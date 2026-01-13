# app.R
library(shiny)
library(openxlsx)

# Google Sheets
library(googlesheets4)
library(jsonlite)

# ---- CONFIG ----
agents <- sprintf("Agent chimique %02d", 1:50)
oel_lists <- c("France", "ACGIH", "MAK", "JSOH", "WEEL")
units <- c("mg/m3", "ppm", "µg/m3", "mg/L")

scale_labels <- c(
  "1" = "Très faible (ex. < 0,01)",
  "2" = "Faible (ex. 0,01–0,1)",
  "3" = "Modérée (ex. 0,1–1)",
  "4" = "Élevée (ex. 1–10)",
  "5" = "Très élevée (ex. > 10)"
)

agent_id <- function(i) paste0("agent_", i)

# ---- GOOGLE AUTH (via secrets) ----
get_secret <- function(name) {
  val <- Sys.getenv(name, unset = NA)
  if (is.na(val) || !nzchar(val)) stop(sprintf("Missing secret env var: %s", name))
  val
}

gsheet_id <- get_secret("GSHEET_ID")
admin_key <- get_secret("ADMIN_KEY")
sa_json <- get_secret("GCP_SA_JSON")

# Authenticate with service account JSON stored in env var
gs4_auth(token = jsonlite::fromJSON(sa_json))

# Ensure worksheet exists and has header
ensure_sheet_ready <- function() {
  # Create 'responses' sheet if missing
  ss <- googlesheets4::gs4_get(gsheet_id)
  sheets <- ss$sheets$name
  if (!("responses" %in% sheets)) {
    googlesheets4::sheet_add(gsheet_id, "responses")
  }
  
  # Ensure header exists
  # We'll store in long format: one row per (user_key, agent_index)
  header <- data.frame(
    user_key = character(),
    agent_index = integer(),
    agent_name = character(),
    selected_lists = character(),
    oel_value = numeric(),
    oel_unit = character(),
    scale_1_5 = character(),
    scale_label = character(),
    updated_at_utc = character(),
    stringsAsFactors = FALSE
  )
  
  # Read first row; if empty sheet, write header
  existing <- tryCatch(
    suppressWarnings(read_sheet(gsheet_id, sheet = "responses", n_max = 1)),
    error = function(e) NULL
  )
  if (is.null(existing) || nrow(existing) == 0 || !all(names(header) %in% names(existing))) {
    # Overwrite sheet with header
    range_write(gsheet_id, data = header, sheet = "responses", reformat = FALSE)
  }
}

ensure_sheet_ready()

# Upsert (update if exists else append) for one agent row
upsert_response_row <- function(row_df) {
  # row_df must have columns matching header
  stopifnot(all(c("user_key", "agent_index") %in% names(row_df)))
  
  # Read existing for this user+agent
  # For small data (<= a few thousand rows), simplest is to read all and filter.
  all_df <- suppressWarnings(read_sheet(gsheet_id, sheet = "responses"))
  
  key <- row_df$user_key[1]
  idx <- row_df$agent_index[1]
  
  if (nrow(all_df) == 0) {
    sheet_append(gsheet_id, sheet = "responses", data = row_df)
    return(invisible(TRUE))
  }
  
  hit <- which(all_df$user_key == key & as.integer(all_df$agent_index) == as.integer(idx))
  
  if (length(hit) == 0) {
    sheet_append(gsheet_id, sheet = "responses", data = row_df)
  } else {
    # Update the first matching row (should be unique)
    r <- hit[1] + 1  # +1 because sheet row 1 is header
    # Compute A1 range for the whole row across known columns
    # We'll update columns A..I (9 cols)
    # A:user_key B:agent_index C:agent_name D:selected_lists E:oel_value F:oel_unit G:scale_1_5 H:scale_label I:updated_at_utc
    range <- sprintf("A%d:I%d", r, r)
    range_write(gsheet_id, sheet = "responses", data = row_df[, c(
      "user_key","agent_index","agent_name","selected_lists","oel_value","oel_unit","scale_1_5","scale_label","updated_at_utc"
    )], range = range, col_names = FALSE, reformat = FALSE)
  }
  
  invisible(TRUE)
}

# Load all existing rows for a user key
load_user_responses <- function(user_key) {
  df <- suppressWarnings(read_sheet(gsheet_id, sheet = "responses"))
  if (nrow(df) == 0) return(df[0, ])
  df <- df[df$user_key == user_key, , drop = FALSE]
  df
}

# ---- UI ----
ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      #toc {
        position: fixed;
        top: 80px;
        right: 20px;
        width: 300px;
        max-height: 75vh;
        overflow-y: auto;
        background: #ffffff;
        border: 1px solid #e5e5e5;
        border-radius: 10px;
        padding: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
        z-index: 9999;
      }
      #toc h4 { margin-top: 0; font-size: 16px; }
      #toc a { display: block; padding: 4px 0; text-decoration: none; }
      #toc a:hover { text-decoration: underline; }
      @media (max-width: 992px) {
        #toc { position: static; width: 100%; max-height: none; }
      }
      .agent-block {
        padding: 14px;
        border: 1px solid #eee;
        border-radius: 12px;
        margin-bottom: 14px;
        background: #fafafa;
      }
      .agent-title { margin-top: 0; }
      .label-box {
        margin-top: 8px;
        padding: 10px;
        border-radius: 10px;
        background: #fff;
        border: 1px dashed #ddd;
      }
      .key-box {
        padding: 10px;
        border-radius: 10px;
        background: #fff7e6;
        border: 1px solid #ffe0a3;
        margin-bottom: 12px;
      }
    "))
  ),
  
  titlePanel("Enquête — Attribution de valeurs limites d’exposition professionnelle (VLEP)"),
  
  tags$div(
    id = "toc",
    tags$h4("Table des matières"),
    uiOutput("key_status"),
    tags$div(
      lapply(seq_along(agents), function(i) {
        tags$a(href = paste0("#", agent_id(i)), agents[i])
      })
    ),
    tags$hr(),
    actionButton("save_all", "Sauvegarder tout"),
    tags$br(), tags$br(),
    downloadButton("download_xlsx_admin", "Télécharger Excel (admin)")
  ),
  
  fluidRow(
    column(12, uiOutput("agents_ui"))
  )
)

# ---- SERVER ----
server <- function(input, output, session) {
  
  # Extract ?key=... and ?admin=...
  query <- reactive(parseQueryString(session$clientData$url_search))
  user_key <- reactive({
    k <- query()[["key"]]
    if (is.null(k) || !nzchar(k)) "" else k
  })
  is_admin <- reactive({
    a <- query()[["admin"]]
    !is.null(a) && nzchar(a) && identical(a, admin_key)
  })
  
  # Store loaded data for prefilling
  loaded <- reactiveVal(NULL)
  
  # Display key status
  output$key_status <- renderUI({
    if (!nzchar(user_key())) {
      tags$div(class = "key-box",
               tags$b("Accès requis : "),
               "ajoutez votre clé dans l’URL, par ex. ",
               tags$code("?key=AB3F9Q")
      )
    } else {
      tags$div(class = "key-box",
               tags$b("Clé répondant : "),
               tags$code(user_key()),
               if (is_admin()) tags$span(" — mode admin", style="margin-left:8px; color:#7a4; font-weight:600;") else NULL
      )
    }
  })
  
  # Load user responses once key is present
  observeEvent(user_key(), {
    if (!nzchar(user_key())) return()
    df <- load_user_responses(user_key())
    loaded(df)
    
    # Prefill inputs
    if (nrow(df) > 0) {
      for (i in seq_along(agents)) {
        row <- df[df$agent_index == i, , drop = FALSE]
        if (nrow(row) == 0) next
        
        # lists
        lists_str <- row$selected_lists[1]
        lists <- if (!is.na(lists_str) && nzchar(lists_str)) strsplit(lists_str, ";\\s*")[[1]] else character()
        updateCheckboxGroupInput(session, paste0("lists_", i), selected = lists)
        
        # if lists exist, prefill oel/unit
        if (length(lists) > 0) {
          if (!is.na(row$oel_value[1])) updateNumericInput(session, paste0("oel_value_", i), value = as.numeric(row$oel_value[1]))
          if (!is.na(row$oel_unit[1]) && nzchar(row$oel_unit[1])) updateSelectInput(session, paste0("unit_", i), selected = row$oel_unit[1])
        } else {
          # else prefill scale
          if (!is.na(row$scale_1_5[1]) && nzchar(row$scale_1_5[1])) updateRadioButtons(session, paste0("scale_", i), selected = row$scale_1_5[1])
        }
      }
    }
  }, ignoreInit = TRUE)
  
  # Dynamic UI for all agents
  output$agents_ui <- renderUI({
    tagList(
      lapply(seq_along(agents), function(i) {
        
        lists_input <- paste0("lists_", i)
        oel_value_input <- paste0("oel_value_", i)
        unit_input <- paste0("unit_", i)
        scale_input <- paste0("scale_", i)
        scale_label_output <- paste0("scale_label_", i)
        
        tags$div(
          id = agent_id(i),
          class = "agent-block",
          tags$h3(class = "agent-title", paste0(i, ". ", agents[i])),
          
          tags$h4("Partie 1 — Sélection d’une ou plusieurs VLEP internationales"),
          checkboxGroupInput(
            inputId = lists_input,
            label = NULL,
            choices = oel_lists,
            inline = TRUE
          ),
          
          conditionalPanel(
            condition = sprintf("input['%s'] && input['%s'].length > 0", lists_input, lists_input),
            fluidRow(
              column(6, numericInput(oel_value_input, "Valeur de la VLEP", value = NA, min = 0)),
              column(6, selectInput(unit_input, "Unité", choices = units, selected = units[1]))
            )
          ),
          
          tags$hr(),
          
          tags$h4("Partie 2 — Uniquement en l’absence de VLEP internationale"),
          conditionalPanel(
            condition = sprintf("!input['%s'] || input['%s'].length === 0", lists_input, lists_input),
            radioButtons(
              inputId = scale_input,
              label = "Sélectionnez une plage de concentrations (1–5)",
              choices = 1:5,
              inline = TRUE
            ),
            uiOutput(scale_label_output)
          )
        )
      })
    )
  })
  
  # Show label for scale selection
  observe({
    lapply(seq_along(agents), function(i) {
      local({
        ii <- i
        scale_input <- paste0("scale_", ii)
        scale_label_output <- paste0("scale_label_", ii)
        
        output[[scale_label_output]] <- renderUI({
          val <- input[[scale_input]]
          if (is.null(val) || val == "") return(NULL)
          tags$div(
            class = "label-box",
            tags$strong("Libellé sélectionné : "),
            scale_labels[[as.character(val)]]
          )
        })
      })
    })
  })
  
  # Build one row for a given agent i
  build_row_for_agent <- function(i) {
    lists <- input[[paste0("lists_", i)]]
    lists_str <- if (is.null(lists) || length(lists) == 0) "" else paste(lists, collapse = "; ")
    
    has_list <- !is.null(lists) && length(lists) > 0
    oel_value <- if (has_list) input[[paste0("oel_value_", i)]] else NA
    unit <- if (has_list) input[[paste0("unit_", i)]] else ""
    
    scale <- if (!has_list) input[[paste0("scale_", i)]] else ""
    scale_label <- if (!has_list && !is.null(scale) && scale != "") scale_labels[[as.character(scale)]] else ""
    
    data.frame(
      user_key = user_key(),
      agent_index = i,
      agent_name = agents[i],
      selected_lists = lists_str,
      oel_value = suppressWarnings(as.numeric(oel_value)),
      oel_unit = unit,
      scale_1_5 = as.character(scale),
      scale_label = scale_label,
      updated_at_utc = format(Sys.time(), tz = "UTC", usetz = TRUE),
      stringsAsFactors = FALSE
    )
  }
  
  # Auto-save when an input changes (debounced-ish: save per agent on key events)
  # Simple, robust approach: observe changes and upsert that agent
  for (i in seq_along(agents)) {
    local({
      ii <- i
      observeEvent(
        list(
          input[[paste0("lists_", ii)]],
          input[[paste0("oel_value_", ii)]],
          input[[paste0("unit_", ii)]],
          input[[paste0("scale_", ii)]]
        ),
        {
          if (!nzchar(user_key())) return()
          row <- build_row_for_agent(ii)
          upsert_response_row(row)
        },
        ignoreInit = TRUE
      )
    })
  }
  
  # Manual save all
  observeEvent(input$save_all, {
    if (!nzchar(user_key())) {
      showNotification("Clé manquante : ajoutez ?key=... dans l’URL.", type = "error")
      return()
    }
    for (i in seq_along(agents)) {
      upsert_response_row(build_row_for_agent(i))
    }
    showNotification("Sauvegarde terminée.", type = "message")
  })
  
  # Admin download: pulls whole sheet and exports xlsx
  output$download_xlsx_admin <- downloadHandler(
    filename = function() {
      paste0("Reponses_enquete_VLEP_", format(Sys.Date(), "%Y-%m-%d"), ".xlsx")
    },
    content = function(file) {
      if (!is_admin()) {
        # Return a tiny workbook explaining access denied
        wb <- createWorkbook()
        addWorksheet(wb, "access_denied")
        writeData(wb, "access_denied", data.frame(message = "Accès refusé : paramètre admin invalide."))
        saveWorkbook(wb, file, overwrite = TRUE)
        return()
      }
      df <- suppressWarnings(read_sheet(gsheet_id, sheet = "responses"))
      wb <- createWorkbook()
      addWorksheet(wb, "responses_long")
      writeData(wb, "responses_long", df)
      freezePane(wb, "responses_long", firstRow = TRUE)
      setColWidths(wb, "responses_long", cols = 1:ncol(df), widths = "auto")
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)
