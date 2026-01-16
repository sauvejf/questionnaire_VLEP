library(shiny)
library(openxlsx)

# =========================
# Load substances Excel at startup
# =========================

SUBSTANCES_FILE <- "input_substances_questionnaire.xlsx"

# Helper: treat NA/"" as "non renseignés"
nr <- function(x) {
  if (length(x) == 0) return("non renseignés")
  x <- as.character(x)
  if (is.na(x) || trimws(x) == "") "non renseignés" else x
}

# Safe loader
load_substances <- function(path) {
  if (!file.exists(path)) {
    stop(sprintf("Fichier substances introuvable: %s (mettez-le dans le même dossier que app.R)", path))
  }
  
  df <- read.xlsx(path)
  
  required <- c("NAME_FR", "SYN_FR", "INDEX", "EC", "CAS")
  missing <- setdiff(required, names(df))
  if (length(missing) > 0) {
    stop(sprintf(
      "Colonnes manquantes dans %s: %s. Colonnes attendues: %s",
      path, paste(missing, collapse = ", "), paste(required, collapse = ", ")
    ))
  }
  
  # Keep only required columns (plus tolerant to extra columns)
  df <- df[, required, drop = FALSE]
  
  # Drop rows without NAME_FR (cannot show substance without name)
  df$NAME_FR <- as.character(df$NAME_FR)
  df <- df[!(is.na(df$NAME_FR) | trimws(df$NAME_FR) == ""), , drop = FALSE]
  
  if (nrow(df) == 0) stop("Aucune substance valide (NAME_FR vide partout).")
  
  df
}

substances <- load_substances(SUBSTANCES_FILE)

# Dynamic agents list (names displayed)
agents <- substances$NAME_FR

# ---- CONFIG (survey) ----
# Added "Autre"
oel_lists <- c("France", "ACGIH", "MAK", "JSOH", "WEEL", "Autre")

units <- c("mg/m3", "ppm", "\u00B5g/m3", "mg/L")
unit_choices <- c("\u2014 S\u00E9lectionner une unit\u00E9 \u2014" = "", setNames(units, units))

scale_labels <- c(
  "1" = "Très faible (ex. < 0,01)",
  "2" = "Faible (ex. 0,01–0,1)",
  "3" = "Modérée (ex. 0,1–1)",
  "4" = "Élevée (ex. 1–10)",
  "5" = "Très élevée (ex. > 10)"
)

agent_id <- function(i) paste0("agent_", i)
`%||%` <- function(x, y) if (is.null(x)) y else x

# =========================
# UI
# =========================
ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      /* Floating TOC */
      #toc {
        position: fixed;
        top: 80px;
        right: 20px;
        width: 360px;
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
      .agent-title { margin-top: 0; margin-bottom: 6px; }
      .subline { margin: 0 0 6px 0; color: #444; }
      .idline { margin: 0 0 10px 0; color: #555; }
      .label-box {
        margin-top: 8px;
        padding: 10px;
        border-radius: 10px;
        background: #fff;
        border: 1px dashed #ddd;
      }
      .toc-tools { margin-top: 8px; }

      .required-warn {
        color: #b00020;
        margin-top: 4px;
        font-size: 0.85em;
      }
    "))
  ),
  
  titlePanel("Enquête — Attribution de valeurs limites d’exposition professionnelle (VLEP)"),
  
  # Floating TOC
  tags$div(
    id = "toc",
    tags$h4("Table des matières"),
    uiOutput("toc_ui"),
    tags$hr(),
    tags$div(
      class = "toc-tools",
      fileInput("load_session", "Reprendre (.rds)", accept = ".rds"),
      downloadButton("save_session", "Sauvegarder (.rds)"),
      tags$br(),
      downloadButton("download_xlsx", "Télécharger le fichier Excel")
    )
  ),
  
  fluidRow(column(12, uiOutput("agents_ui")))
)

# =========================
# SERVER
# =========================
server <- function(input, output, session) {
  
  # Store responses (supports save/load)
  values <- reactiveValues(data = vector("list", length(agents)))
  
  # ---- Agents UI ----
  output$agents_ui <- renderUI({
    tagList(
      lapply(seq_along(agents), function(i) {
        
        lists_input <- paste0("lists_", i)
        other_input <- paste0("other_", i)  # NEW
        oel_value_input <- paste0("oel_value_", i)
        unit_input <- paste0("unit_", i)
        scale_input <- paste0("scale_", i)
        scale_label_output <- paste0("scale_label_", i)
        comment_input <- paste0("comment_", i)
        
        # substance info
        syn <- as.character(substances$SYN_FR[i])
        syn_has <- !(is.na(syn) || trimws(syn) == "")
        
        idx <- nr(substances$INDEX[i])
        ec  <- nr(substances$EC[i])
        cas <- nr(substances$CAS[i])
        
        tags$div(
          id = agent_id(i),
          class = "agent-block",
          
          # Ligne 1: NAME_FR (also used in TOC)
          tags$h3(class = "agent-title", paste0(i, ". ", agents[i])),
          
          # Ligne 2: SYN_FR only if present, smaller
          if (syn_has) tags$p(class = "subline", tags$small(paste0("Synonymes : ", syn))),
          
          # Ligne 3: IDs with "non renseignés" if blank
          tags$p(
            class = "idline",
            tags$small(paste0("# index : ", idx, "  |  # EC : ", ec, "  |  # CAS : ", cas))
          ),
          
          tags$h4("Partie 1 — Sélection d’une ou plusieurs VLEP internationales"),
          checkboxGroupInput(
            inputId = lists_input,
            label = NULL,
            choices = oel_lists,
            inline = TRUE
          ),
          
          # NEW: if "Autre" selected, require text
          conditionalPanel(
            condition = sprintf("input['%s'] && input['%s'].includes('Autre')", lists_input, lists_input),
            textInput(
              inputId = other_input,
              label = "Si « Autre », préciser (obligatoire)",
              value = "",
              placeholder = "Ex : NIOSH, SUVA, littérature…"
            ),
            tags$div(class = "required-warn", "Champ obligatoire si « Autre » est coché.")
          ),
          
          conditionalPanel(
            condition = sprintf("input['%s'] && input['%s'].length > 0", lists_input, lists_input),
            fluidRow(
              column(6, numericInput(oel_value_input, "Valeur de la VLEP", value = NA, min = 0)),
              column(
                6,
                selectInput(
                  inputId = unit_input,
                  label = "Unité",
                  choices = unit_choices,
                  selected = ""
                )
              )
            )
          ),
          
          tags$hr(),
          
          tags$h4("Partie 2 — Uniquement en l’absence de VLEP internationale"),
          conditionalPanel(
            condition = sprintf("!input['%s'] || input['%s'].length === 0", lists_input, lists_input),
            radioButtons(
              inputId = scale_input,
              label = "Sélectionnez une bande de danger",
              choices = 1:5,
              selected = character(0),
              inline = TRUE
            ),
            uiOutput(scale_label_output)
          ),
          
          tags$hr(),
          
          textAreaInput(
            inputId = comment_input,
            label = "Commentaires (optionnel)",
            value = "",
            rows = 3,
            placeholder = "Ajoutez ici toute précision utile…"
          )
        )
      })
    )
  })
  
  # ---- Scale label display ----
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
  
  # ---- Persist current state ----
  observe({
    lapply(seq_along(agents), function(i) {
      values$data[[i]] <- list(
        lists      = input[[paste0("lists_", i)]],
        other_text = input[[paste0("other_", i)]],   # NEW
        oel_value  = input[[paste0("oel_value_", i)]],
        unit       = input[[paste0("unit_", i)]],    # "" until selected
        scale      = input[[paste0("scale_", i)]],
        comment    = input[[paste0("comment_", i)]]
      )
    })
  })
  
  # ---- Strict completion logic for TOC ----
  is_done <- function(v) {
    if (is.null(v)) return(FALSE)
    
    lists <- v$lists %||% character(0)
    has_list <- length(lists) > 0
    
    if (has_list) {
      val_num <- suppressWarnings(as.numeric(v$oel_value))
      val_ok <- !is.null(v$oel_value) && !is.na(val_num) && is.finite(val_num)
      unit_ok <- !is.null(v$unit) && nzchar(trimws(v$unit))
      
      autre_selected <- "Autre" %in% lists
      autre_ok <- TRUE
      if (autre_selected) {
        autre_ok <- !is.null(v$other_text) && nzchar(trimws(v$other_text))
      }
      
      return(val_ok && unit_ok && autre_ok)
    }
    
    return(!is.null(v$scale) && nzchar(trimws(as.character(v$scale))))
  }
  
  # ---- TOC with checkmarks (uses NAME_FR) ----
  output$toc_ui <- renderUI({
    tagList(
      lapply(seq_along(agents), function(i) {
        v <- values$data[[i]]
        icon <- if (is_done(v)) "✔" else "⬜"
        tags$a(href = paste0("#", agent_id(i)), paste(icon, agents[i]))
      })
    )
  })
  
  # ---- Save session (.rds) ----
  output$save_session <- downloadHandler(
    filename = function() paste0("enquete_VLEP_", format(Sys.Date(), "%Y-%m-%d"), ".rds"),
    content = function(file) saveRDS(values$data, file)
  )
  
  # ---- Load session (.rds), tolerant to changed number of substances ----
  observeEvent(input$load_session, {
    req(input$load_session$datapath)
    restored <- readRDS(input$load_session$datapath)
    if (!is.list(restored)) return()
    
    # If substances list length changed, adapt:
    restored2 <- vector("list", length(agents))
    n <- min(length(restored), length(restored2))
    restored2[seq_len(n)] <- restored[seq_len(n)]
    values$data <- restored2
    
    lapply(seq_along(agents), function(i) {
      v <- values$data[[i]]
      if (is.null(v)) return()
      
      updateCheckboxGroupInput(session, paste0("lists_", i), selected = v$lists %||% character(0))
      updateTextInput(session, paste0("other_", i), value = v$other_text %||% "")  # NEW
      updateNumericInput(session, paste0("oel_value_", i), value = v$oel_value %||% NA)
      
      restored_unit <- v$unit %||% ""
      valid_units <- c("", units)
      if (!restored_unit %in% valid_units) restored_unit <- ""
      updateSelectInput(session, paste0("unit_", i), selected = restored_unit)
      
      updateRadioButtons(session, paste0("scale_", i), selected = v$scale %||% character(0))
      updateTextAreaInput(session, paste0("comment_", i), value = v$comment %||% "")
    })
  })
  
  # ---- Build results (Excel) with binary OEL columns + substance identifiers ----
  build_results <- reactive({
    rows <- lapply(seq_along(agents), function(i) {
      v <- values$data[[i]] %||% list(lists=NULL, other_text="", oel_value=NA, unit="", scale="", comment="")
      
      lists <- v$lists %||% character(0)
      has_list <- length(lists) > 0
      
      # binary columns for each OEL list (incl. Autre)
      bin_cols <- as.list(setNames(oel_lists %in% lists, oel_lists))
      
      scale <- if (!has_list) (v$scale %||% "") else ""
      scale_label <- if (!has_list && !is.null(scale) && scale != "") scale_labels[[as.character(scale)]] else ""
      
      syn <- as.character(substances$SYN_FR[i])
      syn_out <- if (is.na(syn) || trimws(syn) == "") "" else syn
      
      data.frame(
        substance_name_fr = substances$NAME_FR[i],
        synonyms_fr       = syn_out,
        index_id          = nr(substances$INDEX[i]),
        ec_id             = nr(substances$EC[i]),
        cas_id            = nr(substances$CAS[i]),
        
        bin_cols,
        
        autre_details = if ("Autre" %in% lists) (v$other_text %||% "") else "",
        
        oel_value   = suppressWarnings(as.numeric(if (has_list) v$oel_value else NA)),
        oel_unit    = if (has_list) (v$unit %||% "") else "",
        scale_1_5   = as.character(scale),
        scale_label = scale_label,
        comments    = v$comment %||% "",
        
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
    })
    
    do.call(rbind, rows)
  })
  
  # ---- Excel download ----
  output$download_xlsx <- downloadHandler(
    filename = function() paste0("Reponses_enquete_VLEP_", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"),
    content = function(file) {
      df <- build_results()
      wb <- createWorkbook()
      addWorksheet(wb, "reponses")
      writeData(wb, "reponses", df)
      freezePane(wb, "reponses", firstRow = TRUE)
      setColWidths(wb, "reponses", cols = 1:ncol(df), widths = "auto")
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)
