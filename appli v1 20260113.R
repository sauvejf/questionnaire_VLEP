library(shiny)
library(openxlsx)

# ---- CONFIG ----
agents <- sprintf("Agent chimique %02d", 1:50)  # Remplacez par vos 50 substances

oel_lists <- c("France", "ACGIH", "MAK", "JSOH", "WEEL")

# Units with explicit "no selection" placeholder
units <- c("mg/m3", "ppm", "\u00B5g/m3", "mg/L")
unit_choices <- c("\u2014 S\u00E9lectionner une unit\u00E9 \u2014" = "", setNames(units, units))

# Libellés affichés pour l'échelle 1-5
scale_labels <- c(
  "1" = "Très faible (ex. < 0,01)",
  "2" = "Faible (ex. 0,01–0,1)",
  "3" = "Modérée (ex. 0,1–1)",
  "4" = "Élevée (ex. 1–10)",
  "5" = "Très élevée (ex. > 10)"
)

agent_id <- function(i) paste0("agent_", i)
`%||%` <- function(x, y) if (is.null(x)) y else x

# ---- UI ----
ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      /* Floating TOC */
      #toc {
        position: fixed;
        top: 80px;
        right: 20px;
        width: 320px;
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
      .toc-tools { margin-top: 8px; }
    "))
  ),
  
  titlePanel("Enquête — Attribution de valeurs limites d’exposition professionnelle (VLEP)"),
  
  # Floating table of contents
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
  
  fluidRow(
    column(12, uiOutput("agents_ui"))
  )
)

# ---- SERVER ----
server <- function(input, output, session) {
  
  # Store responses for save/load
  values <- reactiveValues(data = vector("list", length(agents)))
  
  # ---- Agents UI ----
  output$agents_ui <- renderUI({
    tagList(
      lapply(seq_along(agents), function(i) {
        
        lists_input <- paste0("lists_", i)
        oel_value_input <- paste0("oel_value_", i)
        unit_input <- paste0("unit_", i)
        scale_input <- paste0("scale_", i)
        scale_label_output <- paste0("scale_label_", i)
        comment_input <- paste0("comment_", i)
        
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
              column(
                6,
                numericInput(oel_value_input, "Valeur de la VLEP", value = NA, min = 0)
              ),
              column(
                6,
                # No default unit; explicit placeholder label
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
              label = "Sélectionnez une plage de concentrations (1–5)",
              choices = 1:5,
              selected = character(0),  # no preselect
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
  
  # ---- Persist current state into values$data continuously ----
  observe({
    lapply(seq_along(agents), function(i) {
      values$data[[i]] <- list(
        lists     = input[[paste0("lists_", i)]],
        oel_value = input[[paste0("oel_value_", i)]],
        unit      = input[[paste0("unit_", i)]],     # will be "" until selected
        scale     = input[[paste0("scale_", i)]],
        comment   = input[[paste0("comment_", i)]]
      )
    })
  })
  
  # ---- Strict completion logic for TOC ----
  is_done <- function(v) {
    if (is.null(v)) return(FALSE)
    
    lists <- v$lists %||% character(0)
    has_list <- length(lists) > 0
    
    # Partie 1: requires list(s) + numeric oel_value + unit selected (non-empty)
    if (has_list) {
      val_num <- suppressWarnings(as.numeric(v$oel_value))
      val_ok <- !is.null(v$oel_value) && !is.na(val_num) && is.finite(val_num)
      unit_ok <- !is.null(v$unit) && nzchar(trimws(v$unit))
      return(val_ok && unit_ok)
    }
    
    # Partie 2: requires scale selection
    return(!is.null(v$scale) && nzchar(trimws(as.character(v$scale))))
  }
  
  # ---- Floating TOC with checkmarks ----
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
  
  # ---- Load session (.rds) ----
  observeEvent(input$load_session, {
    req(input$load_session$datapath)
    restored <- readRDS(input$load_session$datapath)
    if (!is.list(restored)) return()
    
    values$data <- restored
    
    lapply(seq_along(agents), function(i) {
      v <- restored[[i]]
      if (is.null(v)) return()
      
      updateCheckboxGroupInput(session, paste0("lists_", i), selected = v$lists %||% character(0))
      updateNumericInput(session, paste0("oel_value_", i), value = v$oel_value %||% NA)
      
      # keep blank if missing / not selected
      restored_unit <- v$unit %||% ""
      if (!restored_unit %in% c("", units)) restored_unit <- ""
      updateSelectInput(session, paste0("unit_", i), selected = restored_unit)
      
      updateRadioButtons(session, paste0("scale_", i), selected = v$scale %||% character(0))
      updateTextAreaInput(session, paste0("comment_", i), value = v$comment %||% "")
    })
  })
  
  # ---- Build results with binary OEL columns ----
  build_results <- reactive({
    rows <- lapply(seq_along(agents), function(i) {
      v <- values$data[[i]] %||% list(lists=NULL, oel_value=NA, unit="", scale="", comment="")
      
      lists <- v$lists %||% character(0)
      has_list <- length(lists) > 0
      
      # binary columns for each OEL list
      bin_cols <- as.list(setNames(oel_lists %in% lists, oel_lists))
      
      scale <- if (!has_list) (v$scale %||% "") else ""
      scale_label <- if (!has_list && !is.null(scale) && scale != "") scale_labels[[as.character(scale)]] else ""
      
      data.frame(
        agent_index = i,
        agent_name  = agents[i],
        bin_cols,
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
    filename = function() {
      paste0("Reponses_enquete_VLEP_", format(Sys.Date(), "%Y-%m-%d"), ".xlsx")
    },
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
