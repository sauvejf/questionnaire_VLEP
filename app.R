# app.R
library(shiny)
library(openxlsx)

# ---- CONFIG ----
agents <- sprintf("Agent chimique %02d", 1:50)  # Remplacez par vos 50 substances
oel_lists <- c("France", "ACGIH", "MAK", "JSOH", "WEEL")
units <- c("mg/m3", "ppm", "µg/m3", "mg/L")

# Libellés affichés pour l'échelle 1-5
scale_labels <- c(
  "1" = "Très faible (ex. < 0,01)",
  "2" = "Faible (ex. 0,01–0,1)",
  "3" = "Modérée (ex. 0,1–1)",
  "4" = "Élevée (ex. 1–10)",
  "5" = "Très élevée (ex. > 10)"
)

agent_id <- function(i) paste0("agent_", i)

# ---- UI ----
ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      /* Floating TOC */
      #toc {
        position: fixed;
        top: 80px;
        right: 20px;
        width: 260px;
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
    "))
  ),
  
  titlePanel("Enquête — Attribution de valeurs limites d’exposition professionnelle (VLEP)"),
  
  # Floating table of contents
  tags$div(
    id = "toc",
    tags$h4("Table des matières"),
    tags$div(
      lapply(seq_along(agents), function(i) {
        tags$a(href = paste0("#", agent_id(i)), agents[i])
      })
    ),
    tags$hr(),
    downloadButton("download_xlsx", "Télécharger le fichier Excel")
  ),
  
  fluidRow(
    column(
      12,
      uiOutput("agents_ui")
    )
  )
)

# ---- SERVER ----
server <- function(input, output, session) {
  
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
              column(
                6,
                numericInput(oel_value_input, "Valeur de la VLEP", value = NA, min = 0)
              ),
              column(
                6,
                selectInput(unit_input, "Unité", choices = units, selected = units[1])
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
              inline = TRUE
            ),
            uiOutput(scale_label_output)
          )
        )
      })
    )
  })
  
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
  
  build_results <- reactive({
    rows <- lapply(seq_along(agents), function(i) {
      lists <- input[[paste0("lists_", i)]]
      lists_str <- if (is.null(lists) || length(lists) == 0) "" else paste(lists, collapse = "; ")
      
      has_list <- !is.null(lists) && length(lists) > 0
      oel_value <- if (has_list) input[[paste0("oel_value_", i)]] else NA
      unit <- if (has_list) input[[paste0("unit_", i)]] else ""
      
      scale <- if (!has_list) input[[paste0("scale_", i)]] else ""
      scale_label <- if (!has_list && !is.null(scale) && scale != "") scale_labels[[as.character(scale)]] else ""
      
      data.frame(
        agent_index = i,
        agent_name = agents[i],
        selected_lists = lists_str,
        oel_value = suppressWarnings(as.numeric(oel_value)),
        oel_unit = unit,
        scale_1_5 = as.character(scale),
        scale_label = scale_label,
        stringsAsFactors = FALSE
      )
    })
    
    do.call(rbind, rows)
  })
  
  output$download_xlsx <- downloadHandler(
    filename = function() {
      paste0("Reponses_enquete_VLEP_", format(Sys.Date(), "%Y-%m-%d"), ".xlsx")
    },
    content = function(file) {
      wb <- createWorkbook()
      addWorksheet(wb, "reponses")
      writeData(wb, "reponses", build_results())
      freezePane(wb, "reponses", firstRow = TRUE)
      setColWidths(wb, "reponses", cols = 1:ncol(build_results()), widths = "auto")
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)
