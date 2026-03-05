library(shiny)
library(openxlsx)

# =========================
# Load substances Excel at startup
# =========================
SUBSTANCES_FILE <- "input_substances_questionnaire.xlsx"

# =========================
# Load hazard -> bande + bande -> concentrations
# =========================
BCR_DATA_FILE <- "bcr_data.Rds"
BCR_CONC_FILE <- "bcr_conc.Rds"

if (!file.exists(BCR_DATA_FILE)) {
  stop(sprintf("Fichier introuvable: %s (mettez bcr_data.Rds dans le même dossier que app.R)", BCR_DATA_FILE))
}
if (!file.exists(BCR_CONC_FILE)) {
  stop(sprintf("Fichier introuvable: %s (mettez bcr_conc.Rds dans le même dossier que app.R)", BCR_CONC_FILE))
}

bcr_data <- readRDS(BCR_DATA_FILE)
bcr_data <- subset(bcr_data, Bande_Danger > 1) ## Enleve les bandes cutanees
bcr_conc <- readRDS(BCR_CONC_FILE)

# Sanity checks (fail fast)
req_cols_data <- c("Code", "Libelle_FR_short", "Bande_Danger")
miss_data <- setdiff(req_cols_data, names(bcr_data))
if (length(miss_data) > 0) stop(sprintf("bcr_data.Rds: colonnes manquantes: %s", paste(miss_data, collapse = ", ")))

req_cols_conc <- c("Bande_Danger", "Bande_Conc_ppm", "Bande_Conc_mgm3")
miss_conc <- setdiff(req_cols_conc, names(bcr_conc))
if (length(miss_conc) > 0) stop(sprintf("bcr_conc.Rds: colonnes manquantes: %s", paste(miss_conc, collapse = ", ")))

bcr_data$Code <- toupper(as.character(bcr_data$Code))
allowed_codes <- sort(unique(na.omit(bcr_data$Code)))

# Helper: treat NA/"" as "non renseignés"
nr <- function(x) {
  if (length(x) == 0) return("non renseignés")
  x <- as.character(x)
  if (is.na(x) || trimws(x) == "") "non renseignés" else x
}

`%||%` <- function(x, y) if (is.null(x)) y else x

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
  
  df <- df[, required, drop = FALSE]
  df$NAME_FR <- as.character(df$NAME_FR)
  df <- df[!(is.na(df$NAME_FR) | trimws(df$NAME_FR) == ""), , drop = FALSE]
  if (nrow(df) == 0) stop("Aucune substance valide (NAME_FR vide partout).")
  df
}

substances <- load_substances(SUBSTANCES_FILE)
agents <- substances$NAME_FR

# ---- CONFIG (survey) ----
oel_lists <- c("France", "ACGIH", "MAK", "JSOH", "WEEL", "Autre")

units <- c("mg/m³", "ppm", "\u00B5g/m³", "ng/m³", "mg/L", "f/L", "f/ml", "%")
unit_choices <- c("\u2014 Sélectionner une unité \u2014" = "", setNames(units, units))

# ---- Helpers (hazard parsing + computation) ----
parse_hazard_codes <- function(txt) {
  if (is.null(txt) || !nzchar(txt)) return(character(0))
  
  txt_up <- toupper(txt)
  
  # Capture des mentions même si espaces dans la mention (H 361 f d, EUH 066 x, etc.)
  # IMPORTANT : on ne supprime PAS tous les whitespace du texte entier (sinon on colle les lignes)
  matches <- gregexpr("\\bEUH\\s*\\d{3}\\s*[A-Z]*\\b|\\bH\\s*\\d{3}\\s*[A-Z]*\\b", txt_up, perl = TRUE)
  tokens_full <- unlist(regmatches(txt_up, matches))
  
  if (length(tokens_full) == 0) return(character(0))
  
  # Normalise chaque token: enlever les espaces internes, puis garder seulement le prefixe+3 chiffres
  tokens_full <- gsub("\\s+", "", tokens_full)  # OK car c'est par token
  tokens_clean <- sub("^(EUH\\d{3}).*$", "\\1", tokens_full)
  tokens_clean <- sub("^(H\\d{3}).*$", "\\1", tokens_clean)
  
  unique(tokens_clean)
}

compute_from_codes <- function(codes) {
  codes <- unique(toupper(as.character(codes)))
  codes <- intersect(codes, allowed_codes)
  
  if (length(codes) == 0) {
    return(list(
      codes = character(0),
      bande_danger = NA_integer_,
      conc_ppm = "",
      conc_mgm3 = "",
      details = bcr_data[0, c("Code", "Libelle_FR_short", "Bande_Danger"), drop = FALSE]
    ))
  }
  
  filtered <- bcr_data[bcr_data$Code %in% codes, c("Code", "Libelle_FR_short", "Bande_Danger"), drop = FALSE]
  max_bd <- suppressWarnings(max(as.integer(filtered$Bande_Danger), na.rm = TRUE))
  if (!is.finite(max_bd)) max_bd <- NA_integer_
  
  conc_ppm <- ""
  conc_mgm3 <- ""
  if (!is.na(max_bd)) {
    conc_ppm  <- as.character(bcr_conc$Bande_Conc_ppm[bcr_conc$Bande_Danger == max_bd][1])
    conc_mgm3 <- as.character(bcr_conc$Bande_Conc_mgm3[bcr_conc$Bande_Danger == max_bd][1])
    conc_ppm  <- ifelse(is.na(conc_ppm),  "", conc_ppm)
    conc_mgm3 <- ifelse(is.na(conc_mgm3), "", conc_mgm3)
  }
  
  list(
    codes = codes,
    bande_danger = max_bd,
    conc_ppm = conc_ppm,
    conc_mgm3 = conc_mgm3,
    details = filtered
  )
}

# Fallback when user clicks Apply with no applicable hazard statements
fallback_bd1 <- function() {
  conc_ppm  <- as.character(bcr_conc$Bande_Conc_ppm[bcr_conc$Bande_Danger == 1][1])
  conc_mgm3 <- as.character(bcr_conc$Bande_Conc_mgm3[bcr_conc$Bande_Danger == 1][1])
  conc_ppm  <- ifelse(is.na(conc_ppm),  "", conc_ppm)
  conc_mgm3 <- ifelse(is.na(conc_mgm3), "", conc_mgm3)
  
  list(
    codes = character(0),
    bande_danger = 1L,
    conc_ppm = conc_ppm,
    conc_mgm3 = conc_mgm3,
    details = bcr_data[0, c("Code", "Libelle_FR_short", "Bande_Danger"), drop = FALSE],
    none_clicked = TRUE
  )
}

agent_id <- function(i) paste0("agent_", i)

ui <- fluidPage(
  tags$style(HTML("
/* ===== Fixed floating TOC (robust) ===== */
.toc-fixed{
  position: fixed;
  top: 90px;              /* below title bar */
  right: 18px;
  width: 320px;
  max-height: calc(100vh - 120px);
  overflow-y: auto;
  padding: 12px;
  border: 1px solid #ddd;
  border-radius: 10px;
  background: #fff;
  z-index: 9999;
  box-shadow: 0 2px 10px rgba(0,0,0,0.06);
}
.toc-fixed a{
  display: block;
  padding: 4px 0;
  text-decoration: none;
}
.toc-fixed a:hover{ text-decoration: underline; }

/* Reserve space to avoid overlap with fixed TOC */
.main-with-toc{
  padding-right: 360px;   /* slightly > toc width */
}

/* CheckboxGroup horizontal */
.inline-cbg .checkbox {
  display: inline-block;
  margin-right: 14px;
  margin-bottom: 6px;
}
.inline-cbg label { font-weight: normal; }

/* Button aligned right */
.text-right { text-align: right; }

.required-warn { color: #b10000; font-size: 0.9em; margin-top: 4px; }
.agent-card { padding: 14px; border: 1px solid #ddd; border-radius: 10px; margin-bottom: 14px; }
.label-box { background: #f7f7f7; border-radius: 8px; padding: 10px; margin-top: 8px; }
.mono { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace; }
  ")),
  titlePanel("Enquête de validation — méthodologie d'évaluation des expositions pour des substances n'ayant pas de VLEP françaises"),
  
  tags$div(class = "main-with-toc",
           uiOutput("agents_ui")
  ),
  
  tags$div(
    id = "tocFloat",
    class = "toc-fixed",
    tags$h4("Table des substances"),
    uiOutput("toc_ui"),
    tags$hr(),
    tags$p("Sauvegarder / reprendre une session :"),
    downloadButton("save_session", "Télécharger la session (.rds)"),
    fileInput("load_session", "Charger une session (.rds)", accept = ".rds"),
    tags$hr(),
    tags$p(tags$strong("Export Excel")),
    tags$p("⚠️ Télécharger uniquement à la fin de l'enquête pour transmettre vos réponses."),
    downloadButton("download_xlsx", "télécharger l'enquête complétée")
  )
)

server <- function(input, output, session) {
  
  values <- reactiveValues(data = vector("list", length(agents)))
  haz <- reactiveValues()  # applied hazard computations per agent
  
  output$agents_ui <- renderUI({
    tagList(
      lapply(seq_along(agents), function(i) {
        
        lists_input <- paste0("lists_", i)
        other_input <- paste0("other_", i)
        oel_value_input <- paste0("oel_value_", i)
        unit_input <- paste0("unit_", i)
        
        hcodes_input <- paste0("hcodes_", i)
        hpaste_input <- paste0("hpaste_", i)
        apply_input  <- paste0("apply_h_", i)
        hazard_summary_output <- paste0("hazard_summary_", i)
        
        comment_input <- paste0("comment_", i)
        
        tags$div(
          id = agent_id(i),
          class = "agent-card",
          
          tags$h3(agents[i]),
          tags$p(
            tags$strong("Synonymes : "),
            tags$span(style = "font-size: 1.25em;", as.character(substances$SYN_FR[i]))
          ),
          tags$p(
            tags$strong("INDEX : "),
            tags$span(class = "mono", style = "font-size: 1.25em;", nr(substances$INDEX[i])),
            "  |  ",
            tags$strong("EC : "),
            tags$span(class = "mono", style = "font-size: 1.25em;", nr(substances$EC[i])),
            "  |  ",
            tags$strong("CAS : "),
            tags$span(class = "mono", style = "font-size: 1.25em;", nr(substances$CAS[i]))
          ),
          
          tags$hr(),
          
          tags$h4("Partie 1 — Choix d'une VLEP internationale"),
          tags$div(
            class = "inline-cbg",
            checkboxGroupInput(
              inputId = lists_input,
              label = "Cochez la ou les sources applicables :",
              choices = oel_lists
            )
          ),
          
          conditionalPanel(
            condition = sprintf("input['%s'] && input['%s'].includes('Autre')", lists_input, lists_input),
            textInput(
              inputId = other_input,
              label = "Si « Autre », préciser (obligatoire)",
              value = "",
              placeholder = "Ex : NIOSH, Anses, littérature…"
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
          
          tags$h4("Partie 2 — Choix d'une bande de concentration (Uniquement en l’absence de VLEP internationale)"),
          conditionalPanel(
            condition = sprintf("!input['%s'] || input['%s'].length === 0", lists_input, lists_input),
            
            fluidRow(
              column(
                width = 6,
                selectInput(
                  inputId = hcodes_input,
                  label = "Mention(s) de danger (liste) :",
                  choices = allowed_codes,
                  multiple = TRUE
                )
              ),
              column(
                width = 6,
                textAreaInput(
                  inputId = hpaste_input,
                  label = "Ou collez les codes H3xx / EUHxxx (1 par ligne) :",
                  placeholder = "Exemple:\nH331\nH373\nEUH014",
                  rows = 4,
                  width = "100%"
                )
              )
            ),
            
            tags$div(
              class = "text-right",
              actionButton(
                inputId = apply_input,
                label = "Appliquer les mentions de danger",
                class = "btn-primary"
              )
            ),
            
            uiOutput(hazard_summary_output)
          ),
          
          tags$hr(),
          
          textAreaInput(
            inputId = comment_input,
            label = "Commentaires (optionnel)",
            value = "",
            rows = 2,
            placeholder = "Ajoutez ici toute précision utile…"
          )
        )
      })
    )
  })
  
  # Apply hazards on button click
  lapply(seq_along(agents), function(i) {
    local({
      ii <- i
      hcodes_input <- paste0("hcodes_", ii)
      hpaste_input <- paste0("hpaste_", ii)
      apply_input  <- paste0("apply_h_", ii)
      
      observeEvent(input[[apply_input]], {
        sel <- input[[hcodes_input]]
        if (is.null(sel)) sel <- character(0)
        
        pasted <- parse_hazard_codes(input[[hpaste_input]] %||% "")
        merged <- unique(c(sel, pasted))
        merged <- intersect(toupper(merged), allowed_codes)
        
        out <- compute_from_codes(merged)
        
        # NEW RULE: user clicked Apply but none applicable => show "aucune" and set BD=1
        if (length(out$codes) == 0) {
          out <- fallback_bd1()
        }
        
        haz[[paste0("out_", ii)]] <- out
        
        updateSelectInput(session, hcodes_input, selected = out$codes)
      }, ignoreInit = TRUE)
    })
  })
  
  # Render hazard summaries
  lapply(seq_along(agents), function(i) {
    local({
      ii <- i
      hazard_summary_output <- paste0("hazard_summary_", ii)
      
      output[[hazard_summary_output]] <- renderUI({
        out <- haz[[paste0("out_", ii)]]
        
        # Avant premier clic
        if (is.null(out)) {
          return(tags$div(
            class = "label-box",
            tags$em("Sélectionnez/collez des mentions puis cliquez « Appliquer les mentions de danger » pour calculer les bandes de danger/concentration.")
          ))
        }
        
        # Texte "Mentions prises en compte"
        codes_txt <- if (length(out$codes) == 0) "aucune" else paste(out$codes, collapse = ", ")
        
        # Détails par mention : Code + Bande_Danger
        details_ui <- NULL
        if (!is.null(out$details) && nrow(out$details) > 0) {
          d <- out$details
          d$Code <- toupper(as.character(d$Code))
          d$Bande_Danger <- as.character(d$Bande_Danger)
          
          details_ui <- tags$div(
            tags$p(tags$strong("Détail par mention :")),
            tags$ul(
              lapply(seq_len(nrow(d)), function(r) {
                tags$li(
                  tags$span(class = "mono", paste0(d$Code[r], " → BD ", d$Bande_Danger[r])),
                  if (!is.null(d$Libelle_FR_short) && !is.na(d$Libelle_FR_short[r]) && nzchar(d$Libelle_FR_short[r])) {
                    tags$span(paste0(" — ", d$Libelle_FR_short[r]))
                  }
                )
              })
            )
          )
        } else {
          # Cas "aucune" (fallback BD=1) : pas de détails par code
          details_ui <- tags$div(
            tags$p(tags$strong("Détail par mention :")),
            tags$em("aucune")
          )
        }
        
        tags$div(
          class = "label-box",
          tags$p(tags$strong("Mentions prises en compte : "),
                 tags$span(class = "mono", codes_txt)),
          details_ui,
          tags$p(tags$strong("Bande de danger (max) : "), as.character(out$bande_danger)),
          tags$p(tags$strong("Bande de concentration correspondante (gaz/vapeurs) : "), out$conc_ppm),
          tags$p(tags$strong("Bande de concentration correspondante (aérosols) : "), out$conc_mgm3)
        )
      })
    })
  })
  
  # Save inputs to values$data
  observe({
    lapply(seq_along(agents), function(i) {
      
      out <- haz[[paste0("out_", i)]]
      
      values$data[[i]] <- list(
        lists      = input[[paste0("lists_", i)]],
        other_text = input[[paste0("other_", i)]],
        oel_value  = input[[paste0("oel_value_", i)]],
        unit       = input[[paste0("unit_", i)]],
        
        hcodes       = out$codes %||% character(0),
        bande_danger = out$bande_danger %||% NA_integer_,
        conc_ppm     = out$conc_ppm %||% "",
        conc_mgm3    = out$conc_mgm3 %||% "",
        
        comment    = input[[paste0("comment_", i)]],
        
        hpaste_raw = input[[paste0("hpaste_", i)]] %||% "",
        hsel_raw   = input[[paste0("hcodes_", i)]] %||% character(0)
      )
    })
  })
  
  is_done <- function(v, out_applied = NULL) {
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
    
    out <- out_applied
    if (is.null(out)) return(FALSE)
    bd <- out$bande_danger %||% NA_integer_
    # With fallback, BD=1 even when no hazard codes -> accept as done after Apply
    return(!is.na(bd) && is.finite(as.numeric(bd)))
  }
  
  output$toc_ui <- renderUI({
    tagList(
      lapply(seq_along(agents), function(i) {
        v <- values$data[[i]]
        out <- haz[[paste0("out_", i)]]
        icon <- if (is_done(v, out)) "✔" else "⬜"
        tags$a(href = paste0("#", agent_id(i)), paste(icon, agents[i]))
      })
    )
  })
  
  output$save_session <- downloadHandler(
    filename = function() paste0("enquete_VLEP_", format(Sys.Date(), "%Y-%m-%d"), ".rds"),
    content = function(file) saveRDS(values$data, file)
  )
  
  observeEvent(input$load_session, {
    req(input$load_session$datapath)
    restored <- readRDS(input$load_session$datapath)
    if (!is.list(restored)) return()
    
    restored2 <- vector("list", length(agents))
    n <- min(length(restored), length(restored2))
    restored2[seq_len(n)] <- restored[seq_len(n)]
    values$data <- restored2
    
    lapply(seq_along(agents), function(i) {
      v <- values$data[[i]]
      if (is.null(v)) return()
      
      updateCheckboxGroupInput(session, paste0("lists_", i), selected = v$lists %||% character(0))
      updateTextInput(session, paste0("other_", i), value = v$other_text %||% "")
      updateNumericInput(session, paste0("oel_value_", i), value = v$oel_value %||% NA)
      
      restored_unit <- v$unit %||% ""
      valid_units <- c("", units)
      if (!restored_unit %in% valid_units) restored_unit <- ""
      updateSelectInput(session, paste0("unit_", i), selected = restored_unit)
      
      updateSelectInput(session, paste0("hcodes_", i), selected = v$hsel_raw %||% (v$hcodes %||% character(0)))
      updateTextAreaInput(session, paste0("hpaste_", i),
                          value = v$hpaste_raw %||% paste(v$hcodes %||% character(0), collapse = "\n"))
      
      # restore computed output
      out0 <- compute_from_codes(v$hcodes %||% character(0))
      if (length(out0$codes) == 0 && (length(v$hcodes %||% character(0)) == 0)) {
        # if saved with no codes, you likely want fallback to BD=1
        out0 <- fallback_bd1()
      }
      haz[[paste0("out_", i)]] <- out0
      
      updateTextAreaInput(session, paste0("comment_", i), value = v$comment %||% "")
    })
  })
  
  build_results <- reactive({
    rows <- lapply(seq_along(agents), function(i) {
      v <- values$data[[i]] %||% list(
        lists=NULL, other_text="", oel_value=NA, unit="",
        hcodes=character(0), bande_danger=NA, conc_ppm="", conc_mgm3="",
        comment=""
      )
      
      lists <- v$lists %||% character(0)
      has_list <- length(lists) > 0
      bin_cols <- as.list(setNames(oel_lists %in% lists, oel_lists))
      
      syn <- as.character(substances$SYN_FR[i])
      syn_out <- if (is.na(syn) || trimws(syn) == "") "" else syn
      
      bd <- if (!has_list) v$bande_danger else NA
      conc_ppm <- if (!has_list) (v$conc_ppm %||% "") else ""
      conc_mgm3 <- if (!has_list) (v$conc_mgm3 %||% "") else ""
      hcodes_str <- if (!has_list) paste(v$hcodes %||% character(0), collapse = ";") else ""
      
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
        
        hazard_codes = hcodes_str,
        bande_danger = as.character(bd),
        conc_gaz_vapeurs_ppm = conc_ppm,
        conc_aerosols_mg_m3 = conc_mgm3,
        
        comments    = v$comment %||% "",
        
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
    })
    
    do.call(rbind, rows)
  })
  
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
