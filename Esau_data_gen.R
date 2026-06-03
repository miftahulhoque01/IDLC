library(shiny)
library(shinycssloaders)
library(readxl)
library(dplyr)
library(openxlsx)
library(DT)

options(shiny.maxRequestSize = 100 * 1024^2)

# ── Helpers ───────────────────────────────────────────────────────────────────

last_day_of_month <- function(month_num, year = as.integer(format(Sys.Date(), "%Y"))) {
  if (month_num == 12) {
    last_day <- as.Date(paste(year, 12, 31, sep = "-"))
  } else {
    last_day <- as.Date(paste(year, month_num + 1, 1, sep = "-")) - 1
  }
  format(last_day, "%Y-%m-%d")
}

as_num_safe <- function(x) {
  if (inherits(x, c("POSIXct", "POSIXlt", "Date"))) {
    as.numeric(as.Date(x)) + 25569
  } else {
    suppressWarnings(as.numeric(x))
  }
}

safe_to_date <- function(x) {
  if (inherits(x, c("Date", "POSIXt"))) return(as.Date(x))
  if (is.factor(x)) x <- as.character(x)
  out <- rep(as.Date(NA), length(x))
  for (i in seq_along(x)) {
    val <- x[i]
    if (is.na(val) || trimws(val) == "") next
    num_val <- suppressWarnings(as.numeric(val))
    if (!is.na(num_val)) { out[i] <- as.Date(num_val, origin = "1899-12-30"); next }
    parsed <- tryCatch({
      as.Date(val, tryFormats = c("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y",
                                  "%d/%m/%Y", "%d-%b-%Y", "%d-%b-%y"))
    }, error = function(e) as.Date(NA))
    out[i] <- parsed
  }
  return(out)
}

# ── Core Calculation Pipeline ─────────────────────────────────────────────────

esau_monthly <- function(working, tva, sop, bdm_growth, m) {
  q <- case_when(m %in% 1:3 ~ 1, m %in% 4:6 ~ 2, m %in% 7:9 ~ 3, TRUE ~ 4)
  
  gross_disb <- working |>
    group_by(`Initiating Sales Officer`) |>
    summarise(total_disb_amount = sum(`Disbursement Amount`, na.rm = TRUE), .groups = "drop")
  
  relation <- working |>
    group_by(`Initiating Sales Officer`) |>
    summarise(New = sum(`Relationship with IDLC` == "New", na.rm = TRUE), Total = n(), .groups = "drop")
  
  sop_1 <- sop |>
    filter(`...2` %in% c("TPSP", "Sales", "Supervisor (personal)", "BDM")) |>
    select(4, 34, 40, 41, 61, 64) |>
    mutate(across(-1, as_num_safe))
  
  bdm_growth_1 <- bdm_growth |>
    select(1, base = 3 + (q - 1) * 7, cm = 4 + (q - 1) * 7 + ((m - 1) %% 3) * 2) |>
    mutate(across(c(base, cm), as_num_safe))
  
  esau_df <- tva |>
    filter(`...1` == "Active", `...2` %in% c("TPSP", "Sales", "Supervisor (personal)", "BDM")) |>
    select(1:13, 20+(m-1)*6, 21+(m-1)*6, 25+(m-1)*6) |>
    mutate(across(c(14, 15, 16), as_num_safe)) |>
    left_join(gross_disb, by = c("...4" = "Initiating Sales Officer")) |>
    select(1:15, 17, 16) |>
    left_join(relation, by = c("...4" = "Initiating Sales Officer")) |>
    left_join(sop_1, by = "...4") |>
    mutate(`...61` = if_else(`...2` == "BDM", as_num_safe(`...64`), as_num_safe(`...61`))) |>
    select(-24) |>
    distinct() |>
    left_join(bdm_growth_1, by = c("...4" = "...1")) |>
    mutate(
      `...34` = case_when(`...2` == "BDM" ~ as.numeric(cm), TRUE ~ as_num_safe(`...34`)),
      `Month End Date` = last_day_of_month(m)
    ) |> select(-25)
  
  colnames(esau_df) <- c(
    "Status", "Sales/Supervisor", "SIS/APB", "Name", "Member Code/CIF", "CIF",
    "Joining Date", "Joining SEF Business", "Resigning Date", "2nd Last Promotion",
    "Last Promotion", "Designation", "Branch Name", "Target", "Net Disbursement",
    "Gross Disbursement", "ACH %", "New Accounts", "Total Accounts", "Outstanding Portfolio",
    "X DPD PAR", "NPL", "90 DPD PAR", "Base Target (BDM)", "Month End Date"
  )
  return(esau_df)
}

# ── Excel Generator ───────────────────────────────────────────────────────────

create_excel_output <- function(data, month_num, year_num) {
  wb <- createWorkbook()
  month_label <- month.abb[month_num]
  sheet_name  <- paste0(month_label, "-", year_num)
  addWorksheet(wb, sheetName = sheet_name)
  
  make_hdr <- function(fill_color, font_color = "#000000") {
    createStyle(
      fontName = "IDLC", fontSize = 9,
      fontColour     = paste0("#", gsub("^#", "", font_color)),
      fgFill         = paste0("#", gsub("^#", "", fill_color)),
      halign = "CENTER", valign = "CENTER", wrapText = TRUE,
      textDecoration = "bold", border = "TopBottomLeftRight", borderStyle = "thin"
    )
  }
  
  s_yellow       <- make_hdr("#FFFF99", "#000000")
  s_yellow_red   <- make_hdr("#FFFF99", "#FF0000")
  s_amber_red    <- make_hdr("#FFC000", "#FF0000")
  s_blue         <- make_hdr("#4472C4", "#000000")
  s_blue_white   <- make_hdr("#4472C4", "#FFFFFF")
  s_lt_yellow    <- make_hdr("#FFE598", "#000000")
  s_lt_blue_red  <- make_hdr("#B4C6E7", "#FF0000")
  s_peach_red    <- make_hdr("#F7CAAC", "#FF0000")
  s_vlt_yel_red  <- make_hdr("#FFF2CB", "#FF0000")
  
  text_style <- createStyle(fontName = "IDLC", fontSize = 11, halign = "left",   valign = "center", border = "TopBottomLeftRight", borderColour = "#D9D9D9")
  num_style  <- createStyle(fontName = "IDLC", fontSize = 11, halign = "right",  valign = "center", numFmt = "#,##0",     border = "TopBottomLeftRight", borderColour = "#D9D9D9")
  pct_style  <- createStyle(fontName = "IDLC", fontSize = 11, halign = "right",  valign = "center", numFmt = "0.0%",      border = "TopBottomLeftRight", borderColour = "#D9D9D9")
  date_style <- createStyle(fontName = "IDLC", fontSize = 11, halign = "center", valign = "center", numFmt = "dd-mmm-yy", border = "TopBottomLeftRight", borderColour = "#D9D9D9")
  
  row1_headers <- c(
    "Status", "Sales/   Supervisor", "SIS/APB", "Name ", "Member Code/CIF", "CIF",
    "Joining Date", "", "Resigning Date", "2nd Last Promotion Date", "Last Promotion Date",
    "Designation", "Branch Name",
    "Disbursement Excluding RSTL", "", "", "",
    "Account (Excluding Duplicate and RSTL)", "",
    "Initiating/ Monitoring Portfolio", "", "", "",
    "Base Target (BDM)", "Month End Date"
  )
  row2_headers <- c(
    "", "", "", "", "", "",
    "IDLC", "SEF Business",
    "", "", "", "", "",
    "Target", "Net Disbursement", "Gross Disbursement", "ACH %",
    "New", "Total",
    "Outstanding", "X DPD PAR", "NPL", "90 DPD PAR",
    "", ""
  )
  
  for (col in c(1,2,3,4,5,6,9,10,11,12,13,24,25)) mergeCells(wb, sheet_name, cols = col,    rows = 1:2)
  mergeCells(wb, sheet_name, cols = 7:8,   rows = 1)
  mergeCells(wb, sheet_name, cols = 14:17, rows = 1)
  mergeCells(wb, sheet_name, cols = 18:19, rows = 1)
  mergeCells(wb, sheet_name, cols = 20:23, rows = 1)
  
  writeData(wb, sheet_name, t(row1_headers), startCol = 1, startRow = 1, colNames = FALSE)
  writeData(wb, sheet_name, t(row2_headers), startCol = 1, startRow = 2, colNames = FALSE)
  
  addStyle(wb, sheet_name, s_yellow,      rows = 1, cols = c(1,2,3,5,7,9,11,12), gridExpand = TRUE)
  addStyle(wb, sheet_name, s_yellow,      rows = 1, cols = c(4,13),              gridExpand = TRUE)
  addStyle(wb, sheet_name, s_yellow_red,  rows = 1, cols = 6,                    gridExpand = TRUE)
  addStyle(wb, sheet_name, s_amber_red,   rows = 1, cols = 10,                   gridExpand = TRUE)
  addStyle(wb, sheet_name, s_blue,        rows = 1, cols = 14:17,                gridExpand = TRUE)
  addStyle(wb, sheet_name, s_lt_yellow,   rows = 1, cols = 18:19,                gridExpand = TRUE)
  addStyle(wb, sheet_name, s_lt_blue_red, rows = 1, cols = 20:23,                gridExpand = TRUE)
  addStyle(wb, sheet_name, s_peach_red,   rows = 1, cols = 24,                   gridExpand = TRUE)
  addStyle(wb, sheet_name, s_vlt_yel_red, rows = 1, cols = 25,                   gridExpand = TRUE)
  
  addStyle(wb, sheet_name, s_yellow,      rows = 2, cols = c(1:6,9:13),          gridExpand = TRUE)
  addStyle(wb, sheet_name, s_yellow,      rows = 2, cols = 7:8,                  gridExpand = TRUE)
  addStyle(wb, sheet_name, s_blue,        rows = 2, cols = 14:17,                gridExpand = TRUE)
  addStyle(wb, sheet_name, s_lt_yellow,   rows = 2, cols = 18:19,                gridExpand = TRUE)
  addStyle(wb, sheet_name, s_blue_white,  rows = 2, cols = 20:23,                gridExpand = TRUE)
  addStyle(wb, sheet_name, s_peach_red,   rows = 2, cols = 24,                   gridExpand = TRUE)
  addStyle(wb, sheet_name, s_vlt_yel_red, rows = 2, cols = 25,                   gridExpand = TRUE)
  
  setRowHeights(wb, sheet_name, rows = 1, heights = 32.55)
  setRowHeights(wb, sheet_name, rows = 2, heights = 25.05)
  setColWidths(wb, sheet_name, cols = 1:25,
               widths = c(13,14.78,13,25.44,9.44,13,11.22,13,9.22,9.44,
                          13,15,17.22,12.78,13.22,14,9.44,13,8.78,15.22,
                          12.78,9.44,12.44,14.22,9.44))
  
  if (nrow(data) > 0) {
    clean_df <- data
    for (dc in c(7,8,9,10,11,25)) clean_df[[dc]] <- safe_to_date(clean_df[[dc]])
    writeData(wb, sheet_name, clean_df, startCol = 1, startRow = 3, colNames = FALSE)
    data_rows <- 3:(2 + nrow(data))
    for (c in c(1,2,3,4,5,6,12,13))   addStyle(wb, sheet_name, text_style,  rows = data_rows, cols = c, gridExpand = TRUE)
    for (c in c(7,8,9,10,11,25))       addStyle(wb, sheet_name, date_style,  rows = data_rows, cols = c, gridExpand = TRUE)
    for (c in c(14,15,16,18,19,20,24)) addStyle(wb, sheet_name, num_style,   rows = data_rows, cols = c, gridExpand = TRUE)
    for (c in c(17,21,22,23))          addStyle(wb, sheet_name, pct_style,   rows = data_rows, cols = c, gridExpand = TRUE)
  }
  
  freezePane(wb, sheet_name, firstActiveRow = 3)
  showGridLines(wb, sheet_name, show = TRUE)
  
  out_path <- tempfile(pattern = paste0("ESAU_", month_label, "_", year_num, "_"), fileext = ".xlsx")
  saveWorkbook(wb, out_path, overwrite = TRUE)
  return(out_path)
}

# ── UI ────────────────────────────────────────────────────────────────────────

ui <- fluidPage(
  tags$head(
    tags$style(HTML("
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
      *, *::before, *::after { box-sizing: border-box; }
      body, .content-wrapper, .right-side {
        background-color: #F8FAFC !important;
        font-family: 'Inter', sans-serif !important;
        color: #0F172A;
      }
      .skin-blue .main-header .logo {
        background-color: #0F172A !important; color: #FFFFFF !important;
        font-weight: 700 !important; font-size: 16px !important;
        letter-spacing: 0.5px; border-bottom: 1px solid #1E293B !important;
      }
      .skin-blue .main-header .navbar { background-color: #FFFFFF !important; box-shadow: 0 1px 2px 0 rgba(0,0,0,0.05) !important; }
      .skin-blue .main-header .sidebar-toggle { color: #64748B !important; }
      .skin-blue .main-header .sidebar-toggle:hover { background-color: #F1F5F9 !important; color: #0F172A !important; }
      .skin-blue .main-sidebar, .sidebar { background-color: #0F172A !important; }
      .sidebar .form-group label, .control-label {
        font-size: 11px !important; font-weight: 700 !important; color: #94A3B8 !important;
        text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 8px !important;
      }
      .selectize-input, .form-control {
        background-color: #1E293B !important; border: 1px solid #334155 !important;
        font-size: 13px !important; color: #F8FAFC !important; border-radius: 6px !important;
        box-shadow: none !important; padding: 10px 12px !important; transition: border-color 0.2s ease;
      }
      .selectize-dropdown { background-color: #1E293B !important; border: 1px solid #334155 !important; color: #F8FAFC !important; border-radius: 6px !important; }
      .selectize-dropdown .active { background-color: #2563EB !important; color: white !important; }
      .sidebar .radio label { color: #CBD5E1 !important; font-size: 13px !important; font-weight: 500 !important; }
      .glass-card {
        background: #FFFFFF; border-radius: 12px; padding: 24px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
        border: 1px solid #E2E8F0; margin-top: 20px;
      }
      .nav-pills { padding-bottom: 5px; }
      .nav-pills > li > a {
        color: #64748B !important; font-weight: 600 !important; font-size: 14px;
        border-radius: 8px !important; padding: 10px 24px !important; margin-right: 8px;
        background-color: transparent !important; transition: all 0.2s ease;
      }
      .nav-pills > li > a:hover { background-color: #E2E8F0 !important; color: #0F172A !important; }
      .nav-pills > li.active > a,
      .nav-pills > li.active > a:hover,
      .nav-pills > li.active > a:focus {
        background-color: #2563EB !important; color: #FFFFFF !important;
        box-shadow: 0 4px 14px 0 rgba(37,99,235,0.39) !important;
      }
      body { background: #F8FAFC; color: #0F172A; }
      .navbar-brand-custom {
        background: #0F172A; padding: 20px; color: white;
        border-radius: 0 0 12px 12px; margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
      }
      .navbar-brand-custom h2 { margin: 0; font-weight: 700; letter-spacing: -0.5px; font-size: 20px; }
      .subtitle { opacity: 0.7; font-size: 13px; margin-top: 5px; font-weight: 400; }
      .main-card {
        background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px;
        padding: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); margin-bottom: 20px;
      }
      .section-label {
        font-weight: 700; color: #0F172A; font-size: 11px; text-transform: uppercase;
        letter-spacing: 1px; margin-bottom: 16px; display: block;
        border-left: 3px solid #2563EB; padding-left: 8px;
      }
      .btn-primary {
        background-color: #2563EB !important; border-color: #2563EB !important;
        font-weight: 600; border-radius: 6px; padding: 10px 20px;
        transition: all 0.2s ease; height: 42px; font-family: 'Inter', sans-serif;
      }
      .btn-primary:hover {
        background-color: #1D4ED8 !important; border-color: #1D4ED8 !important;
        transform: translateY(-1px); box-shadow: 0 4px 14px rgba(37,99,235,0.35);
      }
      .btn-success {
        font-weight: 600; border-radius: 6px; padding: 10px 20px;
        background-color: #16A34A !important; border-color: #16A34A !important;
        font-family: 'Inter', sans-serif;
      }
      .btn-success:hover { background-color: #15803D !important; border-color: #15803D !important; }
      .file-input-wrapper {
        background: #F8FAFC; border: 1px dashed #CBD5E1; padding: 12px;
        border-radius: 8px; text-align: center; transition: border-color 0.2s ease;
      }
      .file-input-wrapper:hover { border-color: #2563EB; }
      table.dataTable thead th {
        background-color: #F1F5F9 !important; color: #0F172A !important;
        font-weight: 600 !important; font-size: 12px !important;
        text-transform: uppercase; letter-spacing: 0.5px;
      }
      .dataTables_wrapper .dataTables_scrollBody { border-bottom: none !important; }
      .alert-danger  { border-radius: 8px; font-size: 13px; font-weight: 500; }
      .alert-success { border-radius: 8px; font-size: 13px; font-weight: 500; }
      @keyframes spin          { to { transform: rotate(360deg); } }
      @keyframes progress-fill { from { width: 20%; } to { width: 95%; } }
    "))
  ),
  
  div(class = "navbar-brand-custom",
      div(class = "container-fluid",
          h2("ESAU Monthly Report Dashboard"),
          div("Enterprise pipeline calculation interface and formatting parser.", class = "subtitle")
      )
  ),
  
  div(class = "container-fluid",
      
      div(class = "main-card",
          div(class = "section-label", "Configuration & Data Uploads"),
          fluidRow(
            column(2, selectInput("month_num", "Target Month", choices = setNames(1:12, month.name), selected = 3)),
            column(2, numericInput("year", "Target Year", value = as.integer(format(Sys.Date(), "%Y")), min = 2020, max = 2100)),
            column(2, div(class = "file-input-wrapper", fileInput("file_working", "Working (.xlsx)",          accept = c(".xlsx",".xls"), width = "100%"))),
            column(2, div(class = "file-input-wrapper", fileInput("file_tva",     "Target vs Achiev (.xlsx)", accept = c(".xlsx",".xls"), width = "100%"))),
            column(2, div(class = "file-input-wrapper", fileInput("file_sop",     "SOP (.xlsx)",              accept = c(".xlsx",".xls"), width = "100%"))),
            column(2, div(class = "file-input-wrapper", fileInput("file_bdm",     "BDM Portfolio (.xlsx)",    accept = c(".xlsx",".xls"), width = "100%")))
          ),
          hr(style = "margin-top: 5px; margin-bottom: 20px; border-color: #E2E8F0;"),
          fluidRow(
            column(8, uiOutput("status_box")),
            column(4, actionButton("run_btn", "Compile Data Structure", icon = icon("cogs"), class = "btn-primary", style = "width: 100%; margin-top: 14px;"))
          )
      ),
      
      div(class = "main-card", style = "min-height: 600px;",
          conditionalPanel(
            condition = "output.table_ready",
            div(style = "display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;",
                div(class = "section-label", style = "margin-bottom: 0;", "Interactive Ledger View"),
                div(
                  downloadButton("download_btn", "Export Formatted Excel", class = "btn-success", icon = icon("file-excel")),
                  uiOutput("download_status_ui")
                )
            ),
            div(style = "background: #F8FAFC; padding: 10px; border-radius: 8px; border: 1px solid #E2E8F0;",
                DT::DTOutput("preview_table") |> withSpinner(color = "#2563EB")
            )
          ),
          conditionalPanel(
            condition = "!output.table_ready",
            div(style = "text-align: center; padding-top: 140px;",
                icon("table", style = "font-size: 72px; display: block; margin: auto; color: #CBD5E1;"),
                br(),
                h4("System Awaiting Configuration Data", style = "color: #64748B; font-weight: 600;"),
                p("Please upload all four required .xlsx files and click 'Compile Data Structure'.", style = "color: #94A3B8;")
            )
          )
      )
      
  )
)

# ── Server ────────────────────────────────────────────────────────────────────

server <- function(input, output, session) {
  
  result    <- reactiveVal(NULL)
  out_path  <- reactiveVal(NULL)
  err_msg   <- reactiveVal(NULL)
  dl_status <- reactiveVal(NULL)
  
  observeEvent(input$run_btn, {
    result(NULL); out_path(NULL); err_msg(NULL); dl_status(NULL)
    
    files_needed <- list(
      "Working" = input$file_working, "Target vs Achievement" = input$file_tva,
      "Sales Officer Performance" = input$file_sop, "BDM Portfolio Growth" = input$file_bdm
    )
    missing <- names(Filter(is.null, files_needed))
    if (length(missing)) { err_msg(paste("Missing attachments:", paste(missing, collapse = ", "))); return() }
    
    withProgress(message = "Executing Pipeline Transformations...", value = 0, {
      tryCatch({
        incProgress(0.1, detail = "Parsing operational logs...")
        working    <- read_xlsx(input$file_working$datapath, skip = 1, sheet = 1)
        incProgress(0.2, detail = "Ingesting metric vectors...")
        tva        <- read_xlsx(input$file_tva$datapath,     skip = 3, sheet = 1, col_names = FALSE)
        incProgress(0.2, detail = "Evaluating sales structures...")
        sop        <- read_xlsx(input$file_sop$datapath,     skip = 6, sheet = 1, col_names = FALSE)
        incProgress(0.2, detail = "Mapping division growth charts...")
        bdm_growth <- read_xlsx(input$file_bdm$datapath,     skip = 3, sheet = 1, col_names = FALSE)
        incProgress(0.2, detail = "Running matrix algorithms...")
        df   <- esau_monthly(working, tva, sop, bdm_growth, as.integer(input$month_num))
        incProgress(0.1, detail = "Generating Excel layouts...")
        path <- create_excel_output(df, as.integer(input$month_num), input$year)
        result(df); out_path(path)
      }, error = function(e) {
        err_msg(paste("Parser exception found:", conditionMessage(e)))
      })
    })
  })
  
  observeEvent(result(), { dl_status(NULL) })
  
  output$table_ready <- reactive({ !is.null(result()) })
  outputOptions(output, "table_ready", suspendWhenHidden = FALSE)
  
  output$preview_table <- DT::renderDT({
    req(result())
    preview_data <- result()
    for (dc in c(7,8,9,10,11,25)) preview_data[[dc]] <- format(safe_to_date(preview_data[[dc]]), "%d-%b-%y")
    datatable(
      preview_data,
      options = list(pageLength = 25, scrollX = TRUE, scrollY = "500px",
                     scrollCollapse = TRUE, autoWidth = FALSE, dom = "Bfrtip"),
      class = "cell-border stripe hover", rownames = FALSE
    ) |> formatPercentage(c("ACH %", "X DPD PAR", "NPL", "90 DPD PAR"), 2)
  })
  
  output$status_box <- renderUI({
    if (!is.null(err_msg())) {
      div(class = "alert alert-danger",  style = "margin-top:14px; margin-bottom:0;",
          icon("exclamation-triangle"), " ", err_msg())
    } else if (!is.null(result())) {
      div(class = "alert alert-success", style = "margin-top:14px; margin-bottom:0;",
          icon("check-circle"),
          sprintf(" Success: Ready for generation (%d records processed).", nrow(result())))
    }
  })
  
  output$download_status_ui <- renderUI({
    status <- dl_status()
    if (is.null(status)) return(NULL)
    
    if (status == "preparing") {
      div(
        style = paste(
          "margin-top: 12px; padding: 12px 16px; border-radius: 8px;",
          "background: #EFF6FF; border: 1px solid #BFDBFE;",
          "display: flex; align-items: center; gap: 12px;"
        ),
        tags$span(style = paste(
          "display: inline-block; width: 18px; height: 18px;",
          "border: 3px solid #BFDBFE; border-top-color: #2563EB;",
          "border-radius: 50%; animation: spin 0.8s linear infinite; flex-shrink: 0;"
        )),
        div(
          tags$strong(style = "color: #1D4ED8; font-size: 13px;", "Preparing download..."),
          tags$div(
            style = "margin-top: 4px;",
            div(
              style = "width: 260px; height: 6px; background: #BFDBFE; border-radius: 3px; overflow: hidden;",
              div(style = paste(
                "height: 100%; width: 100%; background: #2563EB; border-radius: 3px;",
                "animation: progress-fill 1.2s ease-in-out infinite alternate;"
              ))
            )
          )
        )
      )
      
    } else if (status == "done") {
      file_name <- paste0("ESAU_", month.abb[as.integer(input$month_num)], "_", input$year, ".xlsx")
      div(
        style = paste(
          "margin-top: 12px; padding: 12px 16px; border-radius: 8px;",
          "background: #F0FDF4; border: 1px solid #BBF7D0;",
          "display: flex; align-items: center; gap: 10px;"
        ),
        icon("check-circle", style = "color: #16A34A; font-size: 20px;"),
        div(
          tags$strong(style = "color: #15803D; font-size: 13px;", "Download complete!"),
          tags$div(
            style = "color: #64748B; font-size: 12px; margin-top: 2px;",
            icon("file-excel", style = "color: #16A34A; margin-right: 4px;"),
            file_name
          )
        )
      )
    }
  })
  
  output$download_btn <- downloadHandler(
    filename = function() {
      paste0("ESAU_", month.abb[as.integer(input$month_num)], "_", input$year, ".xlsx")
    },
    content = function(file) {
      req(out_path())
      dl_status("preparing")
      Sys.sleep(0.8)
      file.copy(out_path(), file)
      dl_status("done")
    },
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  )
}

shinyApp(ui, server)