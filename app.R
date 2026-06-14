# =====================================================================
# 4_shiny_app.R
# =====================================================================

source("1_libraries.R")
source("2_helper_functions.R")
source("3_main_function.R")

app_css <- HTML("
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

  *, *::before, *::after { box-sizing: border-box; }
  body, .content-wrapper, .right-side { 
    background-color: #F1F5F9 !important; 
    font-family: 'Inter', sans-serif !important; 
    color: #0F172A;
  }
  
  /* OVERRIDE DEFAULT PADDING */
  .content { padding: 20px 30px !important; }

  /* ----------------------------------------- */
  /* NEW: FULL WIDTH TOP CARD STYLING          */
  /* ----------------------------------------- */
  .top-card {
    background-color: #0F172A; 
    background-image: linear-gradient(to right, #0F172A, #1E293B);
    border-radius: 12px; 
    padding: 30px 40px; 
    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
    margin-bottom: 25px;
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    justify-content: center;
    border-left: 4px solid #2563EB; /* Blue accent line */
  }
  .top-card h1 {
    color: #F8FAFC;
    font-weight: 800;
    font-size: 28px;
    letter-spacing: -0.5px;
    margin: 0 0 6px 0;
  }
  .top-card h3 {
    color: #94A3B8;
    font-weight: 600;
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 2px;
    margin: 0;
  }

  /* Glass Card Styling for Main Containers */
  .glass-card {
    background: #FFFFFF; 
    border-radius: 12px; 
    padding: 24px; 
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03); 
    border: 1px solid #E2E8F0; 
    height: 100%;
  }

  /* Input Labels and Fields */
  .form-group label, .control-label { 
    font-size: 12px !important; 
    font-weight: 700 !important; 
    color: #475569 !important; 
    text-transform: uppercase; 
    letter-spacing: 0.5px; 
    margin-bottom: 6px !important; 
  }
  
  .selectize-input, .form-control { 
    background-color: #F8FAFC !important; 
    border: 1px solid #CBD5E1 !important; 
    font-size: 14px !important; 
    color: #0F172A !important; 
    border-radius: 8px !important; 
    box-shadow: none !important; 
    padding: 10px 12px !important; 
    transition: all 0.2s ease;
  }
  .selectize-input:focus, .form-control:focus {
    border-color: #2563EB !important;
    box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1) !important;
  }

  .radio label { color: #334155 !important; font-size: 14px !important; font-weight: 500 !important; }

  /* Primary Button Styling */
  #generate_btn {
    background-color: #2563EB; 
    color: #FFFFFF; 
    font-weight: 600; 
    font-size: 15px;
    border: none; 
    width: 100%; 
    border-radius: 8px; 
    padding: 12px; 
    margin-top: 15px; 
    transition: background 0.3s, transform 0.1s;
    box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
  }
  #generate_btn:hover { background-color: #1D4ED8; }
  #generate_btn:active { transform: translateY(2px); }

  /* Custom Terminal/Console Style for the Log */
  #log_output {
    background-color: #0F172A !important;
    color: #34D399 !important; /* Hacker Green */
    font-family: 'Courier New', Courier, monospace !important;
    font-size: 13px !important;
    border-radius: 8px !important;
    padding: 16px !important;
    border: 1px solid #1E293B !important;
    height: 450px !important;
    overflow-y: auto !important;
    white-space: pre-wrap !important;
    box-shadow: inset 0 2px 4px 0 rgba(0, 0, 0, 0.2);
  }
")

ui <- dashboardPage(
  skin = "blue",
  
  # Completely disable the rigid top header
  dashboardHeader(disable = TRUE),
  
  # Completely disable the sidebar
  dashboardSidebar(disable = TRUE),
  
  dashboardBody(
    tags$head(tags$style(app_css)),
    
    # Custom JS to auto-scroll the log to the bottom when updated
    tags$script(HTML("
      $(document).on('shiny:value', function(event) {
        if (event.name === 'log_output') {
          setTimeout(function() {
            var log = document.getElementById('log_output');
            log.scrollTop = log.scrollHeight;
          }, 50);
        }
      });
    ")),
    
    # --- 1. FULL WIDTH TOP TITLE CARD ---
    fluidRow(
      column(width = 12,
             div(class = "top-card",
                 h1("Branch and HoRB wise File split"),
                 h3("SME-PBM Analytics")
             )
      )
    ),
    
    # --- 2. MAIN TWO-COLUMN LAYOUT ---
    fluidRow(
      
      # LEFT COLUMN: INPUTS (Width 5)
      column(width = 5,
             div(class = "glass-card",
                 h3("Report Configuration", style = "margin-top: 0; color: #0F172A; font-weight: 700; font-size: 20px; border-bottom: 2px solid #F1F5F9; padding-bottom: 10px; margin-bottom: 20px;"),
                 
                 fileInput("file_upload", "1. Upload TVA File (.xlsx)", accept = c(".xlsx")),
                 
                 fluidRow(
                   column(width = 6, selectInput("report_month", "2. Reporting Month", choices = setNames(1:12, month.name), selected = 5)),
                   column(width = 6, numericInput("report_year", "3. Reporting Year", value = 2026, min = 2000, max = 2100))
                 ),
                 
                 textInput("base_dir", "4. Branch Output Directory", value = "D:/R_Testing_Environment"),
                 textInput("zip_dir", "5. HoRB Zip Directory", value = "D:/R_Testing_Environment"),
                 
                 radioButtons("report_type", "6. Select Report Type",
                              choices = c("Branch-wise Only" = "branch",
                                          "HoRB-wise Only" = "horb",
                                          "Generate Both" = "both"),
                              selected = "both", inline = TRUE),
                 
                 actionButton("generate_btn", "Generate Reports", icon = icon("gears"))
             )
      ),
      
      # RIGHT COLUMN: EXECUTION LOG (Width 7)
      column(width = 7,
             div(class = "glass-card",
                 h3("System Status & Logs", style = "margin-top: 0; color: #0F172A; font-weight: 700; font-size: 20px; border-bottom: 2px solid #F1F5F9; padding-bottom: 10px; margin-bottom: 20px;"),
                 verbatimTextOutput("log_output")
             )
      )
    )
  )
)

server <- function(input, output, session) {
  
  # Reactive value to store and append logs
  log_rv <- reactiveVal(paste0(
    "--------------------------------------------------\n",
    "  IDLC MIS REPORTING SYSTEM INITIALIZED\n",
    "  System Time: ", Sys.time(), "\n",
    "--------------------------------------------------\n\n",
    "> Waiting for user input..."
  ))
  
  # Helper function to append new messages to the log smoothly
  append_log <- function(new_msg) {
    current_log <- log_rv()
    log_rv(paste0(current_log, "\n> ", new_msg))
  }
  
  observeEvent(input$generate_btn, {
    req(input$file_upload)
    
    append_log("Initializing generation process...")
    append_log(paste("Target Month:", month.name[as.numeric(input$report_month)], input$report_year))
    
    do_branch <- input$report_type %in% c("branch", "both")
    do_horb   <- input$report_type %in% c("horb", "both")
    
    tryCatch({
      withProgress(message = 'Processing Data', value = 0, {
        
        append_log("Reading and filtering uploaded Excel file...")
        incProgress(0.2, detail = "Loading data...")
        
        tva_raw <- read_excel(input$file_upload$datapath, skip = 3, sheet = "Weight Target vs Achievement", col_names = FALSE) %>% 
          filter(`...1` == "Active") %>% 
          mutate(
            HoRB = branch_horb(`...13`),
            across(9:11, as.numeric)
          )
        
        append_log("Data loaded successfully. Extracting templates and generating regional splits...")
        incProgress(0.5, detail = "Building Excel files...")
        
        # Execute main logic
        result_msgs <- get_split_tva(
          tva = tva_raw, 
          horb_wise = do_horb, 
          branch_wise = do_branch,
          report_month_num = as.numeric(input$report_month), 
          report_year = as.numeric(input$report_year), 
          base_dir = input$base_dir, 
          horb_zip_dir = input$zip_dir
        )
        
        # Append the results returned from 3_main_function.R
        for (msg in result_msgs) {
          append_log(msg)
        }
        
        incProgress(1.0, detail = "Complete!")
        append_log("PROCESS COMPLETE. Ready for next task.")
        
      })
      
      showNotification("Reports successfully generated!", type = "message", duration = 5)
      
    }, error = function(e) {
      append_log(paste("CRITICAL ERROR:", e$message))
      showNotification("Execution failed. Check the log window for details.", type = "error", duration = 8)
    })
  })
  
  output$log_output <- renderText({
    log_rv()
  })
}

shinyApp(ui, server)