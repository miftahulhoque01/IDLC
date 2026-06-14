# =====================================================================
# 3_main_function.R
# =====================================================================

get_split_tva <- function(tva, horb_wise = TRUE, branch_wise = TRUE,
                          report_month_num, report_year, base_dir, horb_zip_dir) {
  
  report_month_name <- month.name[as.integer(report_month_num)]
  
  # Data formatting styles 
  border_type   <- "TopBottomLeftRight"
  comma_style   <- createStyle(fontName = "IDLC", fontSize = 10, numFmt = "#,##0", border = border_type)
  percent_style <- createStyle(fontName = "IDLC", fontSize = 10, numFmt = "0%", border = border_type) 
  date_style    <- createStyle(fontName = "IDLC", fontSize = 10, numFmt = "d-mmm-yy", border = border_type)
  text_style    <- createStyle(fontName = "IDLC", fontSize = 10, numFmt = "@", border = border_type)      
  base_style    <- createStyle(fontName = "IDLC", fontSize = 10, border = border_type)                    
  
  # Row Background Styles 
  team_bg       <- createStyle(fgFill = "#C6E0B4", border = border_type) 
  branch_bg     <- createStyle(fgFill = "#B4C6E7", border = border_type) 
  cluster_bg    <- createStyle(fgFill = "#FCE4D6", border = border_type) 
  sub_branch_bg <- createStyle(fgFill = "#DAC2EC", border = border_type) 
  
  write_branch_report <- function(branch_data, Outpath, file_name) {
    if (!dir.exists(Outpath)) dir.create(Outpath, recursive = TRUE)
    
    wb <- build_custom_template(report_month_num, report_year)
    writeData(wb, sheet = 1, x = branch_data, startRow = 4, colNames = FALSE)
    
    num_rows        <- nrow(branch_data)
    data_rows       <- 4:(4 + num_rows - 1)
    total_data_cols <- ncol(branch_data)
    
    addStyle(wb, sheet = 1, style = base_style, rows = data_rows, cols = 1:total_data_cols, gridExpand = TRUE, stack = TRUE)
    
    col_2_vals <- str_trim(as.character(branch_data[[2]]))
    
    team_rows       <- 3 + which(col_2_vals == "Team")
    branch_rows     <- 3 + which(col_2_vals == "Branch")
    sub_branch_rows <- 3 + which(col_2_vals == "Sub Branch")
    cluster_rows    <- 3 + which(col_2_vals == "Cluster Head")
    
    if (length(team_rows) > 0) addStyle(wb, sheet = 1, style = team_bg, rows = team_rows, cols = 1:total_data_cols, gridExpand = TRUE, stack = TRUE)
    if (length(branch_rows) > 0) addStyle(wb, sheet = 1, style = branch_bg, rows = branch_rows, cols = 1:total_data_cols, gridExpand = TRUE, stack = TRUE)
    if (length(cluster_rows) > 0) addStyle(wb, sheet = 1, style = cluster_bg, rows = cluster_rows, cols = 1:total_data_cols, gridExpand = TRUE, stack = TRUE)
    if (length(sub_branch_rows) > 0) addStyle(wb, sheet = 1, style = sub_branch_bg, rows = sub_branch_rows, cols = 1:total_data_cols, gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb, sheet = 1, style = date_style, rows = data_rows, cols = c(7, 8, 9, 10, 11), gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sheet = 1, style = text_style, rows = data_rows, cols = c(5, 6), gridExpand = TRUE, stack = TRUE)
    
    start_metric_col <- 14
    while (start_metric_col <= total_data_cols) {
      metric_end <- min(start_metric_col + 4, total_data_cols)
      addStyle(wb, sheet = 1, style = comma_style, rows = data_rows, cols = start_metric_col:metric_end, gridExpand = TRUE, stack = TRUE)
      
      percent_col <- start_metric_col + 5
      if (percent_col <= total_data_cols) {
        addStyle(wb, sheet = 1, style = percent_style, rows = data_rows, cols = percent_col, gridExpand = TRUE, stack = TRUE)
      }
      start_metric_col <- start_metric_col + 6
    }
    
    saveWorkbook(wb, file.path(Outpath, file_name), overwrite = TRUE)
  }
  
  log_msgs <- c() # Capture messages for Shiny UI
  
  # --- 1. HoRB-wise Splitting (ZIPPED) ---
  if (horb_wise) {
    if (!dir.exists(horb_zip_dir)) dir.create(horb_zip_dir, recursive = TRUE)
    
    staging_dir <- file.path(tempdir(), paste0("HoRB_Staging_", report_month_num, "_", Sys.Date()))
    if (dir.exists(staging_dir)) unlink(staging_dir, recursive = TRUE) 
    dir.create(staging_dir, recursive = TRUE)
    
    nested_horb <- tva %>% mutate(HoRB_dup = HoRB) %>% group_by(HoRB_dup) %>% nest()
    
    for (i in 1:nrow(nested_horb)) {
      horb_name  <- nested_horb$HoRB_dup[i]
      horb_data  <- nested_horb$data[[i]]
      clean_name <- str_replace_all(horb_name, "[\\\\/:*?\"<>|]", "_")
      
      Outpath   <- file.path(staging_dir, horb_name, paste0(report_month_num, ". ", report_month_name, " ", report_year))
      file_name <- paste0(clean_name, "_Target_vs_Achievement_", report_month_name, "-", report_year, ".xlsx")
      write_branch_report(horb_data, Outpath, file_name)
    }
    
    zip_filename <- paste0("HoRB_Reports_", report_month_name, "-", report_year, ".zip")
    zip_filepath <- file.path(horb_zip_dir, zip_filename)
    
    current_wd <- getwd()
    setwd(staging_dir)
    zip::zip(zipfile = zip_filepath, files = list.files(all.files = TRUE, no.. = TRUE))
    setwd(current_wd)
    unlink(staging_dir, recursive = TRUE)
    
    log_msgs <- c(log_msgs, paste("✅ HoRB Zip created at:", zip_filepath))
  }
  
  # --- 2. Branch-wise Splitting ---
  if (branch_wise) {
    nested_branch <- tva %>% mutate(Branch_dup = `...13`) %>% group_by(Branch_dup) %>% nest()
    
    for (i in 1:nrow(nested_branch)) {
      branch_name <- nested_branch$Branch_dup[i]
      branch_data <- nested_branch$data[[i]]
      clean_name  <- str_replace_all(branch_name, "[\\\\/:*?\"<>|]", "_")
      
      Outpath   <- paste0(base_dir, "/Branchwise_Reports/", branch_name, "/", report_month_num, ". ", report_month_name, " ", report_year, "/")
      file_name <- paste0(clean_name, "_Target_vs_Achievement_", report_month_name, "-", report_year, ".xlsx")
      write_branch_report(branch_data, Outpath, file_name)
    }
    log_msgs <- c(log_msgs, paste("✅ Branch-wise reports saved to:", base_dir))
  }
  
  return(log_msgs)
}