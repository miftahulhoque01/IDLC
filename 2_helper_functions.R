# =====================================================================
# 2_helper_functions.R
# =====================================================================

# Assigns the Head of Retail Banking (HoRB) based on the branch name
branch_horb <- function(BRANCH_NAME) {
  case_when(
    BRANCH_NAME %in% c("Gulshan", "Imamganj", "Keranigonj", "Jessore",
                       "Shatkhira (Jessore)", "Khulna", "Kushtia", "Faridpur",
                       "Mirpur", "Barisal", "Bhola (Barisal)", "Tongi")
    ~ "Ashaqur Rahman",
    
    BRANCH_NAME %in% c("Agrabad", "Chowmuhani", "Feni (Chowmuhani)", "Nandankanon",
                       "Bhulta", "Dilkusha", "Narsingdi", "Narayangonj",
                       "Jatrabari (Dilkusha)")
    ~ "Khandoker Maruf Momin",
    
    BRANCH_NAME %in% c("Elephant Road", "Savar", "Gazipur", "Mymensingh",
                       "Tangail (Gazipur)", "Sherpur (Mymensingh)", "Dinajpur",
                       "Rangpur", "Bogra", "Naogaon (Bogra)", "Natore",
                       "Pabna (Natore)", "Rajshahi", "Hobiganj", "Sylhet",
                       "Comilla", "Brahmanbaria (Comilla)")
    ~ "Mohammad Samiul Alam",
    
    TRUE ~ NA_character_
  )
}

# Dynamically builds the Excel headers and formatting in memory
build_custom_template <- function(report_month_num, report_year) {
  wb <- createWorkbook()
  sheet_name <- "Weight Target vs Achievement"
  addWorksheet(wb, sheet_name)
  
  # Styles
  red_style <- createStyle(fontName = "IDLC", fontSize = 10, fontColour = "#FFFFFF", fgFill = "#C00000",
                           halign = "center", valign = "center", textDecoration = "bold",
                           border = "TopBottomLeftRight", wrapText = TRUE)
  
  month_row1_style <- createStyle(fontName = "IDLC", fontSize = 10, fontColour = "#C00000", fgFill = "#FFFF00",
                                  halign = "center", valign = "center", textDecoration = "bold",
                                  border = "TopBottomLeftRight", wrapText = TRUE)
  
  target_style <- createStyle(fontName = "IDLC", fontSize = 10, fontColour = "#C00000", fgFill = "#FFFF00",
                              halign = "center", valign = "center", textDecoration = "bold",
                              border = "TopBottomLeftRight", wrapText = TRUE)
  
  # Basic info labels
  basic_info <- list(
    list(col = 1,  label = "Status"), list(col = 2,  label = "Sales/   Supervisor"),
    list(col = 3,  label = "SIS/APB"), list(col = 4,  label = "Name "),
    list(col = 5,  label = "Member Code/ CIF"), list(col = 6,  label = "CIF"),
    list(col = 9,  label = "Resigning Date"), list(col = 10, label = "2nd Last Promotion Date"),
    list(col = 11, label = "Last Promotion Date"), list(col = 12, label = "Designation"),
    list(col = 13, label = "Branch")
  )
  
  for (item in basic_info) {
    writeData(wb, sheet_name, item$label, startRow = 2, startCol = item$col)
    mergeCells(wb, sheet_name, cols = item$col, rows = 2:3)
    addStyle(wb, sheet_name, red_style, rows = 2:3, cols = item$col, gridExpand = TRUE)
  }
  
  # Joining Date block
  writeData(wb, sheet_name, "Joining Date", startRow = 2, startCol = 7)
  mergeCells(wb, sheet_name, cols = 7:8, rows = 2)
  addStyle(wb, sheet_name, red_style, rows = 2, cols = 7:8, gridExpand = TRUE)
  writeData(wb, sheet_name, "IDLC", startRow = 3, startCol = 7)
  writeData(wb, sheet_name, "SEF Business", startRow = 3, startCol = 8)
  addStyle(wb, sheet_name, red_style, rows = 3, cols = 7:8, gridExpand = TRUE)
  
  # Total Disbursement block
  writeData(wb, sheet_name, "Total Disbursement", startRow = 2, startCol = 14)
  mergeCells(wb, sheet_name, cols = 14:19, rows = 2)
  addStyle(wb, sheet_name, red_style, rows = 2, cols = 14:19, gridExpand = TRUE)
  
  total_disb_sub <- c("Target", "Total \nAchievement", "Disbursement\nEx. RSTL", "RSTL", "Business\nRef.", "In %")
  for (i in seq_along(total_disb_sub)) {
    col_i <- 13 + i
    writeData(wb, sheet_name, total_disb_sub[i], startRow = 3, startCol = col_i)
    if (i == 1) addStyle(wb, sheet_name, target_style, rows = 3, cols = col_i) else addStyle(wb, sheet_name, red_style, rows = 3, cols = col_i)
  }
  
  # Monthly blocks
  monthly_sub <- c("Target", "Total \nAchievement", "Net Disbursement\nEx. RSTL", "RSTL", "Collection Achievement", "In %")
  start_col <- 20
  
  for (m in 1:report_month_num) {
    month_date <- as.Date(sprintf("%04d-%02d-01", as.integer(report_year), m))
    end_col    <- start_col + 5
    formatted_month <- trimws(format(month_date, "%e-%b-%y"))
    
    writeData(wb, sheet_name, formatted_month, startRow = 1, startCol = start_col)
    addStyle(wb, sheet_name, month_row1_style, rows = 1, cols = start_col)
    writeData(wb, sheet_name, formatted_month, startRow = 2, startCol = start_col)
    mergeCells(wb, sheet_name, cols = start_col:end_col, rows = 2)
    addStyle(wb, sheet_name, red_style, rows = 2, cols = start_col:end_col, gridExpand = TRUE)
    
    for (j in seq_along(monthly_sub)) {
      col_j <- start_col + j - 1
      writeData(wb, sheet_name, monthly_sub[j], startRow = 3, startCol = col_j)
      if (j == 1) addStyle(wb, sheet_name, target_style, rows = 3, cols = col_j) else addStyle(wb, sheet_name, red_style, rows = 3, cols = col_j)
    }
    start_col <- start_col + 6
  }
  
  # HoRB column
  last_col <- 19 + (report_month_num * 6) + 1
  writeData(wb, sheet_name, "HoRB", startRow = 2, startCol = last_col)
  mergeCells(wb, sheet_name, cols = last_col, rows = 2:3)
  addStyle(wb, sheet_name, red_style, rows = 2:3, cols = last_col, gridExpand = TRUE)
  total_cols <- last_col
  
  # Heights & Widths
  setRowHeights(wb, sheet_name, rows = 1:3, heights = c(16.2, 27.0, 20.55))
  setColWidths(wb, sheet_name, cols = 1:4, widths = c(8.0, 18.8, 9.6, 28.6))
  setColWidths(wb, sheet_name, cols = 12:19, widths = c(13.4, 14.2, 15.2, 13.8, 19.4, 10.8, 10.2, 7.4))
  
  for (m in 0:(report_month_num - 1)) {
    base <- 20 + m * 6
    setColWidths(wb, sheet_name, cols = base:(base+3), widths = c(13.6, 15.6, 14.0, 9.2))
  }
  setColWidths(wb, sheet_name, cols = total_cols, widths = 21.8)
  
  # Group & Collapse hidden columns
  setColWidths(wb, sheet_name, cols = 5:11, widths = 11)
  groupColumns(wb, sheet_name, cols = 5:11, hidden = TRUE)
  
  freezePane(wb, sheet_name, firstActiveRow = 4, firstActiveCol = 14)
  
  return(wb)
}