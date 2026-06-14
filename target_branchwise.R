library(openxlsx)
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)

# =====================================================================
# 1. BUILD TEMPLATE FUNCTION (Single Sheet)
# =====================================================================
build_target_template <- function(target_date = "2026-06-30") {
  
  wb <- createWorkbook()
  sheet_name <- "Weighted Target"
  addWorksheet(wb, sheet_name)
  
  parsed_date     <- as.Date(target_date)
  top_row_date    <- format(parsed_date, "%d-%b-%y")   # Creates "30-Jun-26"
  second_row_date <- format(parsed_date, "%b-%y")      # Creates "Jun-26"
  
  # --- Styles ---
  red_style <- createStyle(fontName = "IDLC", fontSize = 10, fontColour = "#FFFFFF", fgFill = "#C00000",
                           halign = "center", valign = "center", textDecoration = "bold",
                           border = "TopBottomLeftRight", wrapText = TRUE)
  
  yellow_style <- createStyle(fontName = "IDLC", fontSize = 10, fontColour = "#C00000", fgFill = "#FFFF00",
                              halign = "center", valign = "center", textDecoration = "bold",
                              border = "TopBottomLeftRight", numFmt = "#,##0.00")
  
  # --- Row 1: Totals and Date (Using Formulas) ---
  writeFormula(wb, sheet_name, x = "SUM(N4:N100000)", startRow = 1, startCol = 14)
  writeFormula(wb, sheet_name, x = "SUM(O4:O100000)", startRow = 1, startCol = 15)
  writeData(wb, sheet_name, top_row_date, startRow = 1, startCol = 16)
  addStyle(wb, sheet_name, yellow_style, rows = 1, cols = 14:16, gridExpand = TRUE)
  
  # --- Row 2 & 3: Basic Info Columns (Cols 1-13) ---
  basic_info <- list(
    list(col = 1,  label = "Status"),               list(col = 2,  label = "Sales/   Supervisor"),
    list(col = 3,  label = "SIS/APB"),              list(col = 4,  label = "Name "),
    list(col = 5,  label = "Member Code/ CIF"),     list(col = 6,  label = "CIF"),
    list(col = 9,  label = "Resigning Date"),       list(col = 10, label = "2nd Last Promotion Date"),
    list(col = 11, label = "Last Promotion Date"),  list(col = 12, label = "Designation"),
    list(col = 13, label = "Branch")
  )
  
  for (item in basic_info) {
    writeData(wb, sheet_name, item$label, startRow = 2, startCol = item$col)
    mergeCells(wb, sheet_name, cols = item$col, rows = 2:3)
    addStyle(wb, sheet_name, red_style, rows = 2:3, cols = item$col, gridExpand = TRUE)
  }
  
  # Joining Date block (Cols 7 & 8)
  writeData(wb, sheet_name, "Joining Date", startRow = 2, startCol = 7)
  mergeCells(wb, sheet_name, cols = 7:8, rows = 2)
  addStyle(wb, sheet_name, red_style, rows = 2, cols = 7:8, gridExpand = TRUE)
  
  writeData(wb, sheet_name, "IDLC", startRow = 3, startCol = 7)
  writeData(wb, sheet_name, "SEF Business", startRow = 3, startCol = 8)
  addStyle(wb, sheet_name, red_style, rows = 3, cols = 7:8, gridExpand = TRUE)
  
  # --- Row 2 & 3: Metric Columns (Cols 14-17) ---
  writeData(wb, sheet_name, second_row_date, startRow = 2, startCol = 14) 
  mergeCells(wb, sheet_name, cols = 14:16, rows = 2)
  addStyle(wb, sheet_name, red_style, rows = 2, cols = 14:16, gridExpand = TRUE)
  
  # Sub-headers
  writeData(wb, sheet_name, "Designation Target", startRow = 3, startCol = 14)
  writeData(wb, sheet_name, "Weighted Target", startRow = 3, startCol = 15)
  writeData(wb, sheet_name, "STM", startRow = 3, startCol = 16)
  addStyle(wb, sheet_name, red_style, rows = 3, cols = 14:16, gridExpand = TRUE)
  
  # Role Column (Col 17)
  writeData(wb, sheet_name, "Role", startRow = 2, startCol = 17)
  mergeCells(wb, sheet_name, cols = 17, rows = 2:3)
  addStyle(wb, sheet_name, red_style, rows = 2:3, cols = 17, gridExpand = TRUE)
  
  # Formatting Widths & Freezing Panes
  setRowHeights(wb, sheet_name, rows = 1:3, heights = c(16, 27, 20.5))
  setColWidths(wb, sheet_name, cols = c(1:4, 14:17), widths = c(8, 18.8, 9.6, 28.6, 16, 16, 22, 10))
  setColWidths(wb, sheet_name, cols = 5:11, widths = 11)
  groupColumns(wb, sheet_name, cols = 5:11, hidden = TRUE) 
  freezePane(wb, sheet_name, firstActiveRow = 4, firstActiveCol = 14)
  
  return(wb)
}


# =====================================================================
# 2. LOAD DATA
# =====================================================================
target_path <- "C:/Users/Nobel/Desktop/Target vs Achievement/6_Target _Jun-2026.xlsx"

rm_target <- readxl::read_excel(target_path, sheet = "Weighted RM Target", skip = 3, col_names = FALSE) %>%
  filter(`...1` == "Active") %>%
  mutate(across(9:11, as.numeric))

team_target <- readxl::read_excel(target_path, sheet = "Weighted Team Target", skip = 4, col_names = FALSE) %>%
  filter(`...1` == "Active") %>%
  mutate(across(9:11, as.numeric))


# =====================================================================
# 3. SPLIT & WRITE FUNCTION (Stacked)
# =====================================================================

get_split_target <- function(rm_target, team_target, target_date = "2026-06-30", base_dir = "D:/R_Testing_Environment") {
  
  parsed_date <- as.Date(target_date)
  month_name  <- format(parsed_date, "%B")
  year_num    <- format(parsed_date, "%Y")
  month_num   <- as.integer(format(parsed_date, "%m"))
  
  # --- Styles ---
  border_type <- "TopBottomLeftRight"
  comma_style <- createStyle(fontName = "IDLC", fontSize = 10, numFmt = "#,##0", border = border_type)
  date_style  <- createStyle(fontName = "IDLC", fontSize = 10, numFmt = "d-mmm-yy", border = border_type)
  text_style  <- createStyle(fontName = "IDLC", fontSize = 10, numFmt = "@", border = border_type)
  base_style  <- createStyle(fontName = "IDLC", fontSize = 10, border = border_type)
  
  # ONLY retaining the Team (Green) filter
  team_bg     <- createStyle(fgFill = "#C6E0B4", border = border_type)
  
  # Set the uniform format width to stretch exactly to the "Role" header (Column 17)
  report_width <- 17 
  
  # --- Group Data ---
  nested_rm   <- rm_target %>% mutate(Branch_dup = `...13`) %>% group_by(Branch_dup) %>% nest()
  nested_team <- team_target %>% mutate(Branch_dup = `...13`) %>% group_by(Branch_dup) %>% nest()
  
  all_branches <- unique(c(nested_rm$Branch_dup, nested_team$Branch_dup))
  all_branches <- all_branches[!is.na(all_branches)]
  
  for (branch_name in all_branches) {
    clean_name <- str_replace_all(branch_name, "[\\\\/:*?\"<>|]", "_")
    
    Outpath <- paste0(base_dir, "/Branchwise_Targets/", branch_name, "/", month_num, ". ", month_name, " ", year_num, "/")
    if (!dir.exists(Outpath)) dir.create(Outpath, recursive = TRUE)
    
    # Generate template
    wb <- build_target_template(target_date)
    
    # Keep track of which row we are writing to, so we can stack data seamlessly
    current_row <- 4 
    
    # --- WRITE 1: RM Target ---
    rm_data_subset <- nested_rm$data[nested_rm$Branch_dup == branch_name]
    if (length(rm_data_subset) > 0) {
      b_data_rm <- rm_data_subset[[1]]
      writeData(wb, sheet = 1, x = b_data_rm, startRow = current_row, colNames = FALSE)
      
      rows_rm <- current_row:(current_row + nrow(b_data_rm) - 1)
      
      # Apply grid lines to the full width (1:17)
      addStyle(wb, sheet = 1, style = base_style, rows = rows_rm, cols = 1:report_width, gridExpand = TRUE, stack = TRUE)
      
      # Apply ONLY Team color formatting to the full width
      col_2_rm <- str_trim(as.character(b_data_rm[[2]]))
      team_idx <- which(col_2_rm == "Team")
      if (length(team_idx) > 0) {
        addStyle(wb, sheet = 1, style = team_bg, rows = (current_row - 1) + team_idx, cols = 1:report_width, gridExpand = TRUE, stack = TRUE)
      }
      
      addStyle(wb, sheet = 1, style = date_style, rows = rows_rm, cols = c(7, 8, 9, 10, 11), gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sheet = 1, style = text_style, rows = rows_rm, cols = c(5, 6), gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sheet = 1, style = comma_style, rows = rows_rm, cols = c(14, 15), gridExpand = TRUE, stack = TRUE)
      
      # Advance the current_row so the next chunk writes directly underneath
      current_row <- current_row + nrow(b_data_rm)
    }
    
    # --- WRITE 2: Team Target (Stacked underneath) ---
    team_data_subset <- nested_team$data[nested_team$Branch_dup == branch_name]
    if (length(team_data_subset) > 0) {
      b_data_team <- team_data_subset[[1]]
      writeData(wb, sheet = 1, x = b_data_team, startRow = current_row, colNames = FALSE)
      
      rows_team <- current_row:(current_row + nrow(b_data_team) - 1)
      
      # Apply grid lines to the full width (1:17) even though data stops at 14
      addStyle(wb, sheet = 1, style = base_style, rows = rows_team, cols = 1:report_width, gridExpand = TRUE, stack = TRUE)
      
      # Apply ONLY Team color formatting to the full width
      col_2_team <- str_trim(as.character(b_data_team[[2]]))
      team_idx <- which(col_2_team == "Team")
      if (length(team_idx) > 0) {
        addStyle(wb, sheet = 1, style = team_bg, rows = (current_row - 1) + team_idx, cols = 1:report_width, gridExpand = TRUE, stack = TRUE)
      }
      
      addStyle(wb, sheet = 1, style = date_style, rows = rows_team, cols = c(7, 8, 9, 10, 11), gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sheet = 1, style = text_style, rows = rows_team, cols = c(5, 6), gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sheet = 1, style = comma_style, rows = rows_team, cols = 14, gridExpand = TRUE, stack = TRUE)
    }
    
    # Save Output
    file_name <- paste0(clean_name, "_Target_", month_name, "-", year_num, ".xlsx")
    saveWorkbook(wb, file.path(Outpath, file_name), overwrite = TRUE)
  }
  
  message("Branch-wise Target reports stacked and generated successfully.")
}

# Run the function:
get_split_target(rm_target = rm_target, team_target = team_target, target_date = "2026-06-30")
