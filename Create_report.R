
#### set up script ####

# Import Libraries
library(tidyverse)
library(openxlsx)
library(readxl)

# Import functions
source("src/functions.r")

# Folders - import & output locations (should really be in a .env file to avoid sharing file paths)
import_location <- "Inputs/"
output_location <- "Outputs/"

#### Get current month   ####
current_month <- get_month()

#### Import data  ####
# set up filename - this is the file to import
current_month_filename <- paste0(current_month,"_dummy_data")
# Import csv file 
data_frame <- import_csv(import_location, current_month_filename)

# Create summary- at National Level 
# Summary - National
# Grouping by date and creating a sum of metric
df_national <- data_frame %>%
  group_by (`Date Period`) %>%
  summarise (`National Total` = sum(Metric)) %>%
  mutate (`Organisation code` = '-',
          `Organisation name` = 'National Total') %>%
  select (`Organisation name`,`Organisation code`,`National Total`)

# Create summary - Regional level
# Grouping by region and creating a sum of metric
df_regional <- data_frame %>%
  group_by (`Region Name`,`Region Code`) %>%
  summarise (`Regional Totals` = sum(Metric)) 

# Worksheet Titles
# Titles for the excel output
title <- "Metrics - National & Regional"
sub_title <- "Metrics calculated at both national and regional levels"
contact <- "name@email.com"

#### creates Excel workbook from scratch - including the formatting ####

# Create excel workbook
wb <- createWorkbook()
addWorksheet(wb, "National and Regional", gridLines = FALSE)
setColWidths(wb,"National and Regional", cols = 1:4, widths = c(2, 52, 22, 20))

# Set up font styles
title_font <-  createStyle(fontName = "Arial" , fontSize = 16, fontColour = "#16365C")
header <- createStyle(fontName = "Arial", fontSize = 14, fontColour = "#FFFFFF", fgFill = "#16365C")
blue_font <-  createStyle(fontName = "Arial" , fontSize = 12, fontColour = "#16365C")

# Adding titles
writeData(wb, "National and Regional", title, startCol = 2, startRow = 2)
writeData(wb, "National and Regional", sub_title, startCol = 2, startRow = 3, headerStyle = blue_font)
writeData(wb, "National and Regional", contact, startCol = 2, startRow = 4, headerStyle = blue_font)

# Add data to excel workbook
writeData(wb, "National and Regional", df_national, startCol = 2, startRow = 6, headerStyle = header, borders = 'all', borderStyle = 'thin')
writeData(wb, "National and Regional", df_regional, startCol = 2, startRow = 8, headerStyle = header, borders = 'all', borderStyle = 'thin')

# Add styles to sheet
addStyle(wb, "National and Regional", title_font, rows = 2, cols = 2)
addStyle(wb, "National and Regional", blue_font, rows = c(3,4,7,9:15), cols = c(2:4), gridExpand = TRUE, stack = TRUE)

# Save final excel output
output_file_name <- paste0(output_location, current_month,"_final_output.xlsx")
saveWorkbook(wb,output_file_name,overwrite = TRUE)

