# Import Libraries
library(tidyverse)
library(openxlsx)
library(readxl)

# Import data
data_frame <- read_excel("Inputs/dummy_data.xlsx")

# Summary - National
# Grouping by date and creating a sum of metric
df_national <- data_frame %>%
  group_by (`Date Period`) %>%
  summarise (`National Total` = sum(Metric)) %>%
  mutate (`Organisation code` = '-',
          `Organisation name` = 'National Total') %>%
  select (`Organisation name`,`Organisation code`,`National Total`)

# Summary - Regional
# Grouping by region and creating a sum of metric
df_regional <- data_frame %>%
  group_by (`Region Name`,`Region Code`) %>%
  summarise (`Regional Totals` = sum(Metric)) 

# Worksheet Titles
# Titles for the excel output
title <- "Metrics - National & Regional"
sub_title <- "Metrics calculated at both national and regional levels"
contact <- "name@email.com"

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
saveWorkbook(wb,"Outputs/final_output.xlsx",overwrite = TRUE)
