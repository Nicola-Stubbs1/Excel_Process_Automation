
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

#### Using a pre-loaded excel template  #####

wb<-loadWorkbook("Template.xlsx") # Template uses formulas, conditional formatting and data bars
# Writes whole data set into All data sheet
writeData(wb, "All Data", data_frame, startCol = 1, startRow = 1, colNames = TRUE)

# creates a file name and saves to Outputs folder
output_file_name <- paste0(output_location, current_month,"_final_output_Template.xlsx")
saveWorkbook(wb,output_file_name,overwrite = TRUE)
