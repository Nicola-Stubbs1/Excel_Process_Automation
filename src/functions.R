
# function to get latest month full name
# This function can be left blank to find current date or a date string can be added in 'yyyy-mm-dd' format
get_month <- function (date = NULL) {
  if (is.null (date)) {
    date = Sys.Date()
  } else {
    date = lubridate::as_date(date)
  }
  
  month = lubridate::month(date,label = TRUE, abbr = FALSE)
  return (month)
}

# Imports a csv file by finding specified file listed in the folder location, 
# then filtering on text then importing the csv file
import_csv <- function(import_location, text) {
  # import_location :  file path      - example '/name/documents/folder'
  # text : text in the file name to find      - example 'month_data_file'

  # Lists all csv files in import location
  all_files <- list.files(import_location, pattern = '.csv')
  
  # finds matches pattern in files in list 
  matching_file <- all_files[grepl(text, all_files)]
  
  # Error message - if number of files found equals 0 and if more than 1 found
  if (length(matching_file) == 0) {
    stop(paste0("There is no csv file containing the file name pattern ", text))
    
  }
  if (length(matching_file) > 1) {
    stop(paste0("There is more than 1 csv file that contains the file name pattern ", text))
    
  }
  
  # Creates a file path using the import path and matching file
  filepath <- paste0(import_location,"//",matching_file)
  
  # csv file is imported using the file path created
  dataframe <- readr::read_csv(filepath)  
  
  return(dataframe)
  
}

# Imports xlsx file by listing files in location, then filtering on xlsx files
# Then finding specified text and importing the xlsx file that matches
import_xlsx <- function(import_location, text, sheet = NULL) {
  # import_location :  file path      - example '/name/documents/folder'
  # text : text in the file name to find      - example 'April_data_file'
  # file_type : type of file you want to import  - example '.csv' or '.xlsx'
  
  # Lists all files in import location
  all_files <- list.files(import_location, pattern = '.xlsx')
  
  # finds matches pattern in files in list 
  matching_file <- all_files[grepl(text, all_files)]
  
  # Error messages if number of files found equals 0 and if more than 1 found
  if (length(matching_file) == 0) {
    stop(paste0("There is no xlsx file containing the file name pattern ", text))
    
  }
  if (length(matching_file) > 1) {
    stop(paste0("There are more than 1 xlsx file that contains the file name pattern ", text))
    
  }
  
  # Creates a file path using the import path and matching file
  filepath <- paste0(import_location,"//",matching_file)
  
  # csv file is imported using the file path created
  dataframe <- readxl::read_excel(filepath, sheet = sheet)  
  
  return(dataframe)
  
}

# Function for SQL Database connection
getdb_connection <- function(server) {
  # SQL Server connection set up  - port automatically set
  
  DBI::dbConnect(odbc::odbc(),
                 
                 Driver = "ODBC Driver 17 for SQL Server", # generic setting
                 
                 Server = server, # uses specified value for the server
                 
                 Database = "Warehouse",
                 
                 Trusted_Connection = "Yes", # Both settings remove need to sign in and trusts connection
                 
                 TrustServerCertificate = "yes")
  
}

get_current_previous_dates <- function (date){
# formats date in ymd format - makes it easier to work with
  formatted_date <- lubridate::ymd(date)
  #Gets month (current)
  currentMonth <- month(formatted_date,label = TRUE,abbr = FALSE)
  #Gets year (current)
  currentyear <- year(formatted_date)
  #gets previous month
  previousMonth <- month(formatted_date - month(1),label = TRUE,abbr = FALSE)
  # gets year for previous month
  previousyear <- year(formatted_date - month(1))
  return (list (current_month = currentMonth,
                current_year = currentyear,
                previous_month = previousMonth,
                previous_month_year = previousyear))
}
