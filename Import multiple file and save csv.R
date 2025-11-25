

# Creates a file path to Downloads folder
filepath <- paste0("C:/Users/",Sys.getenv("USERNAME"),"/Downloads")
setwd(filepath) # Sets the working directory to download folder

if (file.exists(filepath) == TRUE){
  # Lists all files starting with the text pattern
  list <- list.files(filepath, pattern = "text")
  
  # Imports all files in the list
  data_list <- lapply(list, readr::read_csv, stringsAsFactors = FALSE)
  
  # Combines data
  combined_df <- do.call(rbind, data_list)
  
  filename <- paste0("file-",Sys.Date().".csv")
  write.csv(combined_df, "filename.csv", row.names = FALSE)
  
} else {
  Print("Filepath doesn't exist - Check the filepath")
  
}


