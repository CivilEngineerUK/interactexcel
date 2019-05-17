#' Check Folder
#' 
#' Check whether a folder exists and overwrite it if so
#'   
#' @param folder a string with a folder name. Defaults to \code{NULL}
check_folder <- function(folder = NULL) {
  if (!is.null(folder)) {
    if(dir.exists(folder)) {
      ans <- readline(paste(folder, 'already exists. 
                            Press:\n\n 1 to overwrite \n 2 stop program'))
      if (ans == '1') {
        unlink(folder, recursive = TRUE)
      } else {
        stop(paste('Run aborted.', folder, 'not overwritten.'))
      }
    } 
    dir.create(folder)
    folder <- paste0(folder, '/')
  }
}

#' Install RDCOMClient
#' 
#' Installs RDCOMClient from Github
#' 
#' @description RDCOMClient is a necessary package for the package to interact with 
#'   Excel spreadsheets.
#' @return thing
#' @export
install_RDCOMClient <- function() {
  if (!'RDCOMClient' %in% utils::installed.packages()) {
    remotes::install_github('omegahat/RDCOMClient')
  }
}

#' Read results
#' 
#' Read the results files that are in .csv format
#' 
#' @param folder_directory directory to search for folder. Defaults to \code{getwd()}
#' @export
read_results <- function(folder_directory = getwd()) {

  # read the file names
  file_names <- list.files(folder_directory)
  
  # read the files into a list
  results <- lapply(file_names, function(x) read.csv(paste0(folder_directory, '/', x)))
  
  # name each memeber in the list
  names(results) <- gsub('.csv', '', file_names)
  
  lapply(results, remove_top_row)
}

remove_top_row <- function(df) {
  df %>%
    slice(-1) %>%
    apply(c(1, 2),
          function(x)
            readr::parse_number(readr::parse_character(x))) %>%
    tibble::as_tibble()
}