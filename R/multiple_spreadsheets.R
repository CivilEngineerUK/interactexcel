#' Multiple Spreadsheets
#' 
#' Make changes to a spreadsheet and extract the outputs from 
#'   specified cells using a run matrix as an input.
#'   
use_run_matrix <- 
  function(
    run_matrix = NULL, 
    file_location = NULL, 
    sheet = 'Sheet1',
    folder = 'results',
    round_dp = 3) {
    
    library(RDCOMClient)
    
    folder <- check_folder(folder)
    
    if (is.character(run_matrix)) {
      run_matrix <- readxl::read_excel(run_matrix)
    }
  
    # get load case id OR specify own
    if (colnames(run_matrix)[1] == 'LC') {
      LC <- unlist(run_matrix[2:nrow(run_matrix), 1])
      run_matrix <- run_matrix[, 2:length(run_matrix)]
    } else {
      LC <- 1:(nrow(run_matrix) - 1)
    }
    
    # get the input cell ids
    cells <- run_matrix[1, ]
    
    # split into input and output cells
    input_cells <- unlist(cells[, c(!is.na(run_matrix[3, ]))])
    output_cells <- unlist(cells[, c(is.na(run_matrix[3, ]))])
    
    # extract input cell values
    input_values <- run_matrix[2:nrow(run_matrix), names(input_cells)]
      
    # run the code via a for loop
    for (i in 1:(nrow(run_matrix) - 1)) {  
      
      # create name for calculation spreadsheet to be saved to
      calc_name <- 
        stringi::stri_split_fixed(
          file_location,
          '/')[[1]]
      
      calculation_name <- 
        paste0(folder,
          stringi::stri_split_fixed(
            calc_name[length(calc_name)], '.xls')[[1]][1], 
          '_LC', LC[i], '.xls')
      
      # update the output cells in the run_matrix data frame with the
      #   values from the specified output cells in the spreadsheet 
      run_matrix[i + 1, names(output_cells)] <- 
        t(round(update_spreadsheet(file_location, 
                           sheet,
                           readr::parse_number(unlist(input_values[i, ])),
                           input_cells,
                           output_cells,
                           calculation_name)[, 2], round_dp))
    }
    
    # add the load case column
    run_matrix <- data.frame(LC = c(NA, LC), run_matrix) %>%
      as_tibble()

    # return the run_matrix file 
    return(run_matrix)
  }

#' Check Folder
#' 
#' Check whether a folder exists and if so whether it should be
#'   overwritten.
#'   
#' @param folder a string with a folder name. Defaults to \code{NULL}
check_folder <- function(folder = NULL) {
  if (!is.null(folder)) {
    if(dir.exists(folder)) {
      ans <- readline(paste(folder, 'already exists. 
                            Press:\n\n 1 to overwrite \n 2 stop program'))
      if (ans == '1') {
        dir.create(folder)
        folder <- paste0(folder, '/')
      } else {
        stop(paste('Run aborted.', folder, 'not overwritten.'))
      }
    } else {
      dir.create(folder)
      folder <- paste0(folder, '/')
    }
  }
  return(folder)
}