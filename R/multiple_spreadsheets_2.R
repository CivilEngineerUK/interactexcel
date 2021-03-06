#' Run Spreadsheets 2
#' 
#' Make changes to a spreadsheet and extract the outputs from 
#'   specified cells using a run matrix as an input.
#'
#' @param run_matrix a file path to the input matrix OR the matrix itself
#' @param file_location the full path of the calculation spreadsheet
#' @param sheet the sheet in the spreadsheet in which you wish to change values
#' @param folder the folder to save the results to
#' @param round_dp the number of significant figures to round to. Defaults to 3
#' @export
run_spreadsheet_2 <- 
  function(run_matrix = NULL, 
           file_location = NULL, 
           sheet = 'Sheet1',
           folder = 'results',
           round_dp = 3) {
    
    tryCatch({
      return(use_run_matrix_2(run_matrix, 
                              file_location, 
                              sheet,
                              folder,
                              round_dp))
      
    }, warning = function(warning_condition) {
      unlink(folder, recursive = FALSE)
      return(warning('The code has crashed'))
      
    }, error = function(error_condition) {
      unlink(folder, recursive = FALSE)
      return(stop('The code has crashed'))
      
    })
  }

#' Multiple Spreadsheets
#' 
#' Make changes to a spreadsheet and extract the outputs from 
#'   specified cells using a run matrix as an input.
#'
#' @param run_matrix a file path to the input matrix OR the matrix itself
#' @param file_location the full path of the calculation spreadsheet
#' @param sheet the sheet in the spreadsheet in which you wish to change values
#' @param folder the folder to save the results to
#' @param round_dp the number of significant figures to round to. Defaults to 3
#' @export
use_run_matrix_2 <- 
  function(
    run_matrix = NULL, 
    file_location = NULL, 
    sheet = 'Sheet1',
    folder = 'results',
    round_dp = 3) {
    
    check_folder(folder)
    
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
    input_cells <- unlist(cells[, c(!is.na(run_matrix[2, ]))])
    output_cells <- unlist(cells[, c(is.na(run_matrix[2, ]))])
    
    # extract input cell values
    input_values <- run_matrix[2:nrow(run_matrix), names(input_cells)]
    
    
    # open the spreadsheet
    xlApp <- RDCOMClient::COMCreate("Excel.Application")
    wb    <- xlApp[["Workbooks"]]$Open(file_location)
    
    time_start <- Sys.time()
    
    # run the code via a for loop
    for (i in 1:(nrow(run_matrix) - 1)) {  
      
      # not save the results to a folder
      if (is.null(folder)) {
        calculation_name <- NULL
        
      } else {
        
        # create name for calculation spreadsheet to be saved to
        calc_name <- 
          stringi::stri_split_fixed(
            file_location,
            '/')[[1]]
        
        calculation_name <- 
          paste0(#paste0(calc_name[1:(length(calc_name) - 1)], collapse = '/'), '/', 
            folder, '/',
            stringi::stri_split_fixed(
              calc_name[length(calc_name)], '.xls')[[1]][1], 
            '_LC', LC[i], '.xls')
        
      }
      
      # update the output cells in the run_matrix data frame with the
      #   values from the specified output cells in the spreadsheet 
      run_matrix[i + 1, names(output_cells)] <- 
        t(
          round(
            update_spreadsheet_2(
              file_location, 
              sheet,
              readr::parse_number(unlist(input_values[i, ])),
              input_cells,
              output_cells,
              save_spreadsheet = calculation_name,
              wb = wb)[, 2], 
            round_dp)
        )
      
      print(message(paste0('Run ', i, ' of ', (nrow(run_matrix) - 1), 
                           ' complete.\nRun time of ', Sys.time() - time_start)))
    }
    
    # save spreadsheet
    wb$Save()
    
    # close Excel
    xlApp$Quit()  
    
    # add the load case column
    run_matrix <- data.frame(LC = c(NA, LC), run_matrix) 
    
    # return the run_matrix file 
    return(run_matrix)
  }

