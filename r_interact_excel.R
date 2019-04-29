#' R_interact_Excel
#' 
#' Allows a spreadsheet to be modified programatically and the 
#'   outputs captured. Uses the \code{RDCOMClient}, \code{dplyr} and
#'   \code{stringr} packages
#'   
#' @param file_location file location of the spreadsheet to modify. The
#'   spreadsheet should be \code{.xls} format
#' @param sheet the sheet name. Defaults to \code{Sheet1}
#' @param input_values a vector of values corresponding to the cells to be modified
#' @param input_cells a character vector of the input cells 
#'   in the form c('A1', 'A2')
#' @param output_cells the cells from which the output will be read
#' @param save_spreadsheet if ' ' then will overwrite the current spreadsheet.
#'   If \code{NULL} will not save. If any other name, will save the file as an
#'   \code{.xls} object
#' @examples \dontrun{
#' input_cells <- c('A1', 'A2')
#' input_values = c(4, 9)
#' output_cells = c('B1', 'B2')
#' }
#' @export
update_spreadsheet <- 
  function(
    file_location = NULL, 
    sheet = 'Sheet1',
    input_values = NULL,
    input_cells = NULL,
    output_cells = NULL,
    save_spreadsheet = 'new_spreadsheet.xls') {
    
    if (is.null(input_cells)) {
      stop(message('input_cells is empty'))
    }
    if (is.null(input_values)) {
      stop(message('input_values is empty'))
    }
    if (is.null(output_cells)) {
      stop(message('output_cells is empty'))
    }
    
    # load spreadsheet
    xlApp <- RDCOMClient::COMCreate("Excel.Application")
    wb    <- xlApp[["Workbooks"]]$Open(file_location)
    sheet <- wb$Worksheets(sheet)
    
    # change the value of the input cells
    for (i in 1:nrow(input)) {
      id <- stringr::str_split(input_cells[i], '')[[1]]
      cell <- sheet$cells( as.integer(id[2]), which(LETTERS == id[1]))
      cell[["Value"]] <- input_values[i]
    }
    
    # read outputs
    output <- c()
    for (i in 1:length(output_cells)) {
      id <- stringr::str_split(output_cells[i], '')[[1]]
      cell <- sheet$cells(as.integer(id[2]), which(LETTERS == id[1]))
      output[i] <- cell[["Value"]]
    }
    
    if (!is.null(save_spreadsheet)) {
      if (save_spreadsheet == ' ') {
        wb$Save()
      } else {
        if (file.exists(save_spreadsheet)) {
          file.remove(save_spreadsheet)
        }
        wb$SaveAS(save_spreadsheet)  # save as a new workbook
      }
    }
    
    xlApp$Quit()  # close Excel
    
    # return output
    output <- dplyr::tibble(
      cells = output_cells,
      value = output)
    return(output)
  }