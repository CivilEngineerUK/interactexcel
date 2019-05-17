#' R_interact_excel 2
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
#' @param wb a spreadsheet
#' @examples \dontrun{
#' input_cells <- c('A1', 'A2')
#' input_values = c(4, 9)
#' output_cells = c('B1', 'B2')
#' # ADD SAMPLE SPREADSHEET TO PACKAGE AND RUN CODE
#' }
#' @export
update_spreadsheet_2 <- 
  function(
    file_location = NULL, 
    sheet = 'Sheet1',
    input_values = NULL,
    input_cells = NULL,
    output_cells = NULL,
    save_spreadsheet = 'new_spreadsheet.xls',
    wb = NULL) {
    
    if (is.null(input_cells)) {
      stop(message('input_cells is empty'))
    }
    if (is.null(input_values)) {
      stop(message('input_values is empty'))
    }
    if (is.null(output_cells)) {
      stop(message('output_cells is empty'))
    }
    
    # link to excel sheet in workbook
    sheet <- wb$Worksheets(sheet)
    
    # change the value of the input cells
    for (i in 1:length(input_cells)) {
      alpha <- gsub('[[:digit:]]', '', input_cells[i])
      num <- readr::parse_number(input_cells[i])
      cell <- sheet$cells(num, which(LETTERS == alpha))
      cell[["Value"]] <- input_values[i]
    }
    
    # read outputs
    output <- c()
    for (i in 1:length(output_cells)) {
      alpha <- gsub('[[:digit:]]', '', output_cells[i])
      num <- readr::parse_number(output_cells[i])
      cell <- sheet$cells(num, which(LETTERS == alpha))
      output[i] <- cell[["Value"]]
    }
    
    if (!is.null(save_spreadsheet)) {
      if (save_spreadsheet == '') {
        wb$Save()
      } else {
        if (file.exists(save_spreadsheet)) {
          file.remove(save_spreadsheet)
        }
        wb$SaveAS(save_spreadsheet)  # save as a new workbook
      }
    }

    # return output
    output <- dplyr::tibble(
      cells = output_cells,
      value = output)
    
    return(output)
  }