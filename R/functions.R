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


# .onLoad <- function(libname, pkgname) {
#   if (!'RDCOMClient' %in% utils::installed.packages()) {
#     remotes::install_github('omegahat/RDCOMClient')
#   }
# }