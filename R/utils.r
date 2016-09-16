is_url <- function(path) { grepl("^(http|ftp)s?://", path) }

is_xlsx <- function(path) { tolower(file_ext(path)) == "xlsx" }

# used by funtions to make sure they are working with a well-formed docx object
ensure_xlsx <- function(xlsx) {
  if (!inherits(xlsx, "xlsx")) stop("Must pass in an 'xlsx' object read in with read_xlsx()", call.=FALSE)
  invisible(TRUE)
}