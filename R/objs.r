#' @export
print.xlsx <- function(xlsx) {
  ensure_xlsx(xlsx)
  cat(sprintf("%s:\n  A Microsoft Excel xlsx document with %d sheets.\n", xlsx$path, length(xlsx)))
}

#' @export
length.xlsx <- function(xlsx) {
  ensure_xlsx(xlsx)
  length(xlsx$sheets) %||% 0
}
