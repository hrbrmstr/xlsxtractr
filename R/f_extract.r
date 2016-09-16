#' Extract formulas from an Excel workbook sheet
#'
#' @param xlsx an \code{xlsx} object readin with \code{read_xlsx()}
#' @param n sheet number
#' @return a \code{tibble} with three columns (sheet, cell & formuma text). The
#'     data frame may be empty if no formulas were found. The return will be \code{NULL}
#'         if the sheet wasn't found and a warning to said effect will be issued.
#' @export
#' @examples
#' doc <- read_xlsx(system.file("extdata/wb.xlsx", package="xlsxtractr"))
#'
#' extract_formulas(doc, 1)
#' extract_formulas(doc, 2)
#' extract_formulas(doc, 3) # no formula
#'
#' # gotta extract'em all!
#' map_df(seq_along(doc), ~extract_formulas(doc, .))
extract_formulas <- function(xlsx, n) {
  ensure_xlsx(xlsx)
  if ((n > 0) & (n <= length(xlsx))) {
    sheet <- xlsx$sheets[[n]]
    ns <- xlsx$ns[[n]]
    xml_find_all(sheet, ".//x:row", ns) %>%
      map_df(function(row) {
        xml_find_all(row, ".//x:c", ns) %>%
          map_df(function(col) {
            xml_find_all(col, ".//x:f", ns) %>%
              xml_text() -> f
            if (length(f) > 0) {
              data_frame(sheet=n, cell=xml_attr(col, "r"), f=f)
            } else {
              NULL
            }
          })
      })
  } else {
    warning("Sheet number not found.")
    return(NULL)
  }
}
