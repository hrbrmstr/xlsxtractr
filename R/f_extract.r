#' Read in an Excel document for various extractions
#'
#' Local file path or URL pointing to a \code{.xlsx} file.
#'
#' @param path path to the Excel (xlsx) document
#' @importFrom xml2 read_xml
#' @export
#' @examples
#' doc <- read_xlsx(system.file("extdata/wb.xlsx", package="xlsxtractr"))
#' class(doc)
read_xlsx <- function(path) {

  # make temporary things for us to work with
  tmpd <- tempdir()
  tmpf <- tempfile(tmpdir=tmpd, fileext=".zip")

  on.exit({ #cleanup
    unlink(tmpf)
    unlink(sprintf("%s/docdata", tmpd), recursive=TRUE)
  })

  if (is_url(path)) {
    download.file(path, tmpf)
  } else {
    path <- path.expand(path)
    if (!file.exists(path)) stop(sprintf("Cannot find '%s'", path), call.=FALSE)
    # copy docx to zip (not entirely necessary)
    file.copy(path, tmpf)
  }
  # unzip it
  unzip(tmpf, exdir=sprintf("%s/docdata", tmpd))

  # read the actual XML documents for the worksheets
  sheets <- purrr::map(list.files(sprintf("%s/docdata/xl/worksheets", tmpd),
                                  full.names=TRUE), xml2::read_xml)

  # extract the namespaces
  ns <- purrr::map(sheets, ~xml2::xml_ns_rename(xml2::xml_ns(.), d1="x"))

  # make an object for other functions to work with
  xlsx <- list(sheets=sheets, ns=ns, path=path)

  # special class helps us work with these things
  class(xlsx) <- "xlsx"

  xlsx

}

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
#' extract_formulas(doc, 1)
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


is_url <- function(path) { grepl("^(http|ftp)s?://", path) }

is_xlsx <- function(path) { tolower(file_ext(path)) == "xlsx" }


# used by funtions to make sure they are working with a well-formed docx object
ensure_xlsx <- function(xlsx) {
  if (!inherits(xlsx, "xlsx")) stop("Must pass in an 'xlsx' object read in with read_xlsx()", call.=FALSE)
  invisible(TRUE)
}