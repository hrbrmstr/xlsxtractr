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
