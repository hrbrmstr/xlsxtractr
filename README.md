
`xlsxtractr` : Extract Things From Excel (xlsx) Files

Inspired by [this SO question](http://stackoverflow.com/q/39529302/1457051).

The following functions are implemented:

-   `extract_formulas`: Extract formulas from an Excel workbook sheet
-   `read_xlsx`: Read in an Excel document for various extractions

The following data sets are included:

-   `system.file("extdata/wb.xlsx", package="xlsxtractr")`: sample xlsx file

### Installation

``` r
devtools::install_git("https://gitlab.com/hrbrmstr/xlsxtractr.git")

# OR

devtools::install_github("hrbrmstr/xlsxtractr")
```

### Usage

``` r
library(xlsxtractr)
library(purrr)

# current verison
packageVersion("xlsxtractr")
```

    ## [1] '0.1.0'

``` r
doc <- read_xlsx(system.file("extdata/wb.xlsx", package="xlsxtractr"))

class(doc)
```

    ## [1] "xlsx"

``` r
print(doc)
```

    ## /Library/Frameworks/R.framework/Versions/3.3/Resources/library/xlsxtractr/extdata/wb.xlsx:
    ##   A Microsoft Excel xlsx document with 3 sheets.

``` r
length(doc)
```

    ## [1] 3

``` r
extract_formulas(doc, 1)
```

    ## # A tibble: 3 × 3
    ##   sheet  cell          f
    ##   <dbl> <chr>      <chr>
    ## 1     1    A4 SUM(A1:A3)
    ## 2     1    B4 SUM(B1:B3)
    ## 3     1    D4 SUM(A4:B4)

``` r
extract_formulas(doc, 2)
```

    ## # A tibble: 3 × 3
    ##   sheet  cell            f
    ##   <dbl> <chr>        <chr>
    ## 1     2   J11  SUM(J8:J10)
    ## 2     2   K11  SUM(K8:K10)
    ## 3     2   M11 SUM(J11:K11)

``` r
extract_formulas(doc, 3) # no formula
```

    ## # A tibble: 0 × 0

``` r
map_df(seq_along(doc), ~extract_formulas(doc, .))
```

    ## # A tibble: 6 × 3
    ##   sheet  cell            f
    ##   <int> <chr>        <chr>
    ## 1     1    A4   SUM(A1:A3)
    ## 2     1    B4   SUM(B1:B3)
    ## 3     1    D4   SUM(A4:B4)
    ## 4     2   J11  SUM(J8:J10)
    ## 5     2   K11  SUM(K8:K10)
    ## 6     2   M11 SUM(J11:K11)

### Test Results

``` r
library(xlsxtractr)
library(testthat)

date()
```

    ## [1] "Fri Sep 16 09:56:30 2016"

``` r
test_dir("tests/")
```

    ## testthat results ========================================================================================================
    ## OK: 0 SKIPPED: 0 FAILED: 0
    ## 
    ## DONE ===================================================================================================================
