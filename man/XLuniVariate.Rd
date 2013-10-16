\name{XLunivariate}
\alias{XLunivariate}
\title{Univariate Statistics Exported to Excel}
\usage{
  XLunivariate(wb, sheet, calcvar,
    rowvar = rep("All", length(calcvar)),
    fun1 = list(fun = roundmean, name = "Mean"),
    fun2 = list(fun = roundSD, name = "SD"),
    seps = c("", " (", ")"), rowTitle = "",
    rowNames = NULL, ord = NULL, row1 = 1, col1 = 1,
    title = "Summaries", purge = FALSE, ...)
}
\arguments{
  \item{wb}{a \code{\link[XLConnect]{workbook-class}}
  object}

  \item{sheet}{numeric or character: a worksheet name
  (character) or position (numeric) within \code{wb}.}

  \item{calcvar}{vector: variable to calculate the
  statistics for (usually numeric, can be logical).}

  \item{rowvar}{vector: categorical variable to stratify
  \code{calcvar}'s summaries over. Default behavior if left
  unspecified, is to calculate overall summaries with the
  row title "All".}

  \item{fun1,fun2}{two lists describing the utility
  functions that will calculate the statistics. Each list
  has a \code{fun} component for the function, and a
  \code{name} component for its name as it would appear in
  the column header.}

  \item{seps}{character vector of length 3, specifying the
  formatted separators before the output of
  \code{fun1$fun}, between the two outputs, and after the
  output of \code{fun2$fun}. Default behavior encloses the
  second output in parentheses. See 'Examples'.}

  \item{rowTitle}{character: the title to be placed above
  the row name column (default empty string)}

  \item{rowNames}{character vector of row names. Default
  behavior (\code{NULL}): automatically determined from
  data}

  \item{ord}{numeric vector specifying row-index order
  (i.e., a re-ordering of \code{rowvar}'s levels) in the
  produced table. Default (\code{NULL}) is no re-ordering.}

  \item{row1,col1}{numeric: the first row and column
  occupied by the table. In actuality, the first row will
  be \code{row1+2}, to allow for an optional header.}

  \item{purge}{logical: should \code{sheet} be created
  anew, by first removing the previous copy if it exists?
  (default \code{FALSE})}

  \item{...}{parameters passed on to
  \code{fun1$fun,fun2$fun}}

  \item{title}{character: title to be placed above table.}
}
\value{
  The function returns invisibly, after writing the data
  into \code{sheet}.
}
\description{
  Calculates univariate summary statistics (optionally
  stratified), and exports the formatted output to a
  spreadsheet.
}
\details{
  This function evaluates up to 2 univariate, functions on
  the input vector \code{calcvar}, either across the
  dataset or across strata defined by \code{rowvar}, and
  returns a single-column table with formatted results.

  The cell can be formatted to show a combined result, e.g.
  "Mean (SD)" which is the default. Tne function is quite
  mutable: both \code{fun1$fun, fun2$fun} and the strings
  separating their formatted output can be user-defined.
  The functions can return either a string (i.e., a
  formatted output) or a number that will be interpreted as
  a string in subsequent formatting. The default calls
  \code{\link{roundmean},\link{roundSD}} and prints the
  summaries in \code{"mean(SD)"} format.
}
\examples{
book2<-XLwriteOpen("chick2.xls") 
## Plain-vanilla
XLunivariate(book2,"weightByDiet",ChickWeight$weight,ChickWeight$Diet,
             title="Mean Weights by Diet",rowTitle="Diet")

## Replace mean/SD with median/range, put results beside previous
XLunivariate(book2,"weightByDiet",ChickWeight$weight,ChickWeight$Diet,
             title="Median Weights by Diet",rowTitle="Diet",col1=8,
             fun1=list(fun=roundmedian,name="Median"),fun2=list(fun=rangeString,name="range"))

### You can also do only one statistic... by "killing" one of the functions
XLunivariate(book2,"weightByAge",ChickWeight$weight,ChickWeight$Time,
             title="Mean Weights by Age",rowTitle="Age (Days)",seps=rep("",3),
             fun2=list(fun=emptee,name=""))
saveWorkbook(book2)
cat("Look for",paste(getwd(),"chick2.xls",sep='/'),"to see the results!\\n")
}
\author{
  Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
}
\seealso{
  Uses \code{\link{writeWorksheet}} to access the
  spreadsheet, \code{\link{rangeString}} for some utilities
  that can be used as \code{fun1$fun,fun2$fun}. For one-way
  (univariate) contingency tables, \code{\link{XLoneWay}}.
  See \code{\link{setStyleAction}} to control the output
  style.
}

