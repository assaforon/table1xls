\name{XLoneWay}
\alias{XLoneWay}
\title{One-way Contingency Tables exported to a spreadsheet}
\usage{
  XLoneWay(wb, sheet, rowvar, rowTitle = "Value",
    rowNames = NULL, ord = NULL, row1 = 1, col1 = 1,
    purge = FALSE, digits = 1, combine = TRUE)
}
\arguments{
  \item{wb}{a \code{\link[XLConnect]{workbook-class}}
  object}

  \item{sheet}{numeric or character: a worksheet name
  (character) or position (numeric) within \code{wb}.}

  \item{rowvar}{vector: the categorical variable (logical,
  numeric, character, factor, etc.) to be tabulated}

  \item{rowTitle}{character: the title to be placed above
  the row name column (default empty string)}

  \item{rowNames}{character: vector of row names. Default
  behavior (\code{NULL}): automatically determined from
  data}

  \item{ord}{numeric vector specifying row-index order in
  the produced table. Default (\code{NULL}) is no
  re-ordering.}

  \item{row1,col1}{numeric: the first row and column
  occupied by the table.}

  \item{purge}{logical should \code{sheet} be created anew,
  by first removing the previous copy if it exists?
  (default \code{FALSE})}

  \item{digits}{numeric: how many digits (after the decimal
  point) to show in the percents?}

  \item{combine}{logical: should counts and percents be
  combined to the popular \code{"Count(percent)"} format,
  or presented side-by-side? (default \code{TRUE})}
}
\value{
  The function returns invisibly, after writing the data
  into \code{sheet}.
}
\description{
  Calculates a one-way contingency table in counts and
  percents, and exports a formatted output to a
  spreadsheet.
}
\details{
  This function performs a one-way contingency table, also
  calculating the distribution in percents.

  The table is then exported to worksheet \code{sheet} in
  workbook \code{wb}, either using the format
  \code{"Count(percent)"} (if \code{combine=TRUE}), or as
  two separate columns in the same table. Worksheet writing
  uses \code{\link{writeWorksheet}} from the XLConnect
  package.

  The worksheet \code{sheet} does not have to pre-exist;
  the function will create it if it doesn't already exist.
}
\examples{
book1<-XLwriteOpen("chick1.xls") 
XLoneWay(book1,"Diets",ChickWeight$Diet)
XLoneWay(book1,"Diets",ChickWeight$Diet,combine=FALSE,row1=10,rowTitle="Diet")
saveWorkbook(book1)
cat("Look for",paste(getwd(),"chick1.xls",sep='/'),"to see the results!\\n")
}
\author{
  Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
}
\seealso{
  Uses \code{\link{writeWorksheet}} to access the
  spreadsheet. See \code{\link{setStyleAction}} to control
  the output style. If interested in one-way summary, see
  \code{\link{XLunivariate}}. For two-way contingency
  tables, see \code{\link{XLtwoWay}}.
}

