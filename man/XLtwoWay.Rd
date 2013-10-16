\name{XLtwoWay}
\alias{XLtwoWay}
\title{Two-way Contingency Tables exported to a spreadsheet}
\usage{
  XLtwoWay(wb, sheet, rowvar, colvar, sumby = 1,
    rowTitle = "", rowNames = NULL, colNames = NULL,
    ord = NULL, row1 = 1, col1 = 1, header = FALSE,
    purge = FALSE, digits = 1)
}
\arguments{
  \item{wb}{a \code{\link[XLConnect]{workbook-class}}
  object}

  \item{sheet}{numeric or character: a worksheet name
  (character) or position (numeric) within \code{wb}.}

  \item{rowvar}{vector: categorical variable (logical,
  numeric, character, factor, etc.) for the table's rows}

  \item{colvar}{vector: categorical variable (logical,
  numeric, character factor, etc.) for the table's columns}

  \item{sumby}{whether percentages should be calculated
  across rows (1, default) or columns (2).}

  \item{rowTitle}{character: the title to be placed above
  the row name column (default empty string)}

  \item{rowNames,colNames}{character vector of row and
  column names. Default behavior (\code{NULL}):
  automatically determined from data}

  \item{ord}{numeric vector specifying row-index order in
  the produced table. Default (\code{NULL}) is no
  re-ordering.}

  \item{row1,col1}{numeric: the first row and column
  occupied by the table. In actuality, the first row will
  be \code{row1+2}, to allow for an optional header.}

  \item{header}{logical: should a header row with the
  captions "Counts:" and "Percentages:" be added above the
  tables? (default \code{FALSE})}

  \item{purge}{logical: should \code{sheet} be created
  anew, by first removing the previous copy if it exists?
  (default \code{FALSE})}

  \item{digits}{numeric: how many digits (after the decimal
  point) to show in the percents table?}
}
\value{
  The function returns invisibly, after writing the data
  into \code{sheet}.
}
\description{
  Producing 2-way descriptive tables, with percentages, and
  exporting to .xls/.xlsx Format
}
\details{
  This function produces two identical tables side by side,
  one with the cross-tabulated counts of unique values of
  \code{rowvar, colvar} and the other with percentages,
  calculated either by row (\code{sumby=1}) or column
  (\code{sumby=2}). Row and column margins are also
  automatically produced.

  The tables are then exported to worksheet \code{sheet} in
  workbook \code{wb}, using \code{\link{writeWorksheet}}
  from the XLConnect package.

  The worksheet \code{sheet} does not have to pre-exist;
  the function will create it if it doesn't already exist.

  There is an asymmetry between rows and columns, because
  the tables are converted to data frame in order for
  \code{\link{writeWorksheet}} to export them.
}
\note{
  This function uses the internal function \code{fancytab2}
  which produces 2-way tables with counts, percentages and
  margins.

  The tables won't be visible on a spreadsheet, until you
  use \code{\link{saveWorkbook}} to save and close
  \code{wb}.

  If \code{sheet} exists, it will be written into - rather
  than completely cleared and rewritten de novo. However,
  existing data in individual cells will be overwritten.
}
\examples{

### Contrived example looking at, e.g., the distribution of A-K-Q card counts
### in two partners' Bridge hands


hand1=rhyper(1000,12,40,13)
hand2=rhyper(1000,12-hand1,27+hand1,13)
handNames=c("0-1",2:4,"5 or more")


### The problem is ridiculously symmetric, so I de-symmetrize the presentation slightly:

book3<-XLwriteOpen("hands.xls") 
XLtwoWay(book3,"PartnersAKQcounts",cut(hand1,c(0,2:6,14)-0.5),cut(hand2,c(0,2:5,14)-0.5),
         rowTitle="Hand 1 vs. Hand 2",rowNames=c(handNames[-5],5,"6 or more","Total"),
         colNames=c(handNames,"Total"),header=TRUE)

### Same table, but percents now condition on columns rather than rows:
XLtwoWay(book3,"PartnersAKQcounts",cut(hand1,c(0,2:6,14)-0.5),cut(hand2,c(0,2:5,14)-0.5),
         rowTitle="Hand 1 vs. Hand 2",rowNames=c(handNames[-5],5,"6 or more","Total"),
         colNames=c(handNames,"Total"),header=FALSE,row1=10,sumby=2)

saveWorkbook(book3)
cat("Look for",paste(getwd(),"hands.xls",sep='/'),"to see the results!\\n")
}
\author{
  Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
}
\seealso{
  Uses \code{\link{writeWorksheet}} to access the
  spreadsheet. See \code{\link{setStyleAction}} to control
  the output style. If interested in one-way tables, see
  \code{\link{XLoneWay}}.
}

