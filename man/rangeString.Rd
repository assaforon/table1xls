\name{rangeString}
\alias{emptee}
\alias{iqrString}
\alias{rangeString}
\alias{roundmean}
\alias{roundmedian}
\alias{roundSD}
\title{Utility functions for table summaries}
\usage{
  rangeString(x, digits = 1, sep = "-", na.rm = FALSE, ...)

  iqrString(x, digits = 1, sep = "-", quantmeth = 7,
    na.rm = FALSE, ...)

  roundmean(x, digits = 1, na.rm = FALSE, ...)

  roundmedian(x, digits = 1, na.rm = FALSE, ...)

  roundSD(x, digits = 1, na.rm = FALSE, ...)

  emptee(x, ...)
}
\arguments{
  \item{x}{vector (usually numeric, but can be logical) on
  which statistics are to be calculated}

  \item{digits}{numeric: how many digits to round the
  output to?}

  \item{quantmeth}{numeric: for functions calling
  \code{\link{quantile}}, the calculation method for the
  quantiles. Default is 7 to match the R default. Note that
  it is shrunk towards the median and hence biased, but
  typically with lower MSE. A very viable alternative is 6,
  the SAS/SPSS (and Stata?) default, which is unbiased. See
  the help on \code{\link{quantile}} for more details.}

  \item{na.rm}{logical: should missing values be removed?
  (default \code{FALSE}) Passed onto the underlying
  functions}

  \item{sep}{character: separating character for range-
  type functions.}

  \item{...}{this is ignored by the functions, but enables
  the "mixing and matching" of extra parameters between
  functions called by \code{\link{XLunivariate}}, without
  triggering an error.}
}
\value{
  The summary statistic(s), in the format specified via the
  arguments.
}
\description{
  Various auxiliary convenience functions, mostly for
  \code{\link{XLunivariate}}. Functions calculating simple
  statistics and returning the output in a formatted
  manner, making it easier for \code{\link{XLunivariate}}
  to embed them in spreadsheet cells.
}
\details{
  This is a small collection of useful utilities called by
  \code{\link{XLunivariate}}. They return 1-2 summary
  statistics, in a format that will not require additional
  formatting and formula-manipulation in Excel.

  For example, \code{\link{roundmedian}} returns the median
  rounded to the specified number of digits, while
  \code{\link{iqrString}} returns the 1st and 3rd
  quartiles, separated by at least one dash (default 3
  dashes). \code{\link{XLunivariate}} can combine these
  functions' output to produce the formatted summary
  \code{"median (Q1---Q3)"} often used in research
  articles.

  In particular, \code{emptee} returns an empty string,
  enabling the use of \code{\link{XLunivariate}} to produce
  only a single summary statistic per cell rather than a
  pair.
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
  \code{\link{XLunivariate}} which is the main function
  calling these utilities.
}

