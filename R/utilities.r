##'
##'
##' Open a spreadsheet document, while deleting the previous copy.
##'
##' 
##' Open a spreadsheet file (.xls or .xlsx), while deleting the previous copy if it exists.
##' 
##' The XLConnect function \code{\link{loadWorkbook}} can open existing spreadsheets or create new ones if they don't exist. 
##' However, it *cannot* delete the previous copy when opening the new one -- which is the default expected behavior of software such as R. 
##' As a result, analysts might inadvertently mix old and new versions of data and analyses, in the same spreadsheet.
##' 
##' This short utility mitigates the risk, by calling \code{\link{unlink}} first to make sure existing copies are deleted before the new spreadsheet file is opened.
##' 
##' @note Even though the workbook object is created, and is linked to a specific file name, it will only be saved to disk after \code{\link{saveWorkbook}} is called. See example. From \code{table1xls} version 0.3.0 on, all of the package's spreadsheet-export functions save the file by default. The example also illustrates some of the peculiarities of working with \code{\link{XLConnect}}, many of which are taken care of when using \code{table1xls} functions.
##' 
##' @param path character: the spreadsheet's full filename, including the extension. Only \code{.xls, .xlsx} extensions are allowed.
##' 
##' @return an XLConnect workbook object.
##' @example inst/examples/ExOpen.r
##' 
##' @author Assaf P. Oron \code{<aoron.at.idmod.org>}
##' 
##' @seealso \code{\link{loadWorkbook}}, \code{\link{saveWorkbook}}
##' @export

XLwriteOpen<-function(path)
{
unlink(path)
loadWorkbook(path,create=TRUE)
}

##' Utility functions for table summaries
##' 
##' 
##' Various auxiliary convenience functions, mostly for \code{\link{XLunivariate}}.

##' Functions calculating simple statistics and returning the output in a formatted manner, making it easier for \code{\link{XLunivariate}} to embed them in spreadsheet cells.
##' 
##' This is a small collection of useful utilities called by \code{\link{XLunivariate}}. They return 1-2 summary statistics, in a format that will not require additional formatting and formula-manipulation in Excel.
##' 
##' For example, \code{\link{roundmedian}} returns the median rounded to the specified number of digits, while \code{\link{iqrString}} returns the 1st and 3rd quartiles, separated by at least one dash (default 3 dashes). \code{\link{XLunivariate}} can combine these functions' output to produce the formatted summary \code{"median (Q1---Q3)"} often used in research articles.
##' 
##' In particular, \code{emptee} returns an empty string, enabling the use of  \code{\link{XLunivariate}} to produce only a single summary statistic per cell rather than a pair.
##' 
##' @example inst/examples/ExUnivar.r 
##'  
##' @param x vector (usually numeric, but can be logical) on which statistics are to be calculated
##' @param digits numeric: how many digits to round the output to?
##' @param quantmeth numeric: for functions calling \code{\link{quantile}}, the calculation method for the quantiles. Default is 7 to match the R default. Note that it is shrunk towards the median and hence biased, but typically with lower MSE. A very viable alternative is 6, the SAS/SPSS (and Stata?) default, which is unbiased. See the help on \code{\link{quantile}} for more details.
##' @param na.rm logical: should missing values be removed? (default \code{FALSE}) Passed onto the underlying functions
##' @param sep character: separating character for range- type functions.
##' @param ... this is ignored by the functions, but enables the "mixing and matching" of extra parameters between functions called by \code{\link{XLunivariate}}, without triggering an error.
##' 
##' @return The summary statistic(s), in the format specified via the arguments.
##' 
##' @seealso \code{\link{XLunivariate}} which is the main function calling these utilities.
##' 
##' @author Assaf P. Oron \code{<aoron.at.idmod.org>}
##' 
##' @export

rangeString<-function(x,digits=1,sep='-',na.rm=FALSE,...) paste(niceRound(min(x,na.rm=na.rm),digits),'-',niceRound(max(x,na.rm=na.rm),digits),sep=sep)

##' @rdname rangeString
#' @export

iqrString<-function(x,digits=1,sep='-',quantmeth=7,na.rm=FALSE,...) 
{
  tmp<-quantile(x,type=quantmeth,na.rm=na.rm,prob=c(1/4,3/4))
  paste(niceRound(tmp[1],digits),'-',niceRound(tmp[2],digits),sep=sep)
}

##' @rdname rangeString
#' @export

roundmean<-function(x,digits=1,na.rm=FALSE,...) niceRound(mean(x,na.rm=na.rm),digits=digits)

##' @rdname rangeString
#' @export

roundmedian<-function(x,digits=1,na.rm=FALSE,...) niceRound(median(x,na.rm=na.rm),digits=digits)

##' @rdname rangeString
#' @export

roundSD<-function(x,digits=1,na.rm=FALSE,...) niceRound(sd(x,na.rm=na.rm),digits=digits)

##' @rdname rangeString
#' @export

emptee<-function(x,...) ""

##' Write text to a single cell
##' 
##'   
##' Write text to a single cell in a specified file and sheet, and save the file.
##' 
##' 
##' Since XLConnect only exports data to spreadsheets as `data.frame`, this function sends the text as an on-the-fly `data.frame` with one column and one row, and without writing the header or the row name. 
##' 
##' @example inst/examples/ExGeneric.r
##' 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}.
##' @param text character: the text to be written to file.
##' @param row1,col1 integer: the row and column for the output. 
##' 
##' @return The function returns invisibly, after writing the data into \code{sheet} and saving the file.
##' 
##' @note If the specified \code{sheet} does not exist, the function will create it, assuming that was the user's intent (e.g., add a text-only sheet with explanations to a file.) This is hard-coded, because the inadvertent creation of single-text sheets due to typos can be easily discovered upon opening the file :) 
##' 
XLaddText<-function(wb,sheet,text,row1=1,col1=1)
{
  if(!existsSheet(wb,sheet)) createSheet(wb,sheet)  
  writeWorksheet(wb,data=text,sheet=sheet,startRow=row1,startCol=col1,header=FALSE)

#  if(wrap)
#  {
#    cs <- createCellStyle(wb)
#    setWrapText(cs, wrap = TRUE)
#    setCellStyle(wb,sheet=sheet,row=row1,col=col1,cellstyle = cs)
#    setRowHeight(wb, sheet = sheet, row=row1, height=-1)
#  } else  
setColumnWidth(wb, sheet = sheet, column = col1, width=-1)
  
  saveWorkbook(wb)
}


# Inactive:
# @param wrap logical: should text be wrapped keeping its width as is  - or should column width auto-matched to fit text on one line? (default \code{FALSE})

######################################

##' Rounding to a Predictable Number of Digits
##' 
##'   
##' Rounds numbers to always have the specified number of decimal digits, rather than R's "greedy" most-compact rounding convention. Includes optional "<0.0..." override adequate for representing small p-values.
##' 
##' 
##' R's standard \code{\link{round}} utility rounds to at *most* the number of digits specified. When the number happens to round more "compactly", it rounds to fewer digits. Thus, for example, `round(4.03,digits=1)` yields 4 as an answer. This is undesirable when trying to format a table, e.g., for publication.
##' 
##' \code{niceRound} solves this problem by wrapping a \code{\link{format}} call around the \code{\link{round}} call. The result will always have `digits` decimal digits. In addition, since reporting p-values always involves rounding, if the argument `plurb` is \code{TRUE}, then values below the rounding thresholds will be represented using the "less than" convention. For example, with \code{digits=3} and \code{plurb=TRUE}, the number 0.0004 will be represented as \code{<0.001}.
##' 
##' 
##' @param numbers the numbers to be rounded. Can also be a vector or numeric array.
##' @param digits the desired number of decimal digits
##' @param plurb logical, should the p-value-style "less-than blurb" convention be used? Default \code{FALSE}.
##' @export
##' @author Assaf P. Oron \code{<aoron.at.idmod.org>}
##' @seealso \code{\link{round}},\code{\link{format}}
##' 
niceRound<-function(numbers,digits=0,plurb=FALSE)
{
if(digits<0) stop('cannot round to negative number of digits.\n')
cand=format(round(numbers,digits=digits),nsmall=digits,trim=TRUE) 
if(plurb) cand[numbers<0.5/(10^digits-1)]=paste('<',10^(-digits),sep='')
cand
}

