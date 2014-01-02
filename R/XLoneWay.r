##' One-way Contingency Tables exported to a spreadsheet
##'
##' Calculates a one-way contingency table in counts and percents, and exports a formatted output to a spreadsheet.
##' 
##' This function performs a one-way contingency table, also calculating the distribution in percents. 
##' 
##' The table is then exported to worksheet \code{sheet} in workbook \code{wb}, either using the format \code{"Count(percent)"} (if \code{combine=TRUE}), or as two separate columns in the same table. Worksheet writing uses \code{\link{writeWorksheet}} from the XLConnect package.
##' 
##' The worksheet \code{sheet} does not have to pre-exist; the function will create it if it doesn't already exist. 
##' 
##' 
##' @author Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
##' @seealso Uses \code{\link{writeWorksheet}} to access the spreadsheet. See \code{\link{setStyleAction}} to control the output style. If interested in one-way summary, see \code{\link{XLunivariate}}. For two-way contingency tables, see \code{\link{XLtwoWay}}.
##' 
##' 
##' @example inst/examples/Ex1way.r 
##' 
##' 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}. 
##' @param rowvar vector: the categorical variable (logical, numeric, character, factor, etc.) to be tabulated
##' @param rowTitle character: the title to be placed above the row name column (default empty string)
##' @param rowNames character: vector of row names. Default behavior (\code{NULL}): automatically determined from data
##' @param ord numeric vector specifying row-index order in the produced table. Default (\code{NULL}) is no re-ordering.
##' @param row1,col1 numeric: the first row and column occupied by the table. 
##' @param purge logical should \code{sheet} be created anew, by first removing the previous copy if it exists? (default \code{FALSE})
##' @param digits numeric: how many digits (after the decimal point) to show in the percents?
##' @param combine logical: should counts and percents be combined to the popular \code{"Count(percent)"} format, or presented side-by-side? (default \code{TRUE}) 
##' 
##' @return The function returns invisibly, after writing the data into \code{sheet}.
##'

XLoneWay<-function(wb,sheet,rowvar,rowTitle="Value",rowNames=NULL,ord=NULL,row1=1,col1=1,purge=FALSE,digits=1,combine=TRUE)
{ 
  
  if(purge) removeSheet(wb,sheet)
  if(!existsSheet(wb,sheet)) createSheet(wb,sheet)
  
  n=length(rowvar)
  tab=table(rowvar,useNA='ifany')
  percents=round(tab*100/n,digits=digits)
  names(tab)[is.na(names(tab))]="missing"
  tabnames=names(tab)
  tab=c(tab,n)
  names(tab)=c(tabnames,"Total")
  percents=c(percents,100)
  if (is.null(ord)) ord=1:length(tab)
  if (!is.null(rowNames)) names(tab)=rowNames

if(combine) {
    
    dout=data.frame(cbind(names(tab),paste(tab,' (',percents,')',sep="")))
    names(dout)=c(rowTitle,"Count (%)")
    
} else {

  dout=data.frame(cbind(names(tab),tab,percents))
  names(dout)=c(rowTitle,"Count","Percent")
}
  writeWorksheet(wb,dout[ord,],sheet,startRow=row1,startCol=col1)   

setColumnWidth(wb, sheet = sheet, column = col1:(col1+5), width=-1)
  
}  ### Function end
