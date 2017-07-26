##' One-way Contingency Tables exported to a spreadsheet
##'
##' Calculates a one-way contingency table in counts and percents, exports a formatted output to a spreadsheet, and saves the file.
##' 
##' This function performs a one-way contingency table, also calculating the distribution in percents. 
##' 
##' The table is then exported to worksheet \code{sheet} in workbook \code{wb}, either using the format \code{"Count(percent)"} (if \code{combine=TRUE}), or as two separate columns in the same table. 
##' 
##' See the \code{\link{XLtwoWay}} help page, for behavior regarding new-sheet creation, overwriting, etc.
##' 
##' @author Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
##' @seealso If interested in other descriptive summaries, see \code{\link{XLunivariate}}. For two-way contingency tables, see \code{\link{XLtwoWay}}.
##' 
##' 
##' @example inst/examples/Ex1way.r 
##' 
##' 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}. 
##' @param rowvar vector: the categorical variable (logical, numeric, character, factor, etc.) to be tabulated.
##' @param table1mode logical: is the function called from \code{\link{XLtable1}}? If \code{TRUE}, some modifications will be made to the output. Default \code{FALSE}.
##' @param title character: an optional overall title to the table. Default (\code{NULL}) is no title.
##' @param rowTitle character: the title to be placed above the row name column (default empty string)
##' @param rowNames character: vector of row names. Default behavior (\code{NULL}): automatically determined from data
##' @param colNames dummy argument for compatibility with calls from [XLtable1()]. Otherwise ignored by function.
##' @param ord numeric vector specifying row-index order in the produced table. Default (\code{NULL}) is no re-ordering.
##' @param row1,col1 numeric: the first row and column occupied by the table (title included if relevant).
##' @param purge logical should \code{sheet} be created anew, by first removing the previous copy if it exists? (default \code{FALSE})
##' @param digits numeric: how many digits (after the decimal point) to show in the percents? Defaults to 1 if n>=500, 0 otherwise.
##' @param combine logical: should counts and percents be combined to the popular \code{"Count(percent)"} format, or presented side-by-side? (default \code{TRUE}) 
##' @param useNA How to handle missing values. Passed on to \code{\link{table}} (see help on that function for options).
##' @param testname string, the *name* of a function to run a significance test on the table. Default \code{NULL} (no test).
##' @param pround number of significant digits in test p-value representation. Default 3.
##' @param testBelow logical, should test p-value be placed right below the table? Default \code{FALSE}, which places it next to the table's right edge, one row below the column headings
##' @param margins logical: should margins with totals be returned? Default \code{TRUE}.
##' @param ... additional arguments as needed, to pass on to \code{get(textfun)}; for example, the reference frequencies for a Chi-Squared GoF test.
##' @return The function returns invisibly, after writing the data into \code{sheet} and saving the file.
##'
##' @export

XLoneWay<-function(wb,sheet,rowvar,table1mode=FALSE,title=NULL,rowTitle="Value",rowNames=NULL,colNames=NULL,ord=NULL,row1=1,col1=1,digits=ifelse(length(rowvar)>=500,1,0),combine=TRUE,useNA='ifany',testname=NULL,testBelow=FALSE,margins=TRUE,...,purge=FALSE,pround=3)
{ 
  if(table1mode) margins<-FALSE  
  if(purge) removeSheet(wb,sheet)
  if(!existsSheet(wb,sheet)) createSheet(wb,sheet)
  
  if(!is.null(title))  ### Adding a title
  {
    XLaddText(wb,sheet,text=title,row1=row1,col1=col1)
    row1=row1+1
  }
    
  n=length(rowvar)
  tab=table(rowvar,useNA=useNA)
  percentab=niceRound(tab*100/n,digits=digits)
  names(tab)[is.na(names(tab))]="missing"
  tabnames=names(tab)
  if(margins)
  {
    tab=c(tab,n)
    names(tab)=c(tabnames,"Total")
    percentab=c(percentab,niceRound(100,digits))
  }
  if (is.null(ord)) ord=1:length(tab)
  if (!is.null(rowNames)) names(tab)=rowNames

if(combine) {
    
    dout=data.frame(cbind(names(tab),paste(tab,' (',percentab,'%)',sep="")))
    names(dout)=c(rowTitle,"Count (%)")
    
} else {

  dout=data.frame(nam=names(tab),Count=tab,Percent=percentab)
  names(dout)[1]=rowTitle
}
  writeWorksheet(wb,dout[ord,],sheet,startRow=row1,startCol=col1)   

  ### Perform test and p-value on table
  if(!is.null(testname) && length(unique(rowvar))>1)
  {
    pval=suppressWarnings(try(get(testname)(table(rowvar),...)$p.value))
    ptext=paste(testname,'p:',ifelse(is.finite(pval),niceRound(pval,pround,plurb=TRUE),'Error'))
    prow=ifelse(testBelow,row1+dim(dout)[1]+1,row1+1)
    pcol=ifelse(testBelow,col1,col1+dim(dout)[2])
    XLaddText(wb,sheet,ptext,row1=prow,col1=pcol)
  }
  
  
setColumnWidth(wb, sheet = sheet, column = col1:(col1+5), width=-1)
saveWorkbook(wb)
}  ### Function end
