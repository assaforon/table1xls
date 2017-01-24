##' "Table 1" Style List of Tables exported to a spreadsheet
##'
##' Formats and exports a series of shared-structure tables, and saves the file.
##' 
##' 
##' Details forthcoming... 
##' 
##' See the \code{\link{XLtwoWay}} help page, for behavior regarding new-sheet creation, overwriting, etc.
##' 
##' @author Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
##' @seealso \code{\link{XLoneWay}}.
##' 
##' 
##' @example inst/examples/Ex1way.r 
##' 
##' 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}. 
##' @param DF a rectangular array with all variables to be tabulated
##' @param colvar When \code{fun=\link{XLtwoWay}}, this specifies the column variable to cross-tabulate vs. the variables in \code{DF}.
##' @param fun The \code{table1xls} function to apply for each variable. Default \code{\link{XLoneWay}}.
##' @param title character: an optional overall title to the table. Default (\code{NULL}) is no title.
##' @param rowTitle character: the title to be placed above the row name column (default empty string)
##' @param rowNames character: vector of row names. Default behavior (\code{NULL}): automatically determined from data
##' @param ord numeric vector specifying row-index order in the produced table. Default (\code{NULL}) is no re-ordering.
##' @param row1,col1 numeric: the first row and column occupied by the table (title included if relevant).
##' @param purge logical should \code{sheet} be created anew, by first removing the previous copy if it exists? (default \code{FALSE})
##' @param digits numeric: how many digits (after the decimal point) to show in the percents? Defaults to 1 if n>=500, 0 otherwise.
##' @param combine logical: should counts and percents be combined to the popular \code{"Count(percent)"} format, or presented side-by-side? (default \code{TRUE}) 
##' @param useNA How to handle missing values. Passed on to \code{\link{table}} (see help on that function for options).
##' @param testname string, the *name* of a function to run a significance test on the table. Default \code{NULL} (no test).
##' @param pround number of significant digits in test p-value representation. Default 3.
##' @param testBelow logical, should test p-value be placed right below the table? Default \code{FALSE}, which places it next to the table's right edge, one row below the column headings
##' @param ... additional arguments as needed, to pass on to \code{get(textfun)}; for example, the reference frequencies for a Chi-Squared GoF test.
##' @return The function returns invisibly, after writing the data into \code{sheet} and saving the file.
##'
##' @export

XLtable1<-function(wb,sheet,DF,colvar=NULL,fun=XLoneWay,title="Table 1",rowTitle="Variable",row1=1,col1=1,digits=ifelse(dim(DF)[1]>=500,1,0),...,purge=FALSE)
{ 
  dims=dim(DF)
  if(length(dims)!=2) stop("Input must be rectangular array/DF.\n")
  if(!is.null(colvar)) fun<-XLtwoWay
  if(is.null(colvar) && identical(fun,XLtwoWay))
  {
    warning("No column variable specified; assuming last column is it.\n")
    colvar=DF[,dims[2]]
    DF=DF[-dims[2]]
    dims=dim(DF)
  }
  DF=as.data.frame(DF)
  if(purge) removeSheet(wb,sheet)
  if(!existsSheet(wb,sheet)) createSheet(wb,sheet)
  
  if(!is.null(title))  ### Adding a title
  {
    XLaddText(wb,sheet,text=title,row1=row1,col1=col1)
    row1=row1+1
  }
n=dims[1]
nvar=dims[2]

## Tabulating the variables
for(a in 1:nvar)
{
  fun(wb,sheet,DF[,a],colvar=colvar,row1=row1,col1=col1,rowTitle=names(DF)[a],totals=FALSE,digits=digits,...)
  if(a>1) for (b in 1:(1+length(unique(colvar)))) XLaddText(wb,sheet,text="",row1=row1,col1=col1+b)
  row1=row1+ifelse(identical(fun,XLunivariate),length(rowvar),length(unique(DF[,a]))+2
}
## Bottom summaries
if(identical(fun,XLoneWay))
{
  XLaddText(wb,sheet,"Sample Size (n)",row1=row1,col1=col1)
  XLaddText(wb,sheet,paste(n," (",niceRound(100,digits),"%)",sep=''),row1=row1,col1=col1+1)
}
if(identical(fun,XLtwoWay))
{
  XLaddText(wb,sheet,paste("Sample Sizes (total ",n,")",sep=''),row1=row1,col1=col1)
  enns=table(colvar,...)
  npct=paste(enns," (",niceRound(100*enns/n,digits),"%)",sep='')
  for (b in 1:(1+length(unique(colvar)))) XLaddText(wb,sheet,npct[b],row1=row1,col1=col1+b)
}

setColumnWidth(wb, sheet = sheet, column = col1:(col1+5), width=-1)
saveWorkbook(wb)
}  ### Function end
