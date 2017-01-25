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
##' @seealso \code{\link{XLoneWay}},\code{\link{XLtwoWay}},\code{\link{XLunivariate}}.
##' 
##' 
##' @example inst/examples/ 
##' 
##' 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}. 
##' @param DF a rectangular array with all variables to be tabulated.
##' @param colvar vector; specifies the variable to cross-tabulate for \code{fun=\link{XLtwoWay}} (see 'Details' for convenience options), or to stratify for \code{\link{XLunivariate}}. Has to be the entire variable, rather than just a name.
##' @param fun The \code{table1xls} function to apply for each variable. Default \code{\link{XLoneWay}}. Other supported functions are \code{\link{XLtwoWay},\link{XLunivariate}}.
##' @param title character: an optional overall title to the table. Default \code{"Table 1"}.
##' @param rowTitle character: the title to be placed above the row name column (default empty string)
##' @param colName character: when relevant, more descriptive names for columns in case \code{colvar} is used. Default \code{NULL}, which will use the unique values of \code{colvar} as names. 
##' @param row1,col1 numeric: the first row and column occupied by the table (title included if relevant).
##' @param purge logical should \code{sheet} be created anew, by first removing the previous copy if it exists? (default \code{FALSE})
##' @param digits numeric: how many digits (after the decimal point) to show in the percents? Defaults to 1 if n>=500 or if using \code{\link{XLunivariate}}, and 0 otherwise.
##' @param ... additional arguments as needed, to pass on to \code{fun}; for example, non-default summary function choices for \code{\link{XLunivariate}}.
##' @return The function returns invisibly, after writing the data into \code{sheet} and saving the file.
##'
##' @export

XLtable1<-function(wb,sheet,DF,colvar=NULL,fun=XLoneWay,sideBySide=FALSE,title="Table 1",rowTitle="Variable",colNames=NULL,row1=1,col1=1,digits=NULL,...,purge=FALSE)
{ 
  dims=dim(DF)
  if(length(dims)!=2) stop("Input must be rectangular array/DF.\n")
  if(!is.null(colvar) )
  {
    if(length(colvar)!=dims[1]) stop("Column variable length mismatch.\n")
    if(!identical(fun,XLunivariate)) fun<-XLtwoWay
  }
  if(is.null(colvar) && identical(fun,XLtwoWay)) {
    warning("No column variable specified; assuming last column is it.\n")
    colvar=DF[,dims[2]]
    DF=DF[-dims[2]]
    dims=dim(DF)
  }
  if(is.null(digits)) {
    digits=ifelse(dim(DF)[1]>=500,1,0)
    if(identical(fun,XLunivariate)) digits=1
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
  fun(wb,sheet,DF[,a],colvar=colvar,table1mode=TRUE,row1=row1,col1=col1,rowTitle=names(DF)[a],colNames=colNames,digits=digits,...)

  ## Removing some fluff output
  if(a>1) 
  {
    clearRange(wb,sheet,c(row1,ifelse(identical(fun,XLunivariate),col1,col1+1),row1,col1+1+length(unique(colvar))))
   }
  row1=row1+ifelse(identical(fun,XLunivariate),2,length(unique(DF[,a]))+2)
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
