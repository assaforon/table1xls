##' "Table 1" Style List of Tables exported to a spreadsheet
##'
##' Formats and exports a series of shared-structure tables, and saves the file.
##' 
##' 
##' Auto-generation of a series of tables of the same type for a single dataset. One-way and two-way contingency tables and numerical summaries are all supported, but all summaries call the same atomic \code{fun}. 
##' 
##' The function employs convenience conventions for two-way tabulation: first, if \code{colvar} is specified and \code{fun} is left blank, then \code{fun} will be set to \code{\link{XLtwoWay}}. Second, if \code{fun=XLtwoWay} and \code{colvar} is left blank, then \code{colvar} will be set to the last column of \code{DF}.
##' 
##' For numerical summaries, use \code{fun=XLunivariate}. If you specify \code{colvar}, two-way summaries stratified by \code{colvar} will be returned.
##'  
##' Note that this function does not mix and match. Just make several calls to \code{XLtable1} with different sub-datasets and different values of \code{fun}, and combine the results in your report document.
##' 
##' In a similar vein, two-way summaries do not return the marginal one-way summaries as a byproduct. For example, if you use \code{fun=XLtwoWay}, then in order to get column totals for the generated two-way output, you will need to call \code{XLtable1} again on the same data, using the default \code{fun=XLoneWay}.
##' 
##' 
##' See the \code{\link{XLtwoWay}} help page, for behavior regarding new-sheet creation, overwriting, etc.
##' 
##' @author Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
##' @seealso \code{\link{XLoneWay}},\code{\link{XLtwoWay}},\code{\link{XLunivariate}}.
##' 
##' 
##' @example inst/examples/ExTable1.r 
##' 
##' 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}. 
##' @param DF a rectangular array with all variables to be tabulated.
##' @param colvar vector; specifies the variable to cross-tabulate for \code{fun=\link{XLtwoWay}} (see 'Details' for convenience options), or to stratify for \code{\link{XLunivariate}}. Has to be the entire variable, rather than just a name.
##' @param fun The \code{table1xls} function to apply for each variable. Default \code{\link{XLoneWay}}. Other supported functions are \code{\link{XLtwoWay},\link{XLunivariate}}.
##' @param title character: an optional overall title to the table. Default \code{"Table 1"}.
##' @param colTitle character: the title to be placed above the first column of the column variable. Default \code{NULL}.
##' @param colNames character: when relevant, more descriptive names for columns in case \code{colvar} is used. Default \code{NULL}, which will use the unique values of \code{colvar} as names. 
##' @param row1,col1 numeric: the first row and column occupied by the table (title included if relevant).
##' @param purge logical should \code{sheet} be created anew, by first removing the previous copy if it exists? (default \code{FALSE})
##' @param digits numeric: how many digits (after the decimal point) to show in the percents? Defaults to 1 if n>=500 or if using \code{\link{XLunivariate}}, and 0 otherwise.
##' @param useNA How to handle missing values. Passed on to \code{\link{table}} (see help on that function for options).
##' @param ... additional arguments as needed, to pass on to \code{fun}; for example, non-default summary function choices for \code{\link{XLunivariate}}.
##' @return The function returns invisibly, after writing the data into \code{sheet} and saving the file.
##'
##' @export

XLtable1<-function(wb,sheet,DF,colvar=NULL,fun=XLoneWay,title="Table 1",colTitle=NULL,colNames=NULL,row1=1,col1=1,digits=NULL,useNA='ifany',...,purge=FALSE)
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

if(!is.null(colTitle))  ### Adding a column title
  XLaddText(wb,sheet,text=colTitle,row1=row1,col1=col1+1) 
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
  fun(wb,sheet,DF[,a],colvar=colvar,table1mode=TRUE,row1=row1,col1=col1,rowTitle=names(DF)[a],colNames=colNames,digits=digits,useNA=useNA,...)

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
  enns=table(colvar,useNA=useNA)
  npct=paste(enns," (",niceRound(100*enns/n,digits),"%)",sep='')
  for (b in 1:(1+length(unique(colvar)))) XLaddText(wb,sheet,npct[b],row1=row1,col1=col1+b)
}

setColumnWidth(wb, sheet = sheet, column = col1:(col1+5), width=-1)
saveWorkbook(wb)
}  ### Function end
