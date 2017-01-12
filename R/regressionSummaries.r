##' 
##' Regression Summary Tables exported to a spreadsheet
##' 
##' Takes regression effect estimates and the corresponding standard errors, transforms to "human scale" if requested, calculates confidence-intervals and p-values, and exports a standard formatted summary table to a spreadsheet.
##' 
##' This function produces a standard scientific-article regression summary table, given the raw regression output. The resulting table has 4 columns: effect name, its (optionally transformed) magnitude, a probabilistically symmetric confidence interval (likewise transformed), and p-value. The formatted table is exported to \code{sheet}, and the file is immediately saved.
##' 
##' The input can be provided as separate vectors of point estimates (\code{betas}) and standard errors (\code{SE}), or as a single matrix for \code{betas}. In the latter case, as a default the effect names will be \code{rownames(betas)}, unless a vector with more descriptive names is provided via \code{varnames}.
##' 
##' See the \code{\link{XLtwoWay}} help page, for behavior regarding new-sheet creation, overwriting, etc.

##' @author Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
##' @return The function returns invisibly, after writing the data into \code{sheet}.
##' @example inst/examples/ExRegress.r 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}. 
##' @param betas numeric: a vector of point estimates, or a matrix containing estimates and standard errors in columns
##' @param SE numeric: a vector of standard-error estimates for the effects. If \code{NULL} (default), user needs to specify them via the \code{betas} matrix.
##' @param varnames character: a vector of effect names (column 1 of output table). If \code{NULL} (default), user needs to specify them via the \code{betas} matrix.
##' @param colid integer: vector of indices for the columns containing the point estimates and SEs, respectively. Used only if \code{betas} is a matrix.
##' @param transfun transformation function for \code{betas,SE}, to produce columns 2-3 of the output. Defaults to \code{\link{identity}}. use {\code{\link{exp}}} for odds ratio or relative risk.
##' @param effname character: a string explaining what the effect stands for, e.g. "difference" (the default), "Odds Ratio", etc.
##' @param alpha numeric, Type I error for CIs. Default 0.05 for 95\% CIs.
##' @param df numeric, residual degrees of freedom. If a finite value is provided, t-distribution p-value and CIs will be calculated; otherwise Normality is assumed. Default \code{NA}.
##' ##' @param title character: an optional overall title to the table. Default (\code{NULL}) is no title.
##' @param roundig numeric: how many digits (after the decimal point) to round the effect estimate to?
##' @param pround numeric: how many digits (after the decimal point) to round the p-value to? P-values rounded down to zero will show up as "<" the smallest nonzero value, e.g. with the default \code{pround=3} p-values smaller than 0.0005 will show up as "<0.001".
##' @param row1,col1 numeric: the first row and column occupied by the table (title included if relevant).
##' @param purge logical: should \code{sheet} be created anew, by first removing the previous copy if it exists? (default \code{FALSE})

##' @note The default CI's are 95\% and Normal. P-values are also derived from the Normal. If you run any regression whose intervals are calculated differently (e.g., linear regression with not-huge sample size), make sure to change both \code{confac} and \code{pfun} accordingly, as is shown in the example.
##' @export

XLregresSummary=function(wb,sheet,betas,SE=NULL,varnames=NULL,colid=1:2,transfun=identity,title=NULL,effname="Difference",alpha=0.05,df=NA,roundig=2,pround=3,row1=1,col1=1,purge=FALSE)
{	
   
if(purge) removeSheet(wb,sheet)
if(!existsSheet(wb,sheet)) createSheet(wb,sheet)

## Matrix input convenience feature
if(is.matrix(betas))
{
  if(is.null(varnames)) varnames=rownames(betas)
  SE=betas[,colid[2]]
  betas=betas[,colid[1]]
}
if(length(varnames)!=length(betas) | length(varnames)!=length(SE)) stop("Mismatched lengths.\n")

## CIs and p-values
#  Default is Normal:
confac=qnorm(1-alpha/2)
pfun=function(x) 2*pnorm(-abs(x))

# If finite df are specified, then t(df) replaces them:
if(is.finite(df) && df>0) 
{
  confac=qt(1-alpha/2,df=df)
  pfun=function(x) 2*pt(-abs(x),df=df)
}
  
if(!is.null(title))  ### Adding a title
{
  XLaddText(wb,sheet,text=title,row1=row1,col1=col1)
  row1=row1+1
}

dout=data.frame(Effect=varnames,Magnitude=format(round(transfun(betas),roundig),nsmall=roundig,trim=TRUE))
names(dout)[2]=effname
CIlow=transfun(betas-SE*confac)
CIhigh=transfun(betas+SE*confac)
dout$CI=paste("(",niceRound(CIlow,roundig),',',niceRound(CIhigh,roundig),")",sep='')
names(dout)=gsub('CI',paste('CI',100*(1-alpha),sep=''),names(dout))
dout$Pvalue=niceRound(pfun(betas/SE),digits=pround,plurb=TRUE)

writeWorksheet(wb,dout,sheet,startRow=row1,startCol=col1)

setColumnWidth(wb, sheet = sheet, column = col1:(col1+3), width=-1)
saveWorkbook(wb)
} 
               
