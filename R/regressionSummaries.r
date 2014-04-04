##' 
##' Regression Summary Tables exported to a spreadsheet
##' 
##' Takes a vector of regression effect estimates and the corresponding standard errors, transforms to "human scale" if requested, calculates confidence-intervals, and exports a standard formatted summary table to a spreadsheet.
##'
##' This function produces a standard scientific-article regression summary table, given the raw regression output. The output table has 4 columns: effect name, its (optionally transformed) magnitude, a symmetric confidence interval (likewise transformed) and p-value.
##'
##'The formatted table is exported to \code{sheet}.

##' @author Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
##' @return The function returns invisibly, after writing the data into \code{sheet}.
##' @example inst/examples/ExRegress.r 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}. 
##' @param varnames character: a vector of effect names (column 1 of output table)
##' @param betas numeric: a vector of effect estimates
##' @param SE numeric: a vector of standard-error estimates for the effects
##' @param transfun transformation function for \code{betas,SE}, to produce columns 2-3 of the output. Defaults to \code{\link{identity}}. use {\code{\link{exp}}} for odds ratio or relative risk.
##' @param effname character: a string explaining what the effect stands for, e.g. "difference" (the default), "Odds Ratio", etc.
##' @param confac numeric: the proportionality factor for calculating confidence-intervals. Default produces 95% Normal intervals. 
##' @param pfun function used to calculate the p-value, based on the signal-to-noise ratio \code{betas/SE}. Default assumes two-sided Normal p-values.

##' @param title character: title to be placed above table.
##' @param roundig numeric: how many digits (after the decimal point) to round the effect estimate to?
##' @param pround numeric: how many digits (after the decimal point) to round the p-value to?
##' @param row1,col1 numeric: the first row and column occupied by the table. In actuality, the first row will be \code{row1+2}, to allow for an optional title.
##' @param purge logical: should \code{sheet} be created anew, by first removing the previous copy if it exists? (default \code{FALSE})

##' @export

XLregresSummary=function(wb,sheet,varnames,betas,SE,transfun=identity,effname="Difference",confac=qnorm(0.975),roundig=2,pfun=function(x) 2*pnorm(-abs(x)),pround=3,row1=1,col1=1,purge=FALSE,title=NULL)
{	
   
if(purge) removeSheet(wb,sheet)
if(!existsSheet(wb,sheet)) createSheet(wb,sheet)

if(length(varnames)!=length(betas) | length(varnames)!=length(SE)) stop("Mismatched lengths.\n")
  
dout=data.frame(Name=varnames,Effect=round(transfun(betas),roundig))
names(dout)[2]=effname
CIlow=transfun(betas-SE*confac)
CIhigh=transfun(betas+SE*confac)
dout$Confidence=paste("(",round(CIlow,roundig),',',round(CIhigh,roundig),")",sep='')
dout$Pvalue=round(pfun(betas/SE),pround)

writeWorksheet(wb,dout,sheet,startRow=row1+2,startCol=col1)
if(!is.null(title)) {
  writeWorksheet(wb,title,sheet,startRow=row1,startCol=col1)
  clearRange(wb,sheet,c(row1,col1,row1,col1+1))
}
setColumnWidth(wb, sheet = sheet, column = col1:(col1+3), width=-1)
saveWorkbook(wb)
} 
               
