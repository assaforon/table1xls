##' Write generic data frame, matrix or table to a spreadsheet
##'

##' This function is a convenience wrapper for getting practically any rectangular data structure into a spreadsheet, without worrying about conversion or spreadsheet-writing technicalities.
##' 
##' If the structure is not a data frame (or inherited from one), but a table or matrix, the function will convert it into one using \code{\link{as.data.frame.matrix}}, because data frames are what the exporting function \code{\link{writeWorksheet}} can export.
##' 
##' The worksheet \code{sheet} does not have to pre-exist; the function will create it if it doesn't already exist. Also, the changes are automatically saved to file. 
##' 
##' @author Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
##' @seealso Uses \code{\link{writeWorksheet}} to access the spreadsheet. See \code{\link{setStyleAction}} to control the output style. For two-way contingency tables, see \code{\link{XLtwoWay}}.
##' 
##' 
##' @example inst/examples/ExGeneric.r
##' 
##' @title Write generic rectangular data to a spreadsheet
##' 
##' @param wb a \code{\link[XLConnect]{workbook-class}} object
##' @param sheet numeric or character: a worksheet name (character) or position (numeric) within \code{wb}. 
##' @param dataset the rectangular structure to be written. Can be a data frame, table, matrix or similar.
##' @param title character: an optional overall title to the table. Default (\code{NULL}) is no title.
##' @param addRownames logical: should a column of row names be added to the left of the structure? (default \code{FALSE})
##' @param rowTitle character: the title to be placed above the row name column (default "Name")
##' @param rowNames character: vector of row names. Default \code{rownames(dataset)}, but relevant only if \code{addRownames=TRUE}.
##' @param row1,col1 numeric: the first row and column occupied by the output. 
##' @param purge logical: should \code{sheet} be created anew, by first removing the previous copy if it exists? (default \code{FALSE})
##' 
##' @return The function returns invisibly, after writing the data into \code{sheet}.
##'
##' @export
##' @import XLConnect

XLgeneric<-function(wb,sheet,dataset,title=NULL,addRownames=FALSE,rowNames=rownames(dataset),rowTitle="Name",row1=1,col1=1,purge=FALSE)
{ 
  
  if (!("data.frame" %in% class(dataset))) dataset=as.data.frame.matrix(dataset)
  
  if(purge) removeSheet(wb,sheet)
  if(!existsSheet(wb,sheet)) createSheet(wb,sheet)
  
  if(!is.null(title))  ### Adding a title
  {
    XLaddText(wb,sheet,text=title,row1=row1,col1=col1)
    row1=row1+1
  }
  if(addRownames)
  {
      dataset=cbind(rowNames,dataset)
      names(dataset)[1]=rowTitle
  }
  writeWorksheet(wb,dataset,sheet,startRow=row1,startCol=col1)   

setColumnWidth(wb, sheet = sheet, column = col1:(dim(dataset)[2]+1), width=-1)

saveWorkbook(wb)
}  ### Function end
