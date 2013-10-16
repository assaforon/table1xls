\name{XLwriteOpen}
\alias{XLwriteOpen}
\title{Open a spreadsheet document, while deleting the previous copy.}
\usage{
  XLwriteOpen(path)
}
\arguments{
  \item{path}{character: the spreadsheet's full filename,
  including the extension. Only \code{.xls, .xlsx}
  extensions are allowed.}
}
\value{
  an XLConnect workbook object.
}
\description{
  Open a spreadsheet file (.xls or .xlsx), while deleting
  the previous copy if it exists.
}
\details{
  The XLConnect function \code{\link{loadWorkbook}} can
  open existing spreadsheets or create new ones if they
  don't exist. However, it *cannot* delete the previous
  copy when opening the new one -- which is the default
  behavior of software such as R. As a result, analysts
  might inadvertently mix old and new versions of data and
  results, in the same spreadsheet.

  This short utility mitigates this risk, by calling
  \code{\link{unlink}} first to make sure existing copies
  are deleted before the new spreadsheet file is opened.
}
\note{
  Even though the workbook object is created, and is linked
  to a specific file name, it will only be saved to disk
  after \code{\link{saveWorkbook}} is called. See example.
  The example also illustrates some of the peculiarities of
  working with \code{\link{XLConnect}}, many of which are
  taken care of when using \code{table1xls} functions.
}
\examples{
### Run this example in successive copy-paste batches

## Batch 1: be careful to copy and paste only the first 3 lines 
# *without* the white-space below them.
cat("R will now open a new .xls worksheet for you!\\n")
cat("Please enter path and filename, without extension:\\n")
filestring<-readLines(n=1)


# R is waiting for you... enter the filename ... then proceed to next batch.

## Batch 2
newPath<-paste(filestring,'xls',sep='.')
blankbook<-XLwriteOpen(newPath)

# If you check to see whether the file exists - it's not there.
# The spreadsheet is only in R's memory. The next batch will save it.

## Batch 3
saveWorkbook(blankbook)
cat("Now there should be a blank file called",newPath, "- Check it out!\\n")

## Now: writing into the file and resaving
# Make sure you close the file in case you opened it in Excel.
# We'll just write something silly there now:

## Batch 4
# Excel showed 1 blank sheet. But for R, there are 0 sheets until you create some.
createSheet(blankbook,"one") 
writeWorksheet(blankbook,"Something Sillee!!!",sheet='one') 
saveWorkbook(blankbook)

# Now it's not blank anymore - Check it out... 
# You will notice XLConnect has interpreted the string 
# as a data frame. Data transfer can only occur in the form of
# data frames (except some graphics).
# After closing the file run the last batch, which finally demonstrates 
# what XLwriteOpen itself does (open with overwrite).
# Don't forget to close the .xls file first!


## Batch 5

blankbook2<-XLwriteOpen(newPath)
saveWorkbook(blankbook2)

#### Now the file is blank again - Check it out!
#### All done!
}
\author{
  Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
}
\seealso{
  \code{\link{loadWorkbook}}, \code{\link{saveWorkbook}}
}

