\name{table1xls-package}
\alias{table1xlsL-package}
\alias{table1xls}
\docType{package}
\title{Publication-Grade Tables Exported to .xls/.xlsx spreadsheets}
\description{Generate tabular summaries in formats commonly found in scientific articles, and export them to Office-compatible spreadsheet document (.xls/.xlsx format).}

\details{
\tabular{ll}{
Package: \tab table1xls\cr
Type: \tab Package\cr
Version: \tab 0.4.0\cr
Date: \tab 2017-02-28\cr
License: \tab GPL-3\cr
}
Any statistician collaborating with scientists, especially in the health sciences, has to produce publication-grade summary tables. Most commonly these are "Table 1" style basic demographics, contingency tables, or regression-output summaries. As a minimum, statisticians need to provide tables, from which the lead author can easily and accurately generate these tables.

Outside of physics and mathematical fields, the vast majority of scientific investigators edit their manuscripts and tables in office-type software, usually the Microsoft suite. R can easily output data in \code{.csv} format; however, the analyst or someone in the investigative team still has to perform tedious work formatting the tables using office software, into a template that the lead author can use. Moreover, the manner in which regression summaries appear in manuscripts, is quite different from the typical output of \code{\link{summary.glm}} and similar R functions.

More important than the tedious labor involved, this manual "portage" of analysis output from the raw R/csv format to tabular format can be prone to data errors, and is not reproducible.

The package \code{table1xls} is meant to close this gap, enabling analysts and their collaborators to focus on the analysis and the science with peace of mind, while saving precious time. It relies upon the functionality offered by the \code{\link{XLConnect}} package. It can be seen as an Office-compatible baby version of the LaTex (and now also HTML) oriented \code{xtable} (or the SAS-inspired \code{tables}). \code{table1xls} does offer some conveniences that these other packages cannot match, in particular the possibility of packaging all your output tables as separate tabs in the same single spreadsheet document, and the ability to lay out related tables side-by-side or below each other in the same sheet.

Regarding export to non-Microsoft spreadsheet software: with my LibreOffice (Windows version) the .xls exports work perfectly fine, while .xlsx is not as reliable. I recommend using the former, of course.

You can download the latest package source code directly from GitHub, via the command 

\code{devtools::install_github("assaforon/table1xls")}.

If you have \code{roxygen2} installed, then the help pages will show following a GitHub download, thanks to a hack originating/disseminated by Yihui Xie.

I will gladly accept requests for new functionalities, as well as comments and corrections on existing functions.

Thanks, Assaf

PS: \code{\link{XLConnect}} uses the \code{rJava} package. For some systems, additional configuration is needed to enable Java on R. For Windows machines tested so far it was pretty much plug-and-play; the Linux CRAN tests all seem to work; not so for some Macs. 

When manipulating large spreadsheet objects, the Java Virtual Machine might choke and issue error messages. If this happens, you will need to start a new session, and before loading \code{\link{XLConnect}} (either directly or via loading \code{table1xls}), write

\code{options(java.parameters = "-XmxYYYY")}

Where \code{YYYY} is the virtual-memory size in MB. The default is only 128 MB. The string "-Xmx1g" is interpreted as 1 gigabyte. Keep in mind that you need to have substantially more RAM than the amount allocated to the JVM.
}
\author{
Assaf P. Oron.


Maintainer: Assaf P. Oron \code{<aoron@idmod.org>}
}



