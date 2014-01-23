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
Version: \tab 0.3\cr
Date: \tab 2014-01-20\cr
License: \tab GPL 2\cr
}
Any statistician collaborating with scientists, especially in the health sciences, has to produce publication-grade summary tables. Most commonly these are "Table 1" style basic demographics, contingency tables, or regression-output summaries. As a minimum, statisticians need to provide tables, from which the lead author can easily and accurately generate these tables.

Outside of physics and mathematical fields, the vast majority of scientific investigators edit their manuscripts and tables in office-type software, usually the Microsoft suite. R can easily output data in \code{.csv} format; however, the analyst or someone in the investigative team still has to perform tedious work formatting the tables using office software, into a template that the lead author can use. Moreover, the manner in which regression summaries appear in manuscripts, is quite different from the typical output of \code{\link{summary.glm}} and similar R functions.

More important than the tedious labor involved, this manual "portage" of analysis output from the raw R/csv format to tabular format can be prone to data errors, and is not reproducible.

The package \code{table1xls} is meant to close this gap, enabling analysts and their collaborators to focus on the analysis and the science with peace of mind, while saving precious time. It relies upon the functionality offered by the \code{\link{XLConnect}} package. It can be seen as an Office-compatible baby version of the LaTex (and now also HTML) oriented \code{xtable}. While not as broad, it does offer some conveniences that the latter cannot match, in particular the possibility of packaging all your output tables as separate tabs in the same single spreadsheet document.

This is a new package! I will gladly accept requests for new functionalities, as well as comments and corrections on existing functions.

Thanks, Assaf

PS: \code{\link{XLConnect}} uses the \code{rJava} package. For some systems, additional configuration is needed to enable Java on R (for Windows machines tested so far it was plug-and-play; not so for some Macs). When manipulating large spreadsheet objects, the Java Virtual Machine might choke and issue error messages. If this happens, you will need to start a new session, and before loading \code{\link{XLConnect}} (either directly or via loding \code{table1xls}), write

\code{options(java.parameters = "-XmxYYYY")}

Where \code{YYYY} is the virtual-memory size in MB (has to be a multiple of 128 AFAIK). The default is only 128 MB. Keep in mind that you need to have substantially more RAM than the amount allocated to the JVM.
}
\author{
Assaf P. Oron.


Maintainer: Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
}



