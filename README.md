table1xls
=========

This is an R package to generate tabular summaries in formats commonly found in scientific articles, and export them to Office-compatible spreadsheet document (.xls/.xlsx).

It answers a need commonly encountered by 

- analysts working with scientific investigators on academic publications;
- and generally, analysts needing to produce summaries for non-analyst audiences.

The package's functions save tedious work in converting R output to "normal human" summary tables. They also help ensure reproducible data analysis. One can see this as an Office-compatible version of 'xtable'.

The underlying functionality to open, write and manipulate spreadsheets is provided by the `XLConnect` package, which in turn relies upon `rJava`.

List of major functions:

 - `XLwriteOpen` to initiate/reset an export file
 - `XLoneWay` for one-way contingency tables, with or without percent breakdown
 - `XLtwoWay`, the same for two-way tables *(historically this was the first function)*
 - `XLunivariate` for summary statistics of continuous variables, potentially stratified by a categorical variable
 - `XLregresSummary` for standard 4-column regression output with confidence intervals and p-values, including transformations if desired
 - `XLgeneric` to export any rectangular dataset
 - `XLaddText` to add comments anywhere in the spreadsheet
 - **New for 0.4.0!** `XLtable1` to export an omnibus battery of tables typically encountered in Table 1 of an article. All tables must be of the same type, from among one-way, two-way or continuous.
 
Also worthy of mention is `niceRound`, a generally useful utility that reports numbers rounded to the actual specified number of digits, rather than annoyingly dropping digits if the number is "too round".
 

This package is available on GitHub (live copy) CRAN (periodic updates). If you get it from here, then when compiling the package, you might need to roxygenize to get all the help files.

Comments welcome, to assaf.oron at seattlechildrens.org





