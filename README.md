table1xls
=========

This is an R package to generate tabular summaries in formats commonly found in scientific articles, and export them to Office-compatible spreadsheet document (.xls/.xlsx).

It answers a need commonly encountered by 

- analysts working with scientific investigators on academic publications;
- and generally, analysts needing to produce summaries for non-analyst audiences.

The package's functions save tedious work in converting R output to "normal human" summary tables. They also help ensure reproducible data analysis. One can see this as an Office-compatible version of 'xtable'.

The underlying functionality to open, write and manipulate spreadsheets is provided by the XLConnect package.

This package is available on GitHub (live copy) CRAN (periodic updates). If you get it from here, then when compiling the package, you might need to roxygenize to get all the help files.

Comments welcome, to assaf.oron at seattlechildrens.org

Wish List
=========

 - Add association p-values to `XLtwoWay` 
 - perhaps also to `XLoneWay`? (vs. some hypothesis)
 - Add p-value adjustment options to p-values
 - Mega-Table-1 wrapper (different variables in same table)



