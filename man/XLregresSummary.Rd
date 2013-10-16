\name{XLregresSummary}
\alias{XLregresSummary}
\title{Regression Summary Tables exported to a spreadsheet}
\usage{
  XLregresSummary(wb, sheet, varnames, betas, SE,
    transfun = identity, effname = "Difference",
    confac = qnorm(0.975), roundig = 2,
    pfun = function(x) 2 * pnorm(-abs(x)), pround = 3,
    row1 = 1, col1 = 1, purge = FALSE, title = NULL)
}
\arguments{
  \item{wb}{a \code{\link[XLConnect]{workbook-class}}
  object}

  \item{sheet}{numeric or character: a worksheet name
  (character) or position (numeric) within \code{wb}.}

  \item{varnames}{character: a vector of effect names
  (column 1 of output table)}

  \item{betas}{numeric: a vector of effect estimates}

  \item{SE}{numeric: a vector of standard-error estimates
  for the effects}

  \item{transfun}{transformation function for
  \code{betas,SE}, to produce columns 2-3 of the output.
  Defaults to \code{\link{identity}}. use
  {\code{\link{exp}}} for odds ratio or relative risk.}

  \item{effname}{character: a string explaining what the
  effect stands for, e.g. "difference" (the default), "Odds
  Ratio", etc.}

  \item{confac}{numeric: the proportionality factor for
  calculating confidence-intervals. Default produces 95%
  Normal intervals.}

  \item{pfun}{function used to calculate the p-value, based
  on the signal-to-noise ratio \code{betas/SE}. Default
  assumes two-sided Normal p-values.}

  \item{title}{character: title to be placed above table.}

  \item{roundig}{numeric: how many digits (after the
  decimal point) to round the effect estimate to?}

  \item{pround}{numeric: how many digits (after the decimal
  point) to round the p-value to?}

  \item{row1,col1}{numeric: the first row and column
  occupied by the table. In actuality, the first row will
  be \code{row1+2}, to allow for an optional title.}

  \item{purge}{logical: should \code{sheet} be created
  anew, by first removing the previous copy if it exists?
  (default \code{FALSE})}
}
\value{
  The function returns invisibly, after writing the data
  into \code{sheet}.
}
\description{
  Takes a vector of regression effect estimates and the
  corresponding standard errors, transforms to "human
  scale" if requested, calculates confidence-intervals, and
  exports a standard formatted summary table to a
  spreadsheet.
}
\details{
  This function produces a standard scientific-article
  regression summary table, given the raw regression
  output. The output table has 4 columns: effect name, its
  (optionally transformed) magnitude, a symmetric
  confidence interval (likewise transformed) and p-value.

  The formatted table is exported to \code{sheet}.
}
\examples{

### Regression summaries - multiple formats:


book4<-XLwriteOpen("attenu.xls") 

quakenames=c("Magnitude (Richter), per unit","Distance (log km), per x10")

# Ground acceleration as a function of magnitude and distance, all log scale.
quakemod1=summary(lm(log10(accel)~mag+log10(dist),data=attenu))


## Model-scale summaries; we don't care for the intercept.

XLregresSummary(book4,"ModelScale",varnames=quakenames,
                betas=quakemod1$coef[-1,1],SE=quakemod1$coef[-1,2],
                pround=6,title="Log-Ground Acceleration Effects",
                confac=qt(0.975,179),pfun=function(x) 2*pt(-abs(x),df=179))

## Effects arguably more meaningful as percent change. So still same model, but different summaries:

XLregresSummary(book4,"Multiplier",varnames=quakenames,
                betas=quakemod1$coef[-1,1],SE=quakemod1$coef[-1,2],
                pround=6,title="Relative Ground Acceleration Effects",transfun=function(x) 10^x,
                effname="Multiplier",confac=qt(0.975,179),pfun=function(x) 2*pt(-abs(x),df=179))

saveWorkbook(book4)
cat("Look for",paste(getwd(),"attenu.xls",sep='/'),"to see the results!\\n")

### lm() does not take account of station or event level grouping.
### So we use a mixed model, losing 16 data points w/no station data:
### Run this on your own... and ask the authors of "lme4" about p-values at your own risk :)

# library(lme4)
# quakemod2=lmer(log10(accel)~mag+log10(dist)+(1|event)+(1|station),data=attenu)
# 
# XLregresSummary(book4,"MixedModel",varnames=quakenames,betas=fixef(quakemod2)[-1],
# SE=sqrt(diag(vcov(quakemod2)))[-1],
# pround=6,title="Relative Ground Acceleration Effects",
# transfun=function(x) 10^x,effname="Multiplier",
# confac=qt(0.975,160),pfun=function(x) 2*pt(-abs(x),df=160))

}
\author{
  Assaf P. Oron \code{<assaf.oron.at.seattlechildrens.org>}
}

