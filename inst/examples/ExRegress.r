
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

cat("Look for",paste(getwd(),"attenu.xls",sep='/'),"to see the results!\n")

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

