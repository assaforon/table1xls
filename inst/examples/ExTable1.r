table1<-XLwriteOpen("table1.xls") 

## A default, option-free call generates one-way tables
XLtable1(table1,'cars1',mtcars[,c(2,8:11)])
## You can prettify a bit, first by changing variable names

names(mtcars)[c(2,8:11)]=c("Cylinders","V/S","Auto/Manual","Gears","Carbureutors")
XLtable1(table1,'cars1',mtcars[,c(2,8:11)],title="'mtcars': Summary of Categorical Variables",col1=4)

## Now two-way, generated implicitly by specifying 'colvar' (unless fun=XLunivariate)
XLtable1(table1,'cars2',mtcars[,8:11],colvar=mtcars$Cylinders)


cat("Look for",paste(getwd(),"table1.xls",sep='/'),"to see the results!\n")
