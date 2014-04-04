t1<-XLwriteOpen("generic1.xls") 
XLgeneric(t1,"s1",matrix(1:4,nrow=2),row1=5,addRownames=TRUE)
XLgeneric(t1,"s1",matrix(1:4,nrow=2),col1=5,addRownames=TRUE,
          rowTitle="What?",rowNames=c("Hey","You!"))
cat("Look for",paste(getwd(),"generic1.xls",sep='/'),"to see the results!\n")
