book1<-XLwriteOpen("chick1.xls") 
XLoneWay(book1,"Diets",ChickWeight$Diet)
XLoneWay(book1,"Diets",ChickWeight$Diet,combine=FALSE,row1=10,rowTitle="Diet")
cat("Look for",paste(getwd(),"chick1.xls",sep='/'),"to see the results!\n")
