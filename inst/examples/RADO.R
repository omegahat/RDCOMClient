##
## $Id: RADO.R 3340 2003-05-01 18:47:15Z duncan $
##

## R/ADO interface on top of RDCOMClient and SWinTypeLibs

require(RDCOMClient)


con <- COMCreate("ADODB.Connection")

con[["ConnectionString"]] <- paste(
    "DRIVER={MySQL ODBC 3.51 Driver}",
    "SERVER=tukey",
    "DATABASE=verizonWireless",
    "UID=verizon",
    "PWD=w1re1ess",
    sep = ";")

con$Open()

rs <- COMCreate("ADODB.RecordSet")
#rs$Open("show tables", con)

rs$Open("select * from summary_trial3", con)

if(T){
   ## this does not work as of RDCOMClient 0.3-0 (dd is NULL instead
   ## of some kind of array/list/data.frame)
   dd <- rs$GetRows()
print(dd)

} else {

   ## build dd piecemeal
   rs$MoveFirst()
   nc <- rs$Fields()$Count()
   nr <- rs$Fields()$Item(as.integer(0))$Count()
   dd <- vector("list", length=nc)
   nms <- character(nc)
   for(j in 1:nc){
      cat(rs$Fields()$Item(j-1)$Name(), "  ")
      nms[j] <- rs$Fields()$Item(j-1)$Name()
      dd[[i]] <- vector("list", length = nr)
   }
   cat("\n")

   i <- 1
   while(!rs$EOF()){
      cat("Row", i, ": ", sep="")
      for(j in 1:nc){
         v <- rs$Fields()$Item(j-1)$Value()
         cat(v, " ")
         dd[[j]][[i]] <- v
      }
      cat("\n")
      rs$MoveNext()
      i <- i + 1
   }
   dd <- lapply(dd, unlist)
}
