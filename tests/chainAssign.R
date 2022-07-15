if(FALSE) {
library(RDCOMClient)
e <- COMCreate("Excel.Application")
books <- e[["workbooks"]] 

fn <- system.file("examples", "duncan.xls", package = "RDCOMClient")
fn <- gsub("/", "\\\\", fn)
b = books$open(fn)

sheets = b[["sheets"]]
mySheet = sheets$Item(as.integer(1))

e[["Visible"]] <- TRUE

r = mySheet[["Cells"]]
v <- r$Item(as.integer(1), as.integer(1))


#library(SWinTypeLibs)
#f = getFuncs(v)


#file.remove("DCOM.err")
#RDCOMClient::writeErrors("DCOM.err")

v[["Interior"]][["ColorIndex"]] = 3

e$Quit()
}