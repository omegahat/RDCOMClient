# Checks the conversion of the different data types
# by reading a prepared Excel spreadsheet.
#


library(RDCOMClient)

fileName = system.file("examples", "types.xls", package = "RDCOMClient")

if(fileName == "")
 stop("No such file ", "examples/types.xls")

fileName = gsub("/", "\\\\", fileName)

e = COMCreate("Excel.Application")

book = e[["Workbooks"]]$Open(fileName)
# e[["Visible"]] = 1

range = book[["ActiveSheet"]][["UsedRange"]]
d = range[["Value"]]

book$Close()
e$Quit()
e = NULL

names(d[[1]]) = unlist(d[[2]])
d = d[[1]]
print(d)



