require("RDCOMClient") || stop("You need to install the RDCOMClient package")

.COMInit()
e <- COMCreate("Excel.Application")

books <- e[["workbooks"]] 

fn <- system.file("examples", "duncan.xls", package = "RDCOMClient")
fn <- gsub("/", "\\\\", fn)
print(fn)
b = books$open(fn)

sheets = b[["sheets"]]

mySheet = sheets$Item(as.integer(1))

e[["Visible"]] <- TRUE

r = mySheet[["Cells"]]

v <- r$Item(as.integer(1), as.integer(1))
v[["Value"]]

v <- r$Item(as.integer(1), as.integer(3))
v[["Value"]]

