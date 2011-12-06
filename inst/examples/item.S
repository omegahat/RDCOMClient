library(RDCOMClient)
# Misebehaving. 
# Is Item a property or a function or both!
e <- COMCreate("Excel.Application")
books <- e[["Workbooks"]]
books$Open(gsub("/", "\\\\", system.file("examples", "duncan.xls", package="RDCOMClient")))
sheets <- e[["Sheets"]]
# Won't work
k =  sheets$Item("Sheet1")	
print(k)

# Same as.
.COM(sheets, "Item", "Sheet1", .dispatch = as.integer(3))

