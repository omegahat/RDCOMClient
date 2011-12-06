library(RDCOMClient)

e = COMCreate("Excel.Application")

e[["Visible"]] = TRUE

sheet = e[["Workbooks"]]$Add()[["ActiveSheet"]]

k = sheet$Range("A1:A1")
k[["Value"]] = 10

k[["Value"]] = new("COMDate", 10)

k[["Value"]] = new("COMDecimal", 1.3)

k[["Value"]] = new("COMCurrency", 1000)




