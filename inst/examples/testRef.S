library(RDCOMClient)

xls <- COMCreate("Excel.Application")
wb <- xls[["Workbooks"]]$Add()

wbs <- xls[["Workbooks"]]
wbs$Add()
