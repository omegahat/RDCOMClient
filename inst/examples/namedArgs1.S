library(RDCOMClient)
e = COMCreate("Excel.Application")
e[["Workbooks"]]$Add()
wks = e[["Worksheets"]]

e[["Visible"]] = TRUE

#wks$Add(After=as.integer(2))
wks$Add(Count=as.integer(3))


