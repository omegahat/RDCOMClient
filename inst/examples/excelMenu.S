library(RDCOMClient)
library(SWinTypeLibs)

e = COMCreate("Excel.Application")
e$Workbooks()$Add()
e[["Visible"]] = TRUE

bar = e[["CommandBars"]]$Item(3)
bar$Name() == "Standard"

ctrls = bar$Controls()

btn = ctrls$Add("Duncan", Type = as.integer(1), Temporary = TRUE)

btn[["Caption"]] = "Duncan"

