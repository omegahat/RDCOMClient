library(RDCOMClient)
   xls <- COMCreate("Excel.Application")
   xls[["Visible"]] <- TRUE
   xls[["Workbooks"]]$Add()

   sh1 <- xls[["Sheets"]]$Item(1)
   oles <- sh1[["OLEObjects"]]
   oleButton <- oles$Add(ClassType = "Forms.CommandButton.1",
                Top = 10, Left = 10, Width = 144, Height = 30)

   acx <- oleButton[["Object"]]
   acx[["Caption"]] <- "Compute Means"

#   getConnectionPoints(acx)
library(SWinTypeLibs)
formLib = LoadTypeLib("C:\\WINNT\\System32\\FM20.DLL")


pt = findConnectionPoint(acx, formLib[["CommandButtonEvents"]])

library(RDCOMServer)
sinfo = createCOMEventServerInfo(formLib[["CommandButtonEvents"]], methods = list("Click"=function(...){cat("Hi from click\n")}))


server = createCOMEventServer(sinfo)

connectConnectionPoint(pt, server)
