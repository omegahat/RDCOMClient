library(RDCOMClient)
library(SWinTypeLibs)

e = COMCreate("Excel.Application")
e$Workbooks()$Add()
e[["Visible"]] = TRUE

bars = e[["CommandBars"]]

bar = bars$Item(3)

ctrls = bar$Controls()

btns =
 lapply(c("Graphics", "Close"),
        function(txt) {
	  btn = ctrls$Add(as.integer(1))
	  btn[["Style"]] = as.integer(2) # msoButtonCaption
	  btn[["Caption"]] = txt
	  btn
       })




