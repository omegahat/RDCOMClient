library(RDCOMClient)
library(SWinTypeLibs)
library(RDCOMEvents)

setupExcel = function()
{
   e = COMCreate("Excel.Application")
   e$Workbooks()$Add()
   e[["Visible"]] = TRUE

   bars = e[["CommandBars"]]

#   bar = bars$Add("Duncan", as.integer(3), MenuBar = FALSE, Temporary = TRUE)

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

   btns
}

l = LoadTypeLib("C:/Program Files/Common Files/Microsoft Shared/OFFICE11/mso.dll")

# Want getTypeLib(btn) to be able to work, but it doesn't.

addHandler = function(btn, 
                      lib = LoadTypeLib("C:/Program Files/Common Files/Microsoft Shared/OFFICE11/mso.dll"), 
                      connect = FALSE)
{
  library(RDCOMEvents)
  
  event.info = lib[["_CommandBarButtonEvents"]]
  point = findConnectionPoint(btn, event.info)

# point = getConnectionPoints(btn)[[1]]

cat("Got the connection point\n")
print(point)

  library(RDCOMServer)
  sinfo = createCOMEventServerInfo(event.info, 
                                   methods = list("Click"=function(Ctrl, CancelDefault){
				                            print(Ctrl)
				                            cat("Value of CancelDefault", CancelDefault$value, "\n")
				                            CancelDefault$value <- TRUE
				                            cat("Value of CancelDefault", CancelDefault$value, "\n")
				                            CancelDefault$value <- FALSE
				                            cat("Back to FALSE", CancelDefault$value, "\n")
                                                            cat("Hi from click\n")
                                                          }))

cat("Create server info\n")

  server = createCOMEventServer(sinfo)


print(server)

cat("Created the event server\n")
  if(connect) {
    connectConnectionPoint(point, server)
    cat("Connected to the source\n")
  }

  return(list(server = server, point = point, serverInfo = sinfo))
}

doIt =
function(connect = TRUE)
{
  btns = setupExcel()
  addHandler(btns[[1]], l, connect)
}




