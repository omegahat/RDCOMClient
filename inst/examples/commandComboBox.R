library(RDCOMEvents)
library(RDCOMClient)
library(SWinTypeLibs)

l = LoadTypeLib("C:/Program Files/Common Files/Microsoft Shared/OFFICE11/mso.dll")

newDropdown = 
function(type = 3, values = c("Item 1", "Item 2", "Item 3"),
         app = COMCreate("Excel.Application"), caption = "R Test")
{
 if(missing(app)) {
   app[['visible']] = TRUE
   book = app$Workbooks()$Add()
 }

 bars = app$CommandBars()

 stdBar = bars$Item('Standard')

 button = stdBar$Controls()$Add(type, Temporary = TRUE)
 
 button[["Caption"]] = caption
 button[["Style"]] = 1

 for(i in values) 
   button$AddItem(i)

 if(type == 3)
   button[["ListIndex"]] = 1

 button
}

button = newDropdown(type = 4, caption = "Duncan")

Change =
function(Cntrl)
{
  cat("<Change>\n")
  app = Cntrl$Application
  cell = app$Cells(3,1)
  cell[['value']] = "The text in the combo box should show up below me."
  cell = app$Cells(3,2)
  cell[['value']] = Cntrl$Text()
  cat("</Change>\n")
}

print("Getting connection point for a command edit box on the standard toolbar")
if(FALSE) {

 points = getConnectionPoints(button)
 interface = l[[names(points)]]
 point = points[[1]] 

} else {
  interface = l[["_CommandBarComboBoxEvents"]]

  point = findConnectionPoint(button, gsub("[{}]", "", interface@guid))
}

s = createCOMEventServerInfo(interface, methods = list(Change = Change))

server = createCOMEventServer(s)

print("Connecting")
# Error: The server threw an exception
cookie = connectConnectionPoint(point, server)
