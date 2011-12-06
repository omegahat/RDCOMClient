library(RDCOMClient)
e <- COMCreate("Excel.Application")
e[["Visible"]]
e[["StandardFontSize"]]

e[["StartupPath"]]
e[["Path"]] #
e[["PathSeparator"]] # Check characters.	

e[["StatusBar"]] # VARIANT returned.

e[["SheetsInNewWorkbook"]]


# Functions
e$CentimetersToPoints(20.)
e$CheckSpelling("Duncan")
e$CheckSpelling("Duncna")

e$Calculate()
e$CalculateFull()			