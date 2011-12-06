library(RDCOMClient)
e = COMCreate("Excel.Application")
e[["Visible"]]=1
e$InputBox("Hi there", Top=30)

