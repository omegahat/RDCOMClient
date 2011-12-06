library(RDCOMClient)
e <- COMCreate("Excel.Application")
books <- e[["workbooks"]]
fileName = system.file("examples", "duncan.xls", package = "RDCOMClient")
print(fileName)
books$open(gsub("/", "\\\\", fileName), Editable=FALSE, ReadOnly = TRUE)
cat("Finished open\n")
e[["Visible"]] <- TRUE

#XXX The foo.txt appears with "" around it in the dialog box.
e$GetSaveAsFileName(InitialFilename = "foo.txt", Title="Select a name for this file")

# Unnamed first argument and named second argument.
#e$GetSaveAsFileName("foo.txt", Title="Select a name for this file")
