require("RDCOMClient") || stop("You need to install the RDCOMClient package")

# .COMInit()
e <- COMCreate("Excel.Application")

books <- e[["workbooks"]] 

fn <- system.file("examples", "duncan.xls", package = "RDCOMClient")
fn <- gsub("/", "\\\\", fn)
print(fn)
b = books$open(fn)

sheets = b[["sheets"]]

mySheet = sheets$Item(as.integer(1))

e[["Visible"]] <- TRUE

r = mySheet[["Cells"]]

v <- r$Item(as.integer(1), as.integer(1))
v[["Value"]]

v <- r$Item(as.integer(1), as.integer(3))
v[["Value"]]

  # Set the value. Will appear in the Excel sheet.
v[["Value"]] = 10

v = r$Item(1L, 5L)
v[["Value"]]	# NULL
v[["Value"]] = "abc"


library(SWinTypeLibs)
# f = names(getFuncs(v))
v[["Value"]] = "A long title with many words we need to wrap"
v[["WrapText"]] = TRUE
v[["Interior"]][["ColorIndex"]] = 3

#i = v[["Interior"]]
#i[["Bold"]] = TRUE
#i[["Italic"]] = TRUE
#i[["ColorIndex"]] = 4



#err = tryCatch( v[["Interior"]][["ColorIndex"]] <- 4, error = function(e) e)



#sty = v[["Style"]]
#sty[["WrapText"]]
#sty[["WrapText"]] = TRUE



if(FALSE) {


f = getFuncs(sty)

# style currently applies to all cells.
fnt = sty[["Font"]]
names(getFuncs(fnt))

fnt[["Bold"]] = TRUE
fnt[["Italic"]] = TRUE
fnt[["Size"]] = 14
fnt[["ColorIndex"]] = 3
}
