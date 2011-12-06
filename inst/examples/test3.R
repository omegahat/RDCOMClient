library(RDCOMClient)
data(quakes)
source("excelUtils3.R")

#gctorture()
xls <- COMCreate("Excel.Application")
xls[["Visible"]] <- TRUE

wbs <- xls[["Workbooks"]]
book1 <- wbs$Add()

wss <- book1[["Worksheets"]]
ws <- wss$Item(1)
at1 <- ws$Range("A1")

exportDataFrame(quakes, at = at1)

cat("Importing data frame now\n")
qq1 <- importDataFrame(wks = ws)

all.equal(quakes, qq1)

if(FALSE) {

ws <- wss$Item(2)
at2 <- ws$Range("A1")
exportDataFrame(quakes, at = at2)
qq2 <- importDataFrame(wks = ws)
all.equal(quakes, qq2)
}

#wb$Close(saveChanges = FALSE)
#xls$Quit()

