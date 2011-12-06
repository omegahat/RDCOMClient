library(RDCOMClient)
data(quakes)
source("excelUtils.R")

xls <- COMCreate("Excel.Application")
xls[["Visible"]] <- TRUE

wb = xls[["Workbooks"]]$Add()

# gctorture()

quakes = quakes[,1:2]

sh = wb$Worksheets(1)
exportDataFrame(quakes, at = sh$Range("A1"))
cat("Done first export\n")

sh = wb$Worksheets(2)
exportDataFrame(quakes, at = sh$Range("A1"))

cat("Done second export\n")

sh = wb$Worksheets(2)
exportDataFrame(quakes, at = sh$Range("A1"))


cat("Done third export\n")


for(i in 1:15) { 
 print(i)
 sh = wb$Worksheets(2)
 exportDataFrame(quakes, at = sh$Range("A1"))
}

for(i in 1:15) { 
 print(i)
 sh = wb[["Worksheets"]]$Add()
 exportDataFrame(quakes, at = sh$Range("A1"))
}

wb$Close()
xls[["Visible"]] <- FALSE

rm(wb)
rm(xls)
rm(sh)

gc()

