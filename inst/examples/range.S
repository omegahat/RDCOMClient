library(RDCOMClient)

e= COMCreate("Excel.Application")	
print(e[["Workbooks"]])
book = e[["Workbooks"]]$Add()
sheet = book[["ActiveSheet"]]

# r[["Value"]] <- list(1, "a", "b", "c", 20)
# r[["Value"]] <- list(list("a", "b", "c", "d"))  # letters[1:5] # list(1, 2, 3, 4, 5)
# r[["Value"]] <- letters[1:5]

# Can't do the following
# sheet$Range("A1:E3")[["Value"]] <- asCOMArray(matrix(as.numeric(1:(3*5)), 3, 5))

r= sheet$Range("A1:E3")
r[["Value"]] <- asCOMArray(matrix(as.numeric(1:(3*5)), 3, 5))


r = sheet$Range("A5:E7")
r[["Value"]] <- asCOMArray(matrix(as.integer(100 + 1:(3*5)), 3, 5))

r= sheet$Range("A9:E11")
r[["Value"]] <- asCOMArray(matrix(letters[1:(3*5)], 3, 5))

r = sheet$Range("A13:E15")
r[["Value"]] <- asCOMArray(matrix(rep(c(TRUE, FALSE), length = 3*5), 3, 5))


e[["Visible"]] <- 1

#book$Close()
#e <- NULL
gc()
