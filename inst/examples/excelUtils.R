##
## $Id: excelUtils.R 3386 2003-05-16 15:28:47Z duncan $
##

## TODO: 
##   row.names for importing/exporting data.frames
##   dates/currency, general converters
##   error checking

"exportDataFrame" <-
function(df, at, ...)
## export the data.frame "df" into the location "at" (top,left cell)
## output the occupying range.
## TODO: row.names, more error checking
{
   nc <- dim(df)[2]
   if(nc<1) stop("data.frame must have at least one column")
   r1 <- at$Row()                   ## 1st row in range
   c1 <- at$Column()                ## 1st col in range
   c2 <- c1 + nc - 1                ## last col (*not* num of col)
   ws <- at[["Worksheet"]]

   ## headers

   hdrRng <- ws$Range(ws$Cells(r1, c1), ws$Cells(r1, c2))
   hdrRng[["Value"]] <- names(df)

   ## data

   rng <- ws$Cells(r1+1, c1)        ## top cell to put 1st column 
   for(j in seq(from = 1, to = nc)){
       cat("Column", j, "\n")
       exportVector(df[,j], at = rng, ...)
       rng <- rng$Next()            ## next cell to the right
   }
   invisible(ws$Range(ws$Cells(r1, c1), ws$Cells(r1 + nrow(df), c2)))
}

"exportVector" <-
function(obj, at = NULL, byrow = FALSE, ...)
## coerce obj to a simple (no attributes) vector and export to
## the range specified at "at" (can refer to a single starting cell);
## byrow = TRUE puts obj in one row, otherwise in one column.
## How should we deal with unequal of ranges and vectors?  Currently
## we stop, modulo the special case when at refers to the starting cell.
## TODO: converters (currency, dates, etc.)
{
   n <- length(obj)
   if(n<1) return(at)
   d <- c(at$Rows()$Count(), at$Columns()$Count())
   N <- prod(d)
   if(N==1 && n>1){     ## at refers to the starting cell
      r1c1 <- c(at$Row(), at$Column())
      r2c2 <- r1c1 + if(byrow) c(0,n-1) else c(n-1, 0)
      ws <- at[["Worksheet"]]
      at <- ws$Range(ws$Cells(r1c1[1], r1c1[2]), 
                     ws$Cells(r2c2[1], r2c2[2]))
   } 
   else if(n != N)
      stop("range and length(obj) differ")

   ## currently we can only export primitives...

   if(class(obj) %in% c("logical", "integer", "numeric", "character"))
      obj <- as.vector(obj)     ## clobber attributes
   else
      obj <- as.character(obj)  ## give up -- coerce to chars

   ## here we create a C-level COM safearray
   d <- if(byrow) c(1, n) else c(n,1)
   objref <- .Call("R_create2DArray", matrix(obj, nrow=d[1], ncol=d[2]))
   at[["Value"]] <- objref
   invisible(at)
}

"importDataFrame" <-
function(rng = NULL, wks = NULL, n.guess = 5, dateFun = as.chron.excelDate)
## Create a data.frame from the range rng or from the "Used range" in
## the worksheet wks.  The excel data is assumed to be a "database" (sic) 
## excel of primitive type (and possibly time/dates).
## We guess at the type of each "column" by looking at the first
## n.guess entries ... but it is only a very rough guess.
{
   if(is.null(rng) && is.null(wks))
      stop("need to specify either a range or a worksheet")
   if(is.null(rng))
      rng <- wks$UsedRange()          ## actual region
   else
      wks <- rng[["Worksheet"]]       ## need to query rng for its class
   n.areas <- rng$Areas()$Count()     ## must have only one region
   if(n.areas!=1)
      stop("data must be in a contigious block of cells")

   c1 <- rng$Column()                 ## first col
   c2 <- rng$Columns()$Count()        ## last col, provided contiguous region
   r1 <- rng$Row()                    ## first row
   r2 <- rng$Rows()$Count()           ## last row, provided contiguous region

   ## headers

   n.hdrs <- rng$ListHeaderRows()
   if(n.hdrs==0)
      hdr <- paste("V", seq(form=1, to=c2-c1+1), sep="")
   else if(n.hdrs==1) 
      hdr <- unlist(rng$Rows()$Item(r1)$Value2())   
   else {    ## collapse multi-row headers
      h <- vector("list", c2-c1+1)     ## list by column
      r <- rng$Range(rng$Cells(r1,c1), rng$Cells(r1+n.hdrs-1, c2))
      jj <- 1
      for(j in seq(from=c1, to=c2)){
         h[[jj]] <- unlist(r$Columns(j)$Value2()[[1]])
         jj <- jj+1
      }
      hdr <- sapply(h, paste, collapse=".")
   }
   r1 <- r1 + n.hdrs

   ## Data region 

   d1 <- wks$Cells(r1, c1)
   d2 <- wks$Cells(r2, c2)
   dataCols <- wks$Range(d1, d2)$Columns()
   out <- vector("list", length(hdr))
   for(j in seq(along = out)){
      f1 <- dataCols$Item(j)
      f2 <- f1$Value2()[[1]]
      f <- unlist(lapply(f2, function(x) if(is.null(x)) NA else x))
      cls <- guessExcelColType(f1)
      out[[j]] <- if(cls=="logical") as.logical(f) else f
   }
   names(out) <- make.names(hdr)
   as.data.frame(out)
}

"guessExcelColType" <-
function(colRng, n.guess = 5, hint = NULL)
## colRng points to an range object corresponding to one excel column
## e.g., colRng = rng$Columns()$Item("H")
## TODO: currently we return one of "logical", "numeric", "character"
## need to add "SCOMIDispatch"
{
   wf <- colRng[["Application"]][["WorksheetFunction"]]
   S.avail <- c("logical", "numeric", "integer", "character")
   ## we should get the following from the Excel type library
   fmt <- colRng[["NumberFormat"]]
   num.fmt <- c("general", "number", "currency", "accounting", 
                "percentage", "fraction", "scientific")

   fld <- colRng$Rows()
   n <- fld$Count()
   k <- min(n.guess, n)
   cls <- character(k)
   c1 <- colRng$Cells(1,1)
   c2 <- colRng$Cells(k,1)
   for(i in 1:k){
       x <- fld$Item(i)
       if(wf$IsText(x)) 
          cls[i] <- "character"
       else if(wf$IsNumber(x)) {
          if(tolower(fmt) %in% num.fmt)
             cls[i] <- "numeric"
          else 
             cls[i] <- "character"
          }
       else if(wf$IsLogical(x)) 
         cls[i] <- "logical"
       else if(wf$IsNA(x)) 
          cls[i] <- "NA"
       else 
          cls[i] <- "character"
    }
    ## if not all rows agree, use character type
    cls <- cls[cls %in% S.avail]
    if(length(cls)==0 || length(unique(cls))>1)
       return("character")
    else
       return(cls[1])
}

"as.chron.excelDate" <-
function(xlsDate, date1904 = FALSE)
{
   if(date1904){
      orig <- c(month=12, day=31, year=1903)
      off <- 0
   } 
   else {
      orig <- c(month=12, day=31, year=1899)
      off <- 1
   }
   chron(xlsDate - off, origin = c(month=12, day = 31, year = 1899))
}

"as.excelData.chron" <-
function(chronDate, date1904 = FALSE)
{
   if(date1904){
      orig <- c(month=12, day=31, year=1903)
      off <- 0
   } 
   else {
      orig <- c(month=12, day=31, year=1899)
      off <- 1
   }
   if(any(origin(chronDate)!=orig))
      origin(chronDate) <- orig
   as.numeric(chronDate) + off
}
