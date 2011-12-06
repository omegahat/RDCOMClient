# A test for defining  classes so that we can do something like
#  doc$Range()$Text = "Some text"

setClass("Base")

setClass("Range", representation(text = "character"), contains = "Base")
setClass("Document", representation(range = "Range"), contains = "Base")


setMethod("$", c("Document", "character"),
           function(x, name) {
cat("In Document $\n")
              if(name != "Range")
                stop("No field other than Range supported")
 
              function() {
                 x@range
              }
           })


"Range<-" =
function(x, value)
{
 cat("In Range<-\n")
 x@text = value
 x
}

setMethod("$<-", c("Document", "character"),
           function(x, name, value) {
cat("In Document $<-\n")
              if(name != "Range")
                stop("No field other than Range supported")
 
              function() {
browser()
                 x
              }
           })

setMethod("$", c("Range", "character"),
           function(x, name) {
cat("In Range $\n")
              if(name != "Text")
                stop("No field other than Text supported")

              x@text
           })


setMethod("$<-", c("Range", "character"),
           function(x, name, value) {
cat("In Range $<-\n")
              if(name != "Text")
                stop("No field other than Text supported")

              x@text = value
              x 
           })
