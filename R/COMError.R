COMStop =
function(msg, status, class = "COMError")
{
  e = simpleError(msg)
  e$status = status
  class(e) = c(class, class(e))
  stop(e)
}


writeErrors =
function(val = logical())
{
    if(length(val)) {
         if(!is.character(val)) {
	    val = as.logical(val)
	    if(val)
	       val = tempfile()
          } else {
      	    val = path.expand(val)
	    if(!file.exists(dirname(val)))
	       stop("directory/folder ", dirname(val), " does not exist")
          }
	    
        .Call("RDCOM_setWriteError", val, PACKAGE = "RDCOMClient")
	val
    } else
        .Call("RDCOM_getWriteError", PACKAGE = "RDCOMClient")
}
