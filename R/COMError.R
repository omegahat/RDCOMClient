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
         if(!is.character(val))
	    val = as.logical(val)
        .Call("RDCOM_setWriteError", val, PACKAGE = "RDCOMClient")
    } else
        .Call("RDCOM_getWriteError", PACKAGE = "RDCOMClient")
}
