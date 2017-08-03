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
    if(length(val))
        .Call("RDCOM_setWriteError", as.logical(val), PACKAGE = "RDCOMClient")
    else
        .Call("RDCOM_getWriteError", PACKAGE = "RDCOMClient")
}
