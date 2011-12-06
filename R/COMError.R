COMStop =
function(msg, status, class = "COMError")
{
  e = simpleError(msg)
  e$status = status
  class(e) = c(class, class(e))
  stop(e)
}
