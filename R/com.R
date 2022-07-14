# The following constants can be computed by running disp.
DispatchMethods <-
 c("Method"= 1, "PropertyGet" = 2, "PropertyPut" = 4)
storage.mode(DispatchMethods) <- "integer"

.COMInit <-
function(status = TRUE)
{
 .Call("R_initCOM", as.logical(status), PACKAGE = "RDCOMClient")
}


COMCreate <-
function(name, ..., existing = TRUE)
{
     # Will want to allow for class IDs to be specified here.
 name <-  as.character(name)
 if(existing) {
      # force 
   ans = getCOMInstance(name, force = TRUE, silent = TRUE)
   if(!is.null(ans) && !is.character(ans))
     return(ans)
 }

 ans <- .Call("R_create", name, PACKAGE = "RDCOMClient")
 if(is.character(ans))
     stop(ans)
 
 ans
}




looksLikeUUID =
#
# Checks if a given string has the characteristic structure of a UUID.
# Returns
#   0 -  not a UUID
#   1 -  is a UUID with no enclosing braces ({ }),
#   2 -  a complete UUID, with enclosing braces.
#
#  We should use an enumeration for this.
function(str)
{
  # Without the perl = TRUE, we get a segmentation fault on Windows (at least).

  if(length(grep("^\\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\\}$", str, perl = TRUE)) > 0)
    return(2)

  as.integer(length(grep("^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$", str, perl = TRUE)) > 0)
}

getCOMInstance =
function(guid, force = TRUE, silent = FALSE)
{
   guid = as.character(guid)

   status = looksLikeUUID(guid)
   if(status == 1) {
        # substring(guid, 1, 1) != "{" && substring(guid, nchar(guid)) != "}" 
     warning("guid must have form '{.....}', i.e. the curly braces around the id.")
     guid = paste("{", guid, "}", sep = "")
   }

  ans = .Call("R_connect", guid, TRUE, PACKAGE = "RDCOMClient")

  if(is.character(ans)) {
    if(!force) {
      if(silent)
        return(ans)
      els
        stop(ans)
    } else {
       if(!silent)
          warning("creating a new instance of ", guid,
                   " rather than connecting to existing instance.")
       ans = COMCreate(guid, existing = FALSE)
    }
  }

  ans
}


getCLSID =
#
# Converts a human-readable application name to a class id (GUID).
# e.g.  getCLSID("Excel.Application") returns
#     "{00024500-0000-0000-C000-000000000046}"
#
#XXX Should be a UUID.
function(appName)
{
  .Call("R_getCLSIDFromName", as.character(appName), PACKAGE = "RDCOMClient")
}

setMethod("$<-", signature("COMIDispatch"), function(x, name, value) {
            stop("You probably meant to set a COM property. Please use x[[\"",name,"\"]] <- value")

# If we uncomment the following, we would have asymmetry
#  in that x$foo <- 1 would work but x$foo would not.
#	     .Call("R_setProperty", x, as.character(name), list(value), integer(0))
#	      x 

          })

setMethod("$", signature("COMIDispatch"), 
	    function(x, name){
	       function(...) {
	         .COM(x, name, ...)
	       }
	    })

setMethod("[[", c("COMIDispatch", "numeric"), 
	      function(x, i, j, ...) {
# if i is numeric, can check if there is an Item() method.
	       .COM(x, "Item", i)
	      })

setMethod("[[", "COMIDispatch", 
	      function(x, i, j, ...) {
# if i is numeric, can check if there is an Item() method.
	       .Call("R_getProperty", x, as.character(i), NULL, integer(0), PACKAGE = "RDCOMClient")
	      })


setMethod("[[<-", c("COMIDispatch", "character", "character"),
	    function(x, i, j, ..., value) {

            x[[as.character(unlist(c(i, j, ...)))]] <- value
            x
            })

setMethod("[[<-", c("COMIDispatch", "character", "missing"),
	    function(x, i, j, ..., value) {


	    if(length(i) > 1) {
	      tmp = x
	        for(id in i[-length(i)])
		     tmp = tmp[[ id ]]
		       .Call("R_setProperty", tmp, as.character(i[length(i)]), list(value), integer(0), PACKAGE = "RDCOMClient")
	    } else {

	    id = .Call("R_lookupPropName", x, as.character(i), PACKAGE = "RDCOMClient")
	    if(is.na(id)) {
	      
	       stop(structure(list(message = paste("no property ", i, " in this COM object"), call = NULL),
	             class = c("InvalidPropertyName", "DCOMError", "error", "condition")))
            }
            if(!.Call("R_isReadOnly", x, id, PACKAGE = "RDCOMClient"))
               .Call("R_setProperty", x, as.character(i), list(value), id, PACKAGE = "RDCOMClient")
       }
	      x 
  	    })

setMethod("[[<-", c("COMIDispatch", "numeric"),
	    function(x, i, j, ..., value) {
	      warning("x[[ number ]] <- value doesn't change anything in a COMIDispatch object")
	      x 
  	    })

.COM <-
 # Allows one to control the type of dispatch used in the COM Invoke() call.
 # Useful for getting properties and methods.
function(obj, name,  ..., .dispatch = as.integer(3), .return = TRUE, .ids=numeric(0),
           .suppliedArgs)
{
 # if(!missing(.ids)) { }

 .args = list(...)
 if(!missing(.suppliedArgs))
      .args =  .args[!is.na(.suppliedArgs)]
 val = .Call("R_Invoke", obj, as.character(name), .args, 
	  	as.integer(.dispatch), as.logical(.return), as.numeric(.ids), PACKAGE = "RDCOMClient")

 val
}


asCOMArray <-
function(obj)
{
 obj <- as.matrix(obj)
 .Call("R_create2DArray", obj, PACKAGE = "RDCOMClient")
}

createCOMReference <-
function(ref, className)
{
 if(!isClass(className)) {
   className = "COMIDispatch"
   warning("Using COMIDispatch instead of ", className)
 }

 obj = new(className)
 obj@ref = ref

 obj
}


isValidCOMObject = 
function(obj)
{
  if(!is(obj, "COMIDispatch"))
    return(FALSE)

  if(!.Call("R_isValidCOMObject", obj))
    return(FALSE)

   # Now run the .COM(obj, "isValidObject") 
  TRUE
}
