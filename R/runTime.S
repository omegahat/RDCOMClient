#  This file is needed to run code that is generated 
# by the generateInterface() function.  Perhaps
# it should migrate to the RDCOMClient package.
#

setClass("CompiledCOMIDispatch", contains = "COMIDispatch")

setClass("CompiledCOMCoClass",
         representation(coclass = "character"),
         contains = "CompiledCOMIDispatch")	

setMethod("getItemClassName", "CompiledCOMCoClass", function(x) x@coclass[1]) #XXX first one for now.


setMethod("[[", c("CompiledCOMCoClass", "character"),
           function(x, i, j, ...) {
             x = as(x, x@coclass[1]) #XX first one
             x[[i]]
           })

#  x[["name"]]
setMethod("[[<-", c("CompiledCOMCoClass", "character"),
           function(x, i, j, ..., value) {
             orig = x
             x = as(x, x@coclass[1]) 
             x[[i]] <- value
             orig
           })

setMethod("$", c("CompiledCOMCoClass"), 
           function(x, name) {
             x = as(x, x@coclass)
	     do.call("$", list(x, name))
           })

setMethod("$<-", c("CompiledCOMCoClass", "character"),
           function(x, name, value) {
	     do.call("$<-", list(as(x, x@coclass), name, value))
             x
           })


#######################################################################################################

# Return the names of the methods and properties.

COMNames = 
 function(x) {
               ids = createTypeVarName(x, 
                                       c("GetProperty", "SetProperty", "Methods"), quote = FALSE)
               pkg = attr(class(x), "package")
	       env = if(pkg == ".GlobalEnv") globalenv() else getNamespace(pkg)
               
               sort(as.character(unlist(sapply(mget(ids, env), names))))
            }
setMethod("names", "CompiledCOMIDispatch", COMNames)

# Fetch the value of a property or return a function to invoke
# the named method.
setMethod("$", "CompiledCOMIDispatch",
            function(x, name) {

               # Do partial or complete matching depending on the value of 
               # a variable the user can set named StrictMethodNameExpansion.
             if(!exists("StrictMethodNameExpansion") || StrictMethodNameExpansion) 
                m = match
             else
                m = pmatch

               # Find out what the names of the lists of functions for accessing
               # properties and methods are called.
             ids = createTypeVarName(x, c("GetProperty", "Methods"), quote = FALSE)

               # Fetch the property accessor functions.
             pkg = attr(class(x), "package")	
	     env = if(pkg == ".GlobalEnv") globalenv() else getNamespace(pkg)

	     GetProperty = get(ids[1], env, mode = "list")
             Methods = get(ids[2],  env, mode = "list")

               # See if there is an element in the property accessor
               # with that name.
	     idx = m(name, names(GetProperty)) 
	     idx.methods = m(name, names(Methods))

               # No method and there is a property, then just return its value.
             if(is.na(idx.methods) && !is.na(idx)) {
                 # If so, invoke it, but first add x as an argument
                 # by changing the formal arguments.
                f = GetProperty[[idx]]
                return(f(x))
             } else if(!is.na(idx.methods) && is.na(idx)) {
                f = Methods[[idx.methods]]
                formals(f)[[".x"]] = x
		return(f)
             } else if(!is.na(idx.methods) && !is.na(idx)) {
                method = Methods[[idx.methods]]
                property = GetProperty[[idx.methods]]
                f = function() {
                   if(nargs() == 0)
                     property(.x)
                   else  {
                     eval(match.call(method))
                   }
                }
                formals(f) = formals(property)
                formals(f)[[".x"]] = x
                return(f)
             } else 
                stop(name, " is not a property or method for ", class(x))

              # If we did find an entry, adapt it so it has access to 
              # x in the future and return it.  Note that we have to 
              # mess with the formals differently here since it may 
              # not be an empty list.
#             f = Methods[[idx]]
#             formals(f) = c(formals(f), ".x" = x)
#	     f
            })



#  Fetch the value of a property. If the name identifies a method
#  then get the function and if it has all the parameters have default 
#  values, then invoke it. 
#  This is not vectorized.
setMethod("[[", c("CompiledCOMIDispatch", "character"),
            function(x, i, j, ...) {

               # Find out what the names of the lists of functions for accessing
               # properties and methods are called.
             ids = createTypeVarName(x, c("GetProperty", "Methods"), quote = FALSE)

             pkg = attr(class(x), "package")	
	     env = if(pkg == ".GlobalEnv") globalenv() else getNamespace(pkg)	

               # Fetch the property accessor functions.
             GetProperty = get(ids[1], env, mode = "list")

               # See if there is an element in the property accessor
               # with that name.
               # XXX what about partial matching?
	     idx = match(i, names(GetProperty)) 
             if(!is.na(idx)) {
                 # If so, invoke it, but first add x as an argument
                 # by changing the formal arguments.
                f = GetProperty[[idx]]
#		formals(f)[[".x"]] = x
                return(f(x))
             } else {

                Methods = get(ids[2], mode = "list")             
                idx = match(i, names(Methods))

		if(!is.na(idx)) {
                   f = Methods[[idx]]
# If we wanted to just return the function...
#		   formals(f)[[length(args)]] = x
#		   return(f)

		   args = formals(f)
		   args = args[-length(args)] # Get rid of .x
		   hasDefault = sapply(args, function(arg) {
                                               !(is.name(arg) && as.character(arg) == "")
                                             } )

		   if(all(hasDefault))
                      return(f(.x = x))
                }

                stop(COMPropertyAccessError(x, i))
             }
})

# For a numeric value, we are assuming that 
# we are dealing with a COM container/list so we call the Item()
# method.
if(FALSE) #XXX
 setMethod("[[", c("CompiledCOMIDispatch", "numeric"),
            function(x, i, j, ...) {
              x$Item(i) 
            })


setCompiledCOMProperty = 
          function(x, name, value) {
               # Find out what the names of the lists of functions for accessing
               # properties and methods are called.
             ids = createTypeVarName(x, c("SetProperty", "GetProperty"), quote = FALSE)

	     k = class(x)
	     ns <- NULL

             pkg = attr(class(x), "package")	
	     env = if(pkg == ".GlobalEnv") globalenv() else getNamespace(pkg)
	
               # Fetch the property accessor functions.
             SetProperty = get(ids[1], env, mode = "list")	

               # If there is no property to set (i.e. a function in the SetProperty)
               # then we have to be more careful. If this is part of an inline 
               # assignment such as doc$Range$Text = "Some text"
               # then we want to allow the assignment of the intermediate value for
               # "Range" to go through without a warning.
	     if(is.na(match(name, names(SetProperty)))) {
	         if(is.null(ns))
                   GetProperty = get(ids[2], env, mode = "list")

                    #XXX
                    # Here is a way to check if this is part of an inline assignment.
                 inlineAssignment = (as.character(sys.call()[[1]]) %in% c("$<-", "[[<-") 
                                        && as.character(sys.call()[[2]]) == "*tmp*")
                 if(!inlineAssignment || is.na(match(name, names(GetProperty)))) {
                   stop(class(x), " has no property (read or write) named ", name)
                 }
             } else {
	         SetProperty[[name]](x, value)
             }
           
             x
          }

setMethod("$<-",  c("CompiledCOMIDispatch", "character"),  setCompiledCOMProperty)

setMethod("[[<-", c("CompiledCOMIDispatch", "character"),
           function(x, i, j, ..., value) {
             setCompiledCOMProperty(x, i, value)
           })
       
setMethod("[", c("COMList", "numeric"),
      function(x, i, j, ..., drop = TRUE) {
	 if(all(i < 1)) 
            i = (1:length(x))[ i]

         sapply(i, function(index) x[[index]])
      })

setMethod("[", c("COMTypedNamedList", "numeric"),
      function(x, i, j, ..., drop = TRUE) {
	 ans = callNextMethod()
	 if(all(i < 1))
            i = (1:length(x))[ i]
	 names(ans) = names(x)[i]
	 ans
      })	

setMethod("[", c("COMTypedNamedList", "character"),
          function(x, i, j, ..., drop = TRUE) {
           ids = names(x)
           i = pmatch(i, ids)
           a = x[i]
  	   names(a) = ids[ i ]
  	   a
          })

setMethod("[[", c("COMTypedNamedList", "character"),
          function(x, i, j, ..., exact = NA) {
	   w = match(i, names(x))
           if(!is.na(w))
             x[[w]]
           else {
#XXX  
            do.call("$", list(x,i))
#             callNextMethod()
           }
          })
		
COMPropertyAccessError =
function(obj, name, class = "COMPropertyAccessError")
{
  e = simpleError(paste("No property named", name))
  e$object = obj
  e$property = name

  class(e) = c(class, class(e))
  e
}

setClass("CompiledCOMAccessor", contains = "function")

CompiledCOMAccessor =
function(f)
{
  new("CompiledCOMAccessor", f)	 
}			


if(FALSE) {
# Define setGeneric("help"...)
setMethod("help", "CompiledCOMAccessor",
          function(topic, offline = FALSE, package = NULL, lib.loc = NULL,
                   verbose = getOption("verbose"),
                   try.all.packages = getOption("help.try.all.packages"),
                   chmhelp = getOption("chmhelp"),
                   htmlhelp = getOption("htmlhelp"),
                   pager = getOption("pager")) {
 
              cat("There is no help yet for these automated accessor functions\n", stderr())
          })
}


#################################################################################################################################

setClass("EnumValue", representation("integer"),
           validity = function(object) {
# Check the names here. Unfortunately, we don't have the class name.
# Have to add the validity to each class.
                         if(length(names(object)) == 0)
                           return(paste("no name on the value for", class(object)))

                          TRUE
                     }
)

setMethod("show", "EnumValue", function(object) {
                                 x = as.integer(object)
                                 names(x) = names(object)
                                 show(x)
                               })

setGeneric("EnumValue",
            function(id, value, obj = new("EnumValue")) {
               standardGeneric("EnumValue")
           })

setMethod("EnumValue", c("character", "numeric", "EnumValue"),
#
# Constructor for EnumValue classes.
#
function(id, value, obj = new("EnumValue"))
{
  value = as.integer(value)
  names(value) = id

  obj@.Data  = value
  obj
}
)

setMethod("EnumValue", c("character", "EnumValue"),
 function(id, value, obj = new("EnumValue")) {
   coerceToEnumValue(id, class(value))
  })

setMethod("EnumValue", c("numeric", "EnumValue"),
 function(id, value, obj = new("EnumValue")) {
   coerceToEnumValue(id, class(value))
  })

setMethod("EnumValue", c("character", "missing", obj = "EnumValue"),
 function(id, value, obj = new("EnumValue")) {
   coerceToEnumValue(id, class(obj))
  })

setMethod("EnumValue", c("numeric", "missing", obj = "EnumValue"),
 function(id, value, obj = new("EnumValue")) {
   coerceToEnumValue(id, class(obj))
  })



# Should do this is in the validation or in general constructor.
coerceToEnumValue =
function(value, targetClass = as.character(sys.call(-2)[[3]]))
{
   # Get the definition for the enumeration values, i.e. the 
   # named vector of values. 

   defName = paste(targetClass, "Enum", sep = "")
   if(exists(defName, mode = "numeric")) {
        def = get(defName, mode = "numeric")

         # Now that we have the definition table, process the 
         # value we were given and get its entry in the table.
        if(is.character(value)) {
             # should we be kind and let partial matching work here with pmatch().
           idx = match(value, names(def))
        } else
           idx = match(value, def)

          # If there is no corresponding entry, signal an error.
        if(is.na(idx))
          stop("No such value (", value, ") in enumeration for class ", 
               targetClass, ". Values must from the set ", 
               paste(names(def), collapse = ", "))
  
          # Now create a new value with the value and name 
          # and virgin instance of the target class.
        EnumValue(names(def)[idx], def[idx], new(targetClass))
   } else {
       # no definition for the enumeration table in the conventional place,
       # so issue a warning and give back an entirely generic, unvalidated
       # EnumValue object.  Perhaps we should just throw an error or
       # allow the EnumValue class tell us where it's information is located
       # like COMNamedTypedList.
  
      warning("No enumeration table (named ", defName, ") defined for class ", 
               targetClass)
      new(targetClass, as.integer(value))
   }
}

# These won't be inherited.
setAs("numeric", 'EnumValue',
	 function(from) {
            coerceToEnumValue(from)
	 })

setAs("character", 'EnumValue',
	 function(from) {
            coerceToEnumValue(from)
	 })


#################################################################################################################################

# Used in generating R code to interface to Type Library definitions
# and also at run-time for the generated code.

setGeneric("createTypeVarName",
            function(className, var, quote = TRUE)
              standardGeneric("createTypeVarName"))

setMethod("createTypeVarName",
            "COMIDispatch",
# Map the given names in var to a unique and legitimate
# R variable name for the given class.
# 
function(className, var, quote = TRUE)
{
  createTypeVarName(class(className), var, quote)
})

setMethod("createTypeVarName",
            "CompiledCOMCoClass",
# Map the given names in var to a unique and legitimate
# R variable name for the given class.
# 
function(className, var, quote = TRUE)
{
  createTypeVarName(className@coclass, var, quote)
})

setMethod("createTypeVarName",
            "character",
function(className, var, quote = TRUE) {
  ans = paste("COM", className, var, sep = ".") 
  if(quote) {
    ans = paste("'", ans, "'", sep = "")
  }
  names(ans) = var

  ans
})

#################################################################################################################################


getCOMElements =
#
# XXX This should be merged with the names() method for a CompiledCOMIDispatch
# object, specifically it should call this 

function(type, env = NA, namesOnly = FALSE)
{
 if(is(type, "CompiledCOMIDispatch"))
   type = class(type)
	
 if(!isClass(type))
  stop(type, " is not the name of a class")

 if(!("CompiledCOMIDispatch" %in% names(getExtends(getClass(type)))))
   stop(type, " is not the name of a COMIDispatch type class. This only workds for CompiledCOMIDispatch classes.\nIf you want to know about a DCOM type, use the SWinTypeLibs package or the Object Browser in the Visual Basic Editor in Word/Excel")

 ids = paste("COM", type, c("GetProperty", "SetProperty", "Methods"), sep = ".")
 ans = lapply(ids,
               function(x) {
                  if(!is.na(env)) {
                     if(exists(x, env))
                          return(get(x, env))
                  } else {
		     if(exists(x))
                        return(get(x))
                  }
                  NULL
         })

 if(namesOnly) 
   ans = sapply(ans, function(x) sort(names(x)))

	
 names(ans) = c("Readable Properties", "Writeable Properties", "Methods")
 ans
}			
