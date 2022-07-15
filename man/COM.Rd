\name{.COM}
\alias{.COM}
\title{Full access to client COM invocation}
\description{
  This is the S function that provides full
  access to the C routines that perform the
  invocation of methods in COM servers.
  This allows one to control the specification
  of the dispatch method, whether the result
  is returned. 
}
\usage{
.COM(obj, name, ..., .dispatch = 3L, .return = TRUE, 
      .ids = numeric(0), .suppliedArgs)
}
\arguments{
  \item{obj}{the COM object reference, usually obtained via
    \code{\link{COMCreate}}.}
  \item{name}{the name of the method or property to be accessed.}
  \item{\dots}{arguments to be passed to the method.
    If names are provided for these arguments,
    these are used as the names in the COM call.
    (Not yet!)}
  \item{.dispatch}{one or more of the \code{\link{DispatchMethods}}
    values indicating the target of the invocation: a method or a
    property, and whether to get or set the property.
    In some cases, one wants to specify both a method and a property
    which is done by OR'ing the values in \code{\link{DispatchMethods}}
    in the bit-wise sense of OR'ing.
   }
  \item{.return}{a logical value indicating whether to bother
    returning the result of the call.
    This can be used to discard the result of the invocation
    when only the side-effect is of interest.
  }
  \item{.ids}{an optional numeric vector which, if given, 
   provides the \code{MEMBERID} values which identify the
   names of the method and parameters used in the call.
   Supplying these avoids the need for an extra communication step
   with the COM object to map the names to identifiers.
   One must compute these values using the type library
   (see \code{\link[SWinTypeLibs]{getNameIDs}}) or 
   via other type information gathered from the object or 
   another type library tool, e.g. oleviewer, Visual Basic Editor's type brower.
   }
  \item{.suppliedArgs}{a logical vector indicating which of the
    arguments provided by \dots are to be used. In general, this argument
    is not used in interactive use. However, when the code calling the \code{.COM}
    function is generated, this provides a way to specify which are actual arguments and which
    are default values for arguments.}
}
\value{
  An arbitrary value obtained from
  converting the value returned from
  the COM invocation.
}
\references{
\url{http://www.omegahat.net/RDCOMClient}
\url{http://www.omegahat.net/RDCOMServer}
\url{http://www.omegahat.net/SWinTypeLibs}
\url{http://www.omegahat.net/SWinRegistry}
}
\author{Duncan Temple Lang (duncan@r-project.org)}

\seealso{
  \code{\link{COMCreate}}
  \code{\link{COMAccessors}}
  \code{\link[SWinTypeLibs]{getNameIDs}}
}

\note{
 We have madee PROPERTYGET|METHOD the default
 for method invocation. In this case, this 
 function would become less commonly used.

 Also, we will add code to handle the
 DispatchMethods enumeration symbollically
 in the same we have for Gtk enumerations in 
 RGtk.
}

\examples{

 e <- COMCreate("Excel.Application")
 books <- e[["Workbooks"]]
 books$Add()
 
  # Now for the example!
 books$Item(1)

 sheets <- e[["Sheets"]]
 sheets$Item(1)
\dontrun{
  # We can index the list of sheets by sheet name.
  # This is not run here as the name is different for 
  # different languages.  
 sheets$Item("Sheet1")
}

# Now tidy up.
 e$Quit()
 rm(list = c("books", "e", "sheets"))
 gc()

\dontrun{
o = COMCreate("Excel.Application")

.COM(o, "Count", .dispatch = 2, .ids = id)
}
}
\keyword{interface}
