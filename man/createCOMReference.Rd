\name{createCOMReference}
\alias{createCOMReference}
\title{Creates S object for COM reference}
\description{
 This is the function that is called anytime a COM object is being
created or returned from a C language call to S. This function can
examine the object and determine what is the best representation in S
for it. It can use the generic \code{IUnknown} or \code{COMIDispatch}
classes to simply represent the pointer value.
Alternatively, it might dynamically generate a new S class
and accessor methods for accessing properties and functions for
that object using the \code{SWinTypeLibs}
package. Or it might lookup a previously compiled collection
of type information and match the GUID of the object's type 
to find the associated S class name.

This function is not intended to be called from S, but primarily from
C code that has access to the COM references.
}
\usage{
createCOMReference(ref, className)
}

\arguments{
  \item{ref}{the S object that is an external pointer containing the reference to the
   COM object. }
  \item{className}{The name of the class that is ``suggested'' by the caller.}
}
\value{
 An S object that should almost definitely contain
 the \code{ref} value. The class of the object should probably extend
 \code{IUnknown}.
}
\references{
\url{http://www.omegahat.net/RDCOMClient}
\url{http://www.omegahat.net/RDCOMServer}
\url{http://www.omegahat.net/SWinTypeLibs}
}
\author{Duncan Temple Lang (duncan@r-project.org)}

\seealso{
 \code{\link{COMCreate}}
 \code{\link{.COM}}
}
%\examples{}
\keyword{interface}

