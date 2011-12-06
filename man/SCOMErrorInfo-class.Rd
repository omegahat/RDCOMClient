\name{SCOMErrorInfo-class}
\docType{class}
\alias{SCOMErrorInfo-class}
\title{S version of COM Error structure}
\description{

 This is a class that we use to represent an error that was generated
in a COM call.  If an object of this class is returned, then we know
that a call to \code{\link{.COM}} failed.  An object of this class
contains information about the error status, a description of the
error and a string giving the source of the error (e.g. the name of
the COM server).

The C++ code to use this is in place but is currently not activated.

}
\section{Objects from the Class}{
Created in the C++ code when a COM call fails.
}
\section{Slots}{
  \describe{
    \item{\code{status}:}{Object of class \code{"numeric"} the error status for the error.
     This can be used with system codes to find the symbolic name.
     See \code{.COMErrors} in the \code{\link[RDCOMServer]{}} package.}
    \item{\code{source}:}{Object of class \code{"character"} a string giving the source of the error, e.g. the application name. }
    \item{\code{description}:}{Object of class \code{"character"} the human-readable description of the error.}
  }
}

\section{Methods}{
No methods defined with class "SCOMErrorInfo" in the signature.
}

\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
}
\author{Duncan Temple Lang (duncan@wald.ucdavis.edu)}

\seealso{
 \code{\link{.COM}}
}

\examples{
}
\keyword{classes}
