\name{VARIANT-class}
\docType{class}
\alias{VARIANT-class}
\alias{CurrencyVARIANT-class}
\alias{DateVARIANT-class}
\title{Reference to a C-level VARIANT object}
\description{
 The \code{VARIANT} class provides a way to 
 represent a C-level VARIANT object in S
 and assign it to S variable names.
 The other classes are trivial extensions that provide type information about a
 variant so that we can dispatch methods on such an object.
}
\section{Objects from the Class}{
 Currently created in C code when returning values from COM calls.
 In the future, one will be able to create variant objects directly
 from R and use them as arguments in \code{\link{.COM}} calls.
 Also we will provide methods for accessing the values of the variants,
 and converting them to different S values.
}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr"} reference to the 
    C-level address of the variant data structure.}
    \item{\code{kind}:}{Object of class \code{"integer"} the kind of the variant,
   e.g. I8, I4, R8, BSTR, DISPATCH, etc. represented as a named integer 
   which has the value from the C-level enumeration.}
  }
}

\section{Methods}{
No methods defined with class "VARIANT" in the signature.
}
\references{
\url{http://www.omegahat.net/RDCOMClient}
\url{http://www.omegahat.net/RDCOMServer}
\url{http://www.omegahat.net/SWinTypeLibs}
}
\author{Duncan Temple Lang (duncan@r-project.org)}


\seealso{
 \code{\link{.COM}}
}

%\examples{}
\keyword{classes}
