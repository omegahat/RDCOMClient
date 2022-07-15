\name{COMDate-class}
\docType{class}
\alias{COMDate-class}
\alias{COMCurrency-class}
\alias{COMDecimal-class}
\alias{HResult-class}
\title{Classes for representing COM VARIANT values as numbers}
\description{
  Several types of VARIANTs in COM represent their value
  as a number which is interpreted appropriately based
  on the type of the VARIANT. For example, a date is
  the number of days since January 1, 1900,
  i.e. Midnight Jan 1, 1900 is 2.
  Similarly, a date-time value contains the hour, minute
  and second information in the decimal part of the number.
  In order to be able to associate the VARIANT type with
  the value when converting the VARIANT to R, we use
  these classes which provide the numeric value but also
  class information corresponding to the VARIANT.
}
\section{Objects from the Class}{
  Objects can be created by calls of the form \code{new("COMDate",
    ...)},
  etc. They are currently created implicitly in the C code that
  performs the conversion from VARIANT objects to R.
  In the future, we will provide facilities for the reverse direction.
}
\section{Slots}{
  \describe{
    \item{\code{.Data}:}{Object of class \code{"numeric"}
     this is the numeric vector in which the value is stored.}
  }
}
\section{Extends}{
Class \code{"numeric"}, from data part.
Class \code{"vector"}, by class "numeric".
}
\section{Methods}{
No methods defined with class "COMDate" in the signature.
}

\references{
\url{http://www.omegahat.net/RDCOMClient}
\url{http://www.omegahat.net/RDCOMServer}
\url{http://www.omegahat.net/SWinTypeLibs}
\url{http://www.omegahat.net/SWinRegistry}
}
\author{Duncan Temple Lang (duncan@r-project.org)}

\seealso{
  \code{\link{.COM}}
}

%\examples{}
\keyword{classes}
