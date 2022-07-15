\name{EnumValue}
\alias{EnumValue}
\alias{EnumValue,character,numeric,EnumValue-method}
\alias{EnumValue,character,EnumValue,ANY-method}
\alias{EnumValue,numeric,EnumValue,ANY-method}
\alias{EnumValue,character,missing,EnumValue-method}
\alias{EnumValue,numeric,missing,EnumValue-method}
\alias{EnumValue,-method}
\title{Create instance of enumeration value class in R}
\description{
 This generic function is used to create an instance of a
 value from an enumeration constant definition type.
 See \code{\link{EnumValue-class}} for more information
 about these named symbolic constants.
}
\usage{
EnumValue(id, value, obj = new("EnumValue"))
}
\arguments{
  \item{id}{a string giving the symbolic name of the value.}
  \item{value}{the numeric value for this \code{EnumValue} instance.}
  \item{obj}{an instance of the class to be created.
    This allows the caller to reuse this function with 
   for different derived classes of \code{\link{EnumValue-class}}.
   One creates the new object via its default prototype and passes this to 
   \code{EnumValue}.
  }
}
\value{
 The value of the \code{obj} argument with is \code{.Data} slot
 filled in with a named integer vector of length 1.
}

\author{Duncan Temple Lang (duncan@r-project.org)}

\seealso{
 \code{\link{EnumValue-class}}
 \code{\link[methods]{setAs}}
}
%\examples{}
\keyword{programming}
\keyword{interface}


