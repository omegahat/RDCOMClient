\name{asCOMArray}
\alias{asCOMArray}
\title{Create COM SAFEARRAY from R matrix}
\description{

 This creates a COM array from the contents of a two-dimensional R
matrix or data frame.  This can be used within R to explicitly coerce
an R object before it is passed to the COM mechanism as an argument or
return value of a function.  Otherwise, the automatic converter
mechanism creates a dynamic COM object that points to this S object
which loses no information (e.g. column or row names, etc.)
This currently only handles basic types in S, i.e. integer, numeric, 
logical and character elements.
In the future, we will provide functions for creating
an arbitrary SAFEARRAY and populating it in steps
directly from R.
}
\usage{
asCOMArray(obj)
}
\arguments{
  \item{obj}{ an S object that is converted to a matrix
   and whose resulting contents are copied into the 
   COM array.
  }
}
\value{
 An object which is an external pointer to the 
 newly created SAFEARRAY.
}

\references{
\url{http://www.omegahat.net/RDCOMClient}
\url{http://www.omegahat.net/RDCOMServer}
\url{http://www.omegahat.net/SWinTypeLibs}
}
\author{Duncan Temple Lang (duncan@r-project.org)}


\seealso{
 \code{\link{.COM}}
 \code{\link[RDCOMServer]{createCOMObject}}
}
\examples{
\dontrun{
  r = sheet$Range("A1:C10")
  r[["Value"]] <- asCOMArray(matrix(rnorm(30, 10, 3)))
}
}
\keyword{interface}
\keyword{programming}

