\name{writeErrors}
\alias{writeErrors}
\title{Query or set whether DCOM errors are written to a file}
\description{
 This function allows one to query or set the flag that controls
whether DCOM errors are appended to a file (\code{\%TMP\%/RDCOM.err}) when
they are encountered. 
}
\usage{
writeErrors(val = logical())
}
\arguments{
\item{val}{a logical value.  If this is not specified, the function
 queries the current logical value controlling whether DCOM errors are 
 written to a file.  If this is a logical value (not an empty vector),
 that value of the first element is used to set whether the DCOM errors
 are written to a file or not.}
}
\value{
 A logical value (i.e. vector of length 1)
 indicating the value of the flag before the function
  was called. In other words, if we are setting the value,
  we get back the previous value.
}
\author{Duncan Temple Lang}
\examples{
  old = writeErrors()
  writeErrors(FALSE)
  writeErrors()
  writeErrors(old)
}
\keyword{IO}
\keyword{programming}
