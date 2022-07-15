\name{writeErrors}
\alias{writeErrors}
\title{Query or set whether DCOM errors are written to a file}
\description{
 This function allows one to control whether DCOM diagnostic information and errors are written to a file. This is helpful for debugging.
We can now specify the path to the file to which these messages are written
and also pause/suspend and reenable logging these messages.
}
\usage{
writeErrors(val = logical())
}
\arguments{
\item{val}{the full path to a file name  in which to log messages.
  This may exist in which case messages will be apppended.
  If it doesn't exist, it will be created.
  If no value is specified, the function
 queries the current logical value controlling whether DCOM errors are 
 written to a file.  }
}
\value{
  If called with no arguments or an empty vector, 
  the return is a a logical value (i.e. vector of length 1)
  indicating the value of the boolean flag before the function
  was called.
  If called with a file name, that value is returned.
  If called with \code{TRUE}, the name of a file in R's temporary directory
  is returned and this is the file name to which the messages are written.
  If called with \code{FALSE}, the value of boolean flag
  before this call is returned.
}
\author{Duncan Temple Lang}
\examples{
  old = writeErrors()
  f = writeErrors(TRUE)
  writeErrors()
  # next, suspend writing
  writeErrors(FALSE)
  # now reenable writing
  writeErrors(f)
  readLines(f)
}
\keyword{IO}
\keyword{programming}
