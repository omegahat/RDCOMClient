\name{COMStop}
\alias{COMStop}
\title{General Error function for COM errors}
\description{
  This is a general function for raising an exception
  in a DCOM computation from within R.
  It takes a message and a status  code (an integer from the DCOM
  operation) and raises an error.
  The class of the error can be specified allowing for
  different types of errors to be distinguished.
  It is the ability to specify the class of the error
  that makes this general.
}
\usage{
COMStop(msg, status, class)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{msg}{a message string to display in the error, intended to be
    human readable.}
  \item{status}{an integer that identifies the type of the DCOM error.
    These are interpreted relative to a table of errors.
    We should make this available from the R language
    by accessing the C structures.
  }
  \item{class}{a character vector giving the name of the class(es) for
    the error being created.
    One can specify different class names to allow error handlers
    to easily differentiate between types of errors using
     \code{\link[base]{tryCatch}}.}
}
\details{
  This uses the exception mechanism in R to provide extensible
  error objects that can be caught in a flexible manner.
}
\value{
 An error object of the class specified by \code{class}.
}
\references{
  http://www.omegahat.org/RDCOMServer
  http://www.omegahat.org/RDCOMClient
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}

\seealso{
 \code{\link[RDCOMClient]{.COM}}
}
\examples{
\dontrun{
   COMStop("A fake error message", 1, "FakeCOMError")
}
}
\keyword{programming}

