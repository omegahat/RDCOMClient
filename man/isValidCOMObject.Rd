\name{isValidCOMObject}
\alias{isValidCOMObject}
\title{Checks whether a COM object seems to be still active/valid}
\description{
This attempts to determine whether a reference to a COM object
corresponds to an active/valid COM object or if it has been destroyed.
}
\usage{
isValidCOMObject(obj)
}
\arguments{
  \item{obj}{a COMIDispatch object}
}
%\details{}
\value{
A logical value with \code{TRUE} indicating valid.
}
%\references{}
\author{
Duncan Temple Lang
}

\seealso{
\code{\link{COMCreate}}
}
%\examples{}
\keyword{interface}

