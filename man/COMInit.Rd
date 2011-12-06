\name{.COMInit}
\alias{.COMInit}
\title{Activate and de-activate COM facilities in R}
\description{
 This function allows one to turn on and off the
 COM facilities in an R session.
 Generally, one calls this at the start of the session
 and does not turn it off.
}
\usage{
.COMInit(status = TRUE)
}

\arguments{
  \item{status}{a logical value with \code{TRUE} meaning to
  activate the COM mechanism and \code{FALSE} meaning to turn it off.}
}
\details{
 This merely calls \code{CoInitialize}
 or \code{CoUninitialize}.
}
\value{
 \code{NULL}.
}
\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang (duncan@wald.ucdavis.edu)}

\seealso{
 \code{\link{COMCreate}}
 \code{\link{.COM}}
}

\examples{
 .COMInit()
 .COMInit(TRUE)

 .COMInit(FALSE)
}
\keyword{interface}

