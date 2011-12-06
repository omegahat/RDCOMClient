\name{getCLSID}
\alias{getCLSID}
\title{Get the UUID/GUID from the human-readable name of an application.}
\description{
 This function provides a way to find the 
 unique identifier for an application or interface
 given its human-readable form, e.g. "Excel.Application".
 This is convenient if you want to attempt to load the
 type library without having an instance of the application
 or if you want to use this information for looking
 in the Windows registry ( - see the \code{SWinRegistry} package).
}
\usage{
getCLSID(appName)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{appName}{the human-readable string (i.e. character vector of length 1)
    whose UUID we want to find.}
}
\details{
 This is a interface to the C routines
 \code{CLSIDFromString} and \code{CLSIDFromProgID}
}
\value{
 A string which gives the UUID.
 If the UUID could not be found, 
 this (currently) returns the UUID of all zeros:
 \code{"{00000000-0000-0000-0000-000000000000}"}
}
\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang (duncan@wald.ucdavis.edu)}

\seealso{

}
\examples{
  getCLSID("Excel.Application")
}
\keyword{programming}
\keyword{interface}
