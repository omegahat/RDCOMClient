\name{COMCreate}
\alias{COMCreate}
\alias{getCOMInstance}
\title{Create a COM object}
\description{
 These functions are used from R to either create a COM object of a specific type
 or to obtain a reference to an existing COM application/object.
}
\usage{
COMCreate(name, ..., existing = TRUE)
getCOMInstance(guid, force = TRUE, silent = FALSE)
}
\arguments{
  \item{name}{the name identifying the type of COM object.
  This is usually the name of the class, such as
  Excel.Application. 
  In the future, this may support finding class IDs
  separately and passing them as this argument}
  \item{\dots}{currently ignored. In the future, 
   we will have arguments controlling how
   the server is created, e.g. the location of the machine, etc.}
  \item{guid}{a string (character vector of length 1) that identifies the 
   GUID - Global Unique Identifier - for the COM interface of interest.
   This should have the curly braces ('{...}').}
  \item{existing}{a logical value. If this is \code{TRUE}, then first we try to 
    connect to an existing instance of the specified DCOM server,
    If no such instance exists, then we create a new one.
    If this is \code{FALSE}, then we just create a new instance and don't check whether
    there is one already running. This guarantees that we get }
  \item{force}{a logical value. If this is \code{TRUE}, this will (attempt to) create a new instance
   of the specified DCOM application if there is no instance already in existence.
   In other words, \code{getCOMInstance} will look for an existing instance and if that fails,
   it will create a new one. If this is \code{FALSE}, \code{getCOMInstance} will return
   after attempting to find an existing instance. As a result, the
   return value might be \code{NULL}.}
  \item{silent}{a logical value that is used when \code{force} is
    \code{TRUE} and a new object needs to be created. If this is \code{TRUE}, the warning
    message about creating the new instance is supressed. Otherwise, the
    warning is issued.}
}
\details{
 This resolves the class ID given the name and
 then creates the instance using \code{CoCreateInstance}.
 The class ID is found using \code{CLSIDFromString}
 and if that fails using \code{CLSIDFromProgID}.
}
\value{
 An object of class \code{COMIDispatch}.
}
\note{
  \code{getCOMInstance} works with Word and Excel
  but for some reason on at least some machines fails
  to find an existing Internet Explorer application.
  As a result, it will create a new instance of the application.
}
\references{
\url{http://www.omegahat.net/RDCOMClient}
\url{http://www.omegahat.net/RDCOMServer}
\url{http://www.omegahat.net/SWinTypeLibs}
\url{http://www.omegahat.net/SWinRegistry}
}
\author{Duncan Temple Lang (duncan@r-project.org)}

\seealso{
 The \code{$} and \code{[[} operators
 for the \code{COMIDispatch} class - \code{\link{COMAccessors}}.

 Also, see how one can generate "compiled" or processed interfaces
 to any or all of the types in an application using the
 \code{SWinTypeLibs}.
}

\examples{

\dontrun{
 COMCreate("Excel.Application")
 getCOMInstance("{000208D5-0000-0000-C000-000000000046}")
}
}
\keyword{interface}

