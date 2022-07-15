\name{COMList}
\alias{COMList}
\alias{[,COMList,numeric-method}
\alias{[[<-,COMList,numeric,ANY-method}
\title{Create an instance of COMList class}

\description{
  This is a function that creates an instance
  of a \code{\link{COMList-class}} class
  or some derived class of this.
  The \code{\link{COMList-class}} class allows
  certain types of COM objects to be treated as
  an ordered container with a length, much like
  the R list concept.
}
\usage{
COMList(obj, class = "COMList")
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{obj}{the \code{\link{COMIDispatch-class}} object which is to be treated as a class.}
  \item{class}{the name of the class which is to be instantiated.
    This argument allows one to use the name of a derived class
    such as ExcelCOMList or COMTypedList.}
}

\value{
  An object of the specified class.
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
\keyword{interface}
