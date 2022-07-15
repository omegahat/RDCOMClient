\name{createTypeVarName}
\alias{createTypeVarName}
\alias{createTypeVarName,COMIDispatch-method}
\alias{createTypeVarName,CompiledCOMCoClass-method}
\alias{createTypeVarName,character-method}

\alias{getCOMElements}
\alias{COMNames}


\title{Create name for specific COM type}
\description{
It is possible to creating bindings to a particular COM interface, e.g., Excel.Application
so that we can take advantage of knowledge of the set of
methods and properties and their specific IDs, rather than dynamically
looking them up in each operation.
These functions relate to creating and querying these "compiled"/generated R interfaces.

The \code{createTypeVarName} function generates the name for the R class representing this COM interface.
\code{getCOMElements} and \code{COMNames} query the elements.
}
\usage{
createTypeVarName(className, var, quote = TRUE)
}
\arguments{
  \item{className}{the COM object or "concept" for which we want an R name}
  \item{var}{character string that is combined with className}
  \item{quote}{a logical value. If \code{TRUE}, enclose the returned class name within single quotes.}
}
\value{
a character string, i.e., vector of length 1.
}
%\references{}
\author{
Duncan Temple Lang
}

%\seealso{}
%\examples{}
\keyword{interface}
\keyword{programming}
