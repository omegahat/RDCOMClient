\name{CompiledCOMIDispatch-class}
\docType{class}
\alias{CompiledCOMIDispatch-class}
\alias{$,CompiledCOMIDispatch-method}
\alias{$<-,CompiledCOMIDispatch,character-method}
\alias{[[,CompiledCOMIDispatch,character-method}

\title{DCOM Interface object with knowledge of its methods and their types.}
\description{
 This class is used as a simple extension of 
 \code{\link{COMIDispatch-class}} to illustrate that 
 it has  information about its methods and properties
 and can make use of this information in coercing
 arguments and return types apporpriately.
 This class is used by "compiled" or generated
 S code that is derived from reading type libraries
 via the \code{SWinTypeLibs} package.
 
 The key things we use this class for are to 
 provide special methods for
 the \code{$} and \code{[[} operators 
 which can take advantage of the pre-computed 
 information. The class is "never" used directly
 but rather is extended to created derived
 classes for particular DCOM interfaces.

}
\section{Objects from the Class}{
  Objects of this class are not created directly, 
  but typically are implicitly created 
  by calling methods in the compiled/generated code
  created from the type library processing.
}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr"}, inherited from
          \code{\link{IUnknown-class}}}
  }
}
\section{Extends}{
Class \code{"COMIDispatch"}, directly.
Class \code{"IUnknown"}, by class \code{"COMIDispatch"}.
}
\section{Methods}{
  \describe{
    \item{\$}{\code{signature(x = "CompiledCOMIDispatch")}: 
        access a property or call a method.}
    \item{\$<-}{\code{signature(x = "CompiledCOMIDispatch", name = "character")}: ... }
    \item{[[}{\code{signature(x = "CompiledCOMIDispatch", i = "character")}: ... }
  }
}

\references{
\url{http://www.omegahat.org/RDCOMClient}
\url{http://www.omegahat.org/RDCOMServer}
\url{http://www.omegahat.org/SWinTypeLibs}
\url{http://www.omegahat.org/SWinRegistry}
}
\author{Duncan Temple Lang (duncan@wald.ucdavis.edu)}

\seealso{
 \code{\link{COMIDispatch-class}}
 \code{\link[SWinTypeLibs]{generateInterface}}
 \code{\link[SWinTypeLibs]{writeCode}}
}
\examples{

}
\keyword{classes}
