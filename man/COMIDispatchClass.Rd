\name{COMIDispatch-class}
\docType{class}
\alias{COMIDispatch-class}
\alias{IUnknown-class}
\alias{$<-,COMIDispatch,ANY-method}
\alias{[[<-,COMIDispatch,character-method}
\alias{[[<-,COMIDispatch,numeric-method}
\title{Representation of generic COM object in R}
\description{ 
 These classes are used to represent in S an arbitrary
 COM object.
 \code{IUnknown} is the most basic and provides us with
 very little information about the underlying COM object.
 \code{COMIDispatch} is the work-horse for this package
 and provides a mechanism by which one
 can access the methods and properties of the COM object using the IDispatch
 interface, i.e. dynamic bindings rather than compiled ones.

 \code{COMIDispatch} is a trivial extension of \code{IUnknown} 
 that provides type information which we use to dispatch methods.
 The \code{IUnknown} class is merely a reference to the C/C++-level
 COM object.

 Reference counting is done automatically in the C code so that
 the COM object should persist as long as there is an S object
 that refers to it and will be released when no S value refers to it.
 Of course, other clients can clobber the COM object and the S references
 will be meaningless.
}
\section{Objects from the Class}{
Objects can be created by calls of the form \code{\link{COMCreate}}
or implicitly when a COM object is returned from a 
COM method  call.
}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr"} 
   this is the C++ value identifying the COM object.}
  }
}

\section{Methods}{
  \describe{
    \item{\$}{\code{signature(x = "COMIDispatch")}: 
       accessor for a method in the COM object. This returns
       a function that can be used to invoke the named method
       See \code{$} in \code{\link{COMAccessors}} }
    \item{\$<-}{\code{signature(x = "COMIDispatch", "ANY")}: 
     generates an error as one cannot assign to a function/method.
     This is implemeted this way for symmetry so that 
     assigning to a property (\code{x[["foo"]] <- 1})
     has the same basic syntax as accessing  it
     \code{x[["foo"]]}}.
     If we made this operator a property assignment, 
     we would have \code{x$foo <- 1} work but \code{x$foo} 
     failing.
     See \code{\link{COMAccessors}} 
    \item{[[}{\code{signature(x = "COMIDispatch", "numeric")}: 
     access an element of a list via the DCOM object's \code{Item} method,
     assuming we it has one.
     See \code{\link{COMAccessors}} }
    \item{[[<-}{\code{signature(x = "COMIDispatch", "character")}: 
     set a property value using the specified name.
     See \code{\link{COMAccessors}} }
  }
}

\references{
\url{http://www.omegahat.net/RDCOMClient}
\url{http://www.omegahat.net/RDCOMServer}
\url{http://www.omegahat.net/SWinTypeLibs}
\url{http://www.omegahat.net/SWinRegistry}
}
\author{Duncan Temple Lang (duncan@r-project.org)}

\seealso{
 \code{\link{COMCreate}}
}
\examples{

}
\keyword{classes}
