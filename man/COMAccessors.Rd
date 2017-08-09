\name{COMAccessors}
\docType{methods}
\title{COM Object Accesors}
\alias{COMAccessors}
\alias{$,COMIDispatch-method}
\alias{$<-,COMIDispatch-method}
\alias{[[,COMIDispatch,ANY-method}
\alias{[[,COMIDispatch,numeric-method}
\alias{[[<-,COMIDispatch-method}
% \alias{[[,COMTypedList,ANY-method} - in COMTypedList-class.Rd
\description{
 These operators provide a more S-like
 syntax for accessing methods and properties
 in a dynamic COM object.
 One calls a COM object method using the \code{$}
 operator.
 The values of COM object properties are retrieved and set
 using \code{[[} and \code{[[<-}, respectively.
}
\section{Methods}{
\describe{

  \item{x = "COMIDispatch"}{the COMIDispatch object
   whose method or property is to be invoked.}
}}
%  \item{x}{the COM object reference in R}
%  \item{name}{the name of the method}
%  \item{i}{the name of the property or a numeric value in which 
%     case \code{x} is assumed to have a method named \code{Item}
%     and act as a container. In this case, the i-th element is returned.
%    This is not supported for setting an element (i.e. \code{[[<-})}
%  \item{j}{ignored}
%  \item{...}{ignored}  
%  \item{value}{the value to assign to the property.}

\value{
 Setting a property returns \code{NULL}.

 Invoking a method and getting a property value
 returns an S object representing the COM value.
 Primitive COM values are converted to the corresponding
 S objects.
 COM objects are returned as COMIDispatch objects.
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
 \code{\link{COMCreate}}

 \code{\link{COMList-class}}
 \code{\link{COMTypedList-class}}
 \code{\link{COMTypedNamedList-class}}
}

\examples{
 e <- COMCreate("Excel.Application")

  # Boolean/Logical
 e[["Visible"]]
  # Setting a value.
 e[["Visible"]] <- TRUE

  # String
 e[["Path"]]
 e[["Version"]]

  # Double
 e[["Width"]]

  # Long
 e[["SheetsInNewWorkbook"]]
 
  # Object
 books <- e[["Workbooks"]]

 books$Add()

  # Use this as a container, so can have integer indices, 1-based.
 books[[1]]
 e[["Workbooks"]][[1]]

\dontrun{
 books$Open("C:\\mySheet.xls")
}

 e$CheckSpelling("This is a spell check") # okay
 e$CheckSpelling("This is a spell chck") # error

\dontrun{ 
 e$SaveWorkspace()
}
 e$Quit()

 rm(list= c("e", "books"))
 gc()

}
\keyword{interface}
