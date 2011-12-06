\name{COMList-class}
\docType{class}
\alias{COMList-class}
\alias{[[,COMList,numeric-method}
\alias{[[<-,COMList,numeric-method}
\alias{lapply,COMList-method}
\alias{length,COMList-method}
\alias{sapply,COMList-method}

\title{COMList  collection types}
\description{
  The COMList type is used in R to represent a COM type
  which has both a Count() and Item() method.
  Such a type arises frequentyly in Office applications
  such as Excel and Word for representing collections
  or orderd lists of COM objects.  For example,
  Workbooks is a list of Workbook objects,
  Worksheets is a collection of Worksheet objects,
  and Addins is a collection of Addin objects.
  This class provides a way to treat such a container
  as an R list with methods to compute the length,
  access elements by index, and iterate over the elements
  using the s/lapply functions. For each of these, we never
  convert the list to an R list, but perform the computations
  via the COM methods.

  The \code{\link{COMTypedList-class}} class is an extension of COMList
  and should be considered a virtual class.
  (It is not for implementation reasons only.)
  This class has the property that when one extracts individual
  elements from the container in R, the class of that 
  resulting R object is determined from the class of
  the COMTypedList.  This is done in a very simple fashion
  by translating the name of the COMTypedList to its singular
  form (in English). So for a COMTypedList of class
  \code{Workbooks}, say, the expression  \code{x[[1]]}
  would result in an object of class \code{Workbook}.
  It does this computation dynamically. An extension of
  the class could compute this value at the time of definition
  and use that explicitly.  This is an example of the
  utility for class slots.

  The COMTypedList should be treated as virtual.
  One should define an extension of this and the
  associated class for the elements in tandem,
  e.g. \code{setClass("Workbooks", contains = "COMTypedList")}
  and \code{setClass("Workbook", contains = "COMIDispatch")}.
}
\section{Objects from the Class}{
  The constructor function   \code{COMList}
  should be used to create objects of either of these types.
  Alternatively, one can use the canonical form
  \code{new("COMList", comObject)}.

}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr"}. This is
      inherited from the base class \code{\link{COMIDispatch-class}}.}
  }
}
\section{Extends}{
Class \code{"COMIDispatch"}, directly.
Class \code{"IUnknown"}, by class \code{"COMIDispatch"}.
}
\section{Methods}{
  \describe{
    \item{[[}{\code{signature(x = "COMList", i = "numeric")}:
      this is the method to access an individual element in the
      COM container.  This amounts to a call to the method \code{Item}
      in the COM object.}
    \item{[[<-}{\code{signature(x = "COMList", i = "numeric")}:
      this sets the value of an individual element in the
      COM container. In general, this does nothing except
      return the \code{COMList} object so that we can 
      make inline'd assigments of the form
       \code{docs[[1]][["Range"]][["Text"]] = "Some text"}
      and have the \code{docs[[1]]} perform correctly.
    }
    \item{lapply}{\code{signature(X = "COMList")}:
      this iterates over the elements of the list
      and invokes the specified function with that element
      as the first argument.  See \code{\link[base]{lapply}}}
    \item{sapply}{\code{signature(X = "COMList")}:
      a version of \code{\link[base]{sapply}}
      that we currently need to copy to here.
    }    
    \item{length}{\code{signature(x = "COMList")}:
      computes the number of elements currently in the container.}
  }
}
\references{
  Excel types.
  \url{http://www.omegahat.org/RDCOMClient}
}
\author{Duncan Temple Lang <duncan@wald.ucdavis.edu>}


\seealso{
  \code{\link{COMList}}  
  \code{\link{.COM}}
  \code{\link{COMTypedList-class}}  
  \code{\link{COMTypedNamedList-class}}  
}
\examples{
\dontrun{
  e = COMCreate("Excel.Application")
  e$Workbooks()$Add()
  e$Workbooks()$Add()

  l = COMList(e$Workbooks())
  length(l)  # should be 2
  l[[1]]  # First Workbook

  setClass("Workbooks", contains = "COMTypedList")
  setClass("Workbook", contains = "COMIDispatch")

  l = COMList(e$Workbooks, "Workbooks")
  l[[1]]  # class is "Workbook"
}
}
\keyword{classes}
