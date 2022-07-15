\name{COMTypedList-class}
\alias{COMTypedParameterizedNamedList-class}
\docType{class}
\alias{COMTypedList-class}
\alias{COMTypedNamedList-class}
\alias{[[,COMTypedList,ANY-method}
\alias{[,COMTypedNamedList,character-method}
\alias{[,COMTypedNamedList,numeric-method}
\alias{[[,COMTypedList,character-method}
\alias{[[,COMTypedList,numeric-method}
\alias{[[,COMTypedNamedList,character-method}
\alias{names,COMTypedNamedList-method}

\alias{names,COMTypedParameterizedNamedList-method}

\title{Classes for representing DCOM objects that have list-like facilities}
\description{
These classes extend the basic \code{\link{COMList}}.
They are used to represent DCOM types that have methods
that support a list-like interface. 
In Office applications, this means that they have methods 
named \code{Item} and \code{Count}
which correspond to index accessors (\code{[[})
and \code{length} in R.
We can therefore make these objects feel more 
like R lists using methods defined for these classes.

Like \code{\link{COMList-class}}, these
classes are not intended to be used directly, but 
are to be used in creating derived or sub-classes.

These particular classes provide additional functionality for the
\code{\link{COMList-class}} by specifying the type of the elements.
\code{COMTypedList} knows that its elements have a type/class name
given by the singular of its own class.  For example, if we had a
class named \code{Workbooks} that extended \code{COMTypedList}, then
elements returned via its \code{Item} method would be coerced to class
\code{Workbook} the singular of \code{Workbooks}.

The class \code{COMTypedNamedList} does essentially
the same thing, but does not rely on the
class name of its elements being the singular form of
its own class name. Instead, the class name is stored
with the \code{COMTypedNamedList} object
and is used to coerce an element returned from
the \code{Item} method to the appropriate R class.
(If we had class slots in the S4 system, we would 
use that rather than putting the name into each instance
of a class. However, the overhead is small.)
This is used, for example, in the case of
our \code{Workbooks} illustration above. In Excel,
the \code{Workbooks} type behaves like a list
but returns elements which are of class \code{_Workbook}.
In this case, the \code{name} field in the
derived class \code{Workbooks} would be 
\code{_Workbook} and elements would be coerced
to that R type.


These classes form part of the run-time infrastructure
for "compiled" or pre-processed classes to DCOM interfaces
that  are generated from the type library describing a collection
of COM interfaces. See the \code{SWinTypeLibs} package
and the \code{\link[SWinTypeLibs]{generateInterface}} function.

}
\section{Objects from the Class}{
 Objects can be created from these classes manually using
the familiar \code{new} function. However, when
used in the "compiled" DCOM interfaces, they
are automatically created as return values or
arguments in methods where appropriate.
}
\section{Slots}{
  \describe{
    \item{\code{ref}:}{Object of class \code{"externalptr"}, inherited}

    \item{\code{name}:}{character string (i.e. vector of length 1).
                        This is the field in the \code{COMTypedNamedList}
                        that specifies the name of the class for an element
                        of the list.}
  }
}
\section{Extends}{
Class \code{"COMList"}, directly.
Class \code{"COMIDispatch"}, by class \code{"COMList"}.
Class \code{"IUnknown"}, by class \code{"COMList"}.
}
\section{Methods}{
  \describe{
    \item{[[}{\code{signature(x = "COMTypedList", i = "ANY")}:
               get the i-th element of the DCOM list
               and coerce the result to the appropriate type specified by the
               type of elements associated with the class.}
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
 \code{\link{COMList-class}}
 \code{\link{.COM}}
}
%\examples{}
\keyword{classes}
