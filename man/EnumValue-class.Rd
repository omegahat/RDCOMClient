\name{EnumValue-class}
\docType{class}
\alias{EnumValue-class}
\alias{show,EnumValue-method}

\title{A class to represent a value of an enumerated constant}
\description{

This class is intended to be used to represent a value that originates
from an enumerated constant.  The idea is that this value carries
around the value and the symbolic name.  Additionally, if an object of
a enumerated type is expected, an integer or character string could be
supplied and its value compared to a definition of the enumeration in
R.

It is expected that users would create new classes derived from this
one for each new enumeration type that is encountered.  The definition
of the enumerated constant values would be identified from the class
name of the new class. See \code{\link{EnumValue}}.

}
\section{Objects from the Class}{
Objects can be created by calls of the form \code{new("EnumValue", ...)}
or via the function \code{\link{EnumValue}} with an instance
of the appropriate target class specified.

}
\section{Slots}{
  \describe{
    \item{\code{.Data}:}{Object of class \code{"integer"},
       giving the value. This must have the associated name
       which makes it the symbolic constant.
    }
  }
}
\section{Extends}{
Class \code{"integer"}, from data part.
Class \code{"vector"}, by class \code{"integer"}.
Class \code{"numeric"}, by class \code{"integer"}.
}
\section{Methods}{
  \describe{
    \item{show}{\code{signature(object = "EnumValue")}: 
        this takes care of displaying the value as a regular named integer
        giving the human-readable form and its value.}
  }
}

\author{Duncan Temple Lang (duncan@r-project.org)}

\seealso{
 \code{\link{EnumValue}}
}
\examples{

   # Define a new Enumeration class. 
 setClass("Color", contains = "EnumValue")
       # Define the enumeration definition table.
 ColorEnum = c(Red = 1, Green = 2, Blue = 3)
 storage.mode(ColorEnum) = "integer"

  # Set the coercion methods to take a number or a string 
  # to this class type.
 setAs("character", "Color",
         function(from)
            EnumValue(from, new("Color")))

 setAs("numeric", "Color",
         function(from)
            EnumValue(from, new("Color")))


   # Now we can use this class.
  as(1, "Color")
  as("Red", "Color")

  as("Blue", "Color")
 
    # These should give errors, so we enclose them in a call to try. 
  try(as("Orange", "Color"))
  try(as(20, "Color"))
}
\keyword{classes}
