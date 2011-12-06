setClass("Color", contains = "EnumValue")

ColorEnum = c(Red = 1, Green = 2, Blue = 3)
storage.mode(ColorEnum) = "integer"

setAs("character", "Color",
         function(from)
            EnumValue(from, new("Color")))

setAs("numeric", "Color",
         function(from)
            EnumValue(from, new("Color")))
