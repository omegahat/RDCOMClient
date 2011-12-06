library(methods)

setClass("SCOMErrorInfo", representation(status="numeric",
                                         source="character",
	                                 description="character"
	                                 ))

setClass("IUnknown", representation(ref = "externalptr"))
setClass("COMIDispatch", representation("IUnknown"))


setClass("COMDate", representation("numeric"))
setClass("COMCurrency", representation("numeric"))
setClass("COMDecimal", representation("numeric"))
setClass("HResult", representation("numeric"))



setClass("VARIANT", representation(ref= "externalptr", kind="integer"),
                    prototype=list(kind=integer(1)))
setClass("CurrencyVARIANT", representation("VARIANT"))
setClass("DateVARIANT", representation("VARIANT"))




# setClass("CurrencyValue", representation("numeric"))


