.First.lib <-
function(lib, pkg) {
 library.dynam("RDCOMClient", pkg, lib)
 .COMInit()
}

.onLoad <-
function(lib, pkg) {
 library(methods)
 library.dynam("RDCOMClient", pkg, lib)
 .COMInit()
}


#.Last.lib <-
#function() {
#  gc()
#}



