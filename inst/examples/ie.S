library(RDCOMClient)

# Internet explorer.

a = COMCreate("InternetExplorer.application")	

a[["Visible"]] <- TRUE

# Go to a new page.
a$Navigate2("http://www.omegahat.org/")

