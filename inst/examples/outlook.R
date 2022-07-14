# From Michael Bissel (on a github issue)
library(RDCOMClient)
library(SWinTypeLibs)

if(exists("emailAddress"))
{
   out = COMCreate("Outlook.Application")
   eml = out$CreateItem(0)
   eml[["to"]] = emailAddress
   eml[["subject"]] = "RDCOMClient"
   eml[["body"]] = "This is in response to a github issue and illustrates that this is working"
   eml$Send()
}

# check the spam folder.
