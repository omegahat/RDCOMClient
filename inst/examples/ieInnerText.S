library(RDCOMClient)
ie = COMCreate("InternetExplorer.Application")
url. = "http://www.pricelesswarehome.org/2005/2005PL-Nominations.php"

ie$Navigate(url.)
while(ie$Busy()) Sys.sleep(1)

doc = ie[["document"]]
ie[["Visible"]] = TRUE

tabs = doc$getElementsByTagName("table")
maintab <- tabs$item(5)

tds <- maintab$getElementsByTagName("td")
print(tds$item(9)[["innerText"]])
	