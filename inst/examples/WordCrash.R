library(RDCOMClient)
app = COMCreate("Word.Application")
doc = app$Documents()$Add()
range = doc$Range()

# It is the call to Collapse() that causes the problems.
# Without this, we get back a \r.
# Seems fine now - October, 2007
#range$Collapse()
#range$Text()

