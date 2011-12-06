library(RDCOMClient)

# In the registry, there are entries for InternetExplorer.Application.
# HKEY_CLASSES_ROOT\CLSID\0002DF01 0000 0000 C000 000000000046
# and this has a typelib entry which is a GUID
# EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B

app = COMCreate("InternetExplorer.Application")
app$Navigate("http://www.r-project.org/about.html")

app$Document()$Body()$Text()
