# RDCOMClient

This R package allows users to invoke (D)COM methods and access properties in
any (D)COM object that implements the IDispatch interface. This includes
applications such as MS Excel, Word and many others. This is a Windows-specific
package.

## Note on package origins

This package was developed by Duncan Lane, and an his version is still
available here: https://github.com/omegahat/RDCOMClient

That version stopped working with the latest R several years ago, and after
months of trying to make contact, this fork was created. This is not hosted on
CRAN, but can be installed as shown below.

## Install

### For R 3.6.x

```r
library(devtools)
devtools::install_github('dkyleward/RDCOMClient')
```

### For R 4.0

Until this package is updated for R 4.0, installing directly from GitHub will
not work. Instead, use the code snippet below to install from a binary version
already compiled:

```r
dir <- tempdir()
zip <- file.path(dir, "RDCOMClient.zip")
url <- "https://github.com/dkyleward/RDCOMClient/releases/download/v0.94/RDCOMClient_binary.zip"
download.file(url, zip)
install.packages(zip, repos = NULL, type = "win.binary")
```

## Documentation

Check the [full documentation](http://www.omegahat.net/RDCOMClient/) to learn
more about the package.


## Related Packages

+ RDCOMServer
+ RDCOMEvents
+ SWinTypeLibs

+ rdcom
