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

Run the following to install from this repository:

```r
library(devtools)
devtools::install_github('CaliperStaff/RDCOMClient')
```

## Documentation

Check the [full documentation](http://www.omegahat.net/RDCOMClient/) to learn
more about the package.


## Related Packages

+ RDCOMServer
+ RDCOMEvents
+ SWinTypeLibs

+ rdcom
