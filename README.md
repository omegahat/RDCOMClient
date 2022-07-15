# RDCOMClient

This R package allows users to invoke (D)COM methods and access properties in any (D)COM object that implements the IDispatch interface. This includes applications such as MS Excel, Word, PowerPoint, Outlook and many others.

This is a Windows-specific package.

## Install

Run the following to install from the Omegahat repository

```r
install.packages("RDCOMClient", repos = "http://www.omegahat.net/R", type = "win.binary")
```

Please ask for  a binary if there is not one available.
Alternatively, you can build the package from source without much difficulty.
With Rtools installed, one can build the package from source without the need for
any third-party libraries.

From the (Windows) command line

```
R CMD INSTALL RDCOMClient
```

or using devtools in R

```r
devtools::install_github("omegahat/RDCOMClient")
```

## Documentation

Check the [full documentation](http://www.omegahat.net/RDCOMClient/) to learn more about the package.


## Related Packages

+ RDCOMServer
+ RDCOMEvents
+ SWinTypeLibs
+ SWinRegistry

+ rdcom
+ RExcel