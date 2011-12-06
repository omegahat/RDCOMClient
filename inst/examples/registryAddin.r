library(SWinRegistry)
library(SWinTypeLibs)
library(RDCOMServer)
library(RDCOMEvents)
library(RDCOMClient)

# Type Library for COMAddin

lib = LoadTypeLib("C:/Program Files/Common Files/DESIGNER/MSADDNDR.DLL")


register = function()
{
  # Path where Addin is

  path = c("Software", "Microsoft", "Office", "Excel", "Addins", "RFirstAddin")
  top = "HKEY_CURRENT_USER"

  # Create the registry folder, because the Recursive option doesn't work yet

  sapply(1:length(path), function(i) createRegistryKey(path[1:i], top = top))

  # The 5 possible registry values are "FriendlyName", "Description", "LoadBehavior",
  # "SatelliteDllName", and "CommandLineSafe". We only care about the first 3

  # FriendlyName: Name shown in COM Addins list

  setRegistryValue(path,
                   key = "FriendlyName",  value = "RFirstAddin", top = top)

  # Description: A more detailed description of the Addin

  setRegistryValue(path,
                   key = "Description",  value = "A first attempt at an Addin",
                   top = top)

  setRegistryValue(path,
                   key = "LoadBehavior",  value = 8, top = top)

  setRegistryValue(path, key = "CommandLineSafe", value = FALSE, top = top)

  def = createCOMServerInfo(lib[["_IDTExtensibility2"]],
                            complete = TRUE,
                            name = "RFirstAddin",
                            help = "Test Add-in for Excel implemented in R")

  def@methods$OnConnection = function(...) {cat('In OnConnection', file = 'c:\\addin.out')}
  def@methods$OnDisconnection = function(...) {cat('In OnDisconnection', file = 'c:\\addin.out')}
  def@methods$OnBeginShutdown = function(...) {cat('In OnBeginShutdown', file = 'c:\\addin.out')}
  def@methods$OnStartupComplete = function(...) {cat('In OnStartupComplete', file = 'c:\\addin.out')}
  def@methods$OnAddInsUpdate = function(...) {cat('In OnAddInsUpdate', file = 'c:\\addin.out')}


  def@classId = getuuid("a8ddb932-7290-4077-05ae-b9a58db7a97f")

  invisible(registerCOMClass(def, profile = "library(RDCOMEvents)"))
}


activate = function()
{
  app = COMCreate("Excel.Application")
  app$Workbooks()$Add()
  app[["Visible"]] = TRUE

  comAddins = app$COMAddins()

  rFirstAddin = comAddins[[1]]

  cat("Name of addin: ", rFirstAddin$ProgID(), "\n")

  cat("Connected: ", rFirstAddin$Connect(), "\n")

  print("Now to connect it: ")

  #Error: Exception Occured
  try(rFirstAddin[['Connect']] <- TRUE)

  cat("Connected: ", rFirstAddin$Connect(), "\n")

  invisible(list(addin = rFirstAddin, e = app))
}
