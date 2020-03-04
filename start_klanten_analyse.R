### inladen functies ----------
sourceDir <- function(path, trace = TRUE, ...) {
  for (nm in list.files(path, pattern = "\\.[RrSsQq]$")) {
    if(trace) cat(nm,"")           
    source(file.path(path, nm), ...)
    if(trace) cat("\n")
  }
}
script.path <- "\\\\vumc.nl/onderzoek$/s4e-gpfs1/pa-tgac-01/analisten/Dirk/r_scripts/incidenten_analyse/"
sourceDir(file.path(script.path, "scripts"))

###
incidenten.overzicht.klant()
