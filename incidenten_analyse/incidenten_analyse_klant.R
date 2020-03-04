incidenten.overzicht.klant <- function(bestand = NULL, verbose = TRUE, warnings = FALSE, test = FALSE){
  if(warnings == FALSE){
    warning(options(warn = -1))
  }
  library(dplyr)
  library(XLConnect)
  library(xlsx)
 library(svDialogs)
#  library(crayon)
 # library(rJava)
  #.jinit(parameters="-Xmx2g")

  ### INGEVEN JAAR / KWARTAAL GEGEVENS ----
  if(is.null(bestand)){
    bestand <- file.choose() # verplaatsen naar start bestand voor gebruiker.
  }

  klant <- vraag.klant()
  
  wb = XLConnect::loadWorkbook(bestand)
  tabbladen <- XLConnect::getSheets(wb)
  tab.jaar <- tabbladen[grep(pattern = tolower("totaal"), x = tolower(tabbladen))]
  
  if(verbose == TRUE){
    cat("Het bestand '", basename(bestand), "' is gekozen voor analyse.\n", sep = "")
    cat("Klant '", klant, "' is gekozen voor analyse.\n\n", sep = "")
  }

  # ALLE KWARTALEN LANGSLOPEN
  q1 <- lees.kwartaal.klant(wb = wb, kwartaal = 1)
  q2 <- lees.kwartaal.klant(wb = wb, kwartaal = 2)
  q3 <- lees.kwartaal.klant(wb = wb, kwartaal = 3)
  q4 <- lees.kwartaal.klant(wb = wb, kwartaal = 4)
  q.totaal <- rbind(q1, q2, q3, q4)
  
  resultaten.klant <- tellen.incidenten.klant(q1 = q1, q2 = q2, q3 = q3, q4 = q4, klant = klant)
  resultaten.klant <- totalen.klant(x = resultaten.klant)
  
  pad <- gsub(pattern = as.character(basename(bestand)), replacement = "", x = as.character(bestand))
  pad.folder <- file.path(pad, "KLANTEN")
  if(file.exists(pad.folder) == FALSE){
    dir.create(pad.folder)
  }
  output.file <- file.path(paste0(klant,"_", tab.jaar, ".xlsx"))
  output.pad.file <- file.path(pad.folder, output.file)
    XLConnect::writeWorksheetToFile(output.pad.file, data = resultaten.klant, 
                                   sheet = "Overzicht_connect", startRow = 1, startCol = 1)
    
   # xlsx::write.xlsx2(resultaten.klant, file = output.pad.file, sheetName = "Overzicht_writeslsx2",
    #                  col.names = TRUE, row.names = FALSE, append = TRUE)
     
  cat("Analyse is klaar en data staat hier : ", output.pad.file, "\n")
  
  rm(q1,q2,q3,q4,q.totaal,resultaten,resultaten.klant)
  ## welke data moet er behouden blijven?
  ## alle kwartaal codes bij elkaar?
  ## Alle kwartaal codes los opgeslagen?
  
}
