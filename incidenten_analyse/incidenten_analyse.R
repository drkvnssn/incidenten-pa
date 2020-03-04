# ANALYSE PIPELINE
incidenten.analyse <- function(bestand = NULL, print = TRUE, warnings = FALSE, test = FALSE){
  if(warnings == FALSE){
    warning(options(warn = -1))
  }
  library(readxl)
  library(XLConnect)
  library(svDialogs)
  library(crayon)

  ### INGEVEN JAAR / KWARTAAL GEGEVENS ----
  if(is.null(bestand)){
    bestand <- file.choose() # verplaatsen naar start bestand voor gebruiker.
  }
  bestand.backup(bestand)
  kwartaal <- vraag.kwartaal()
  
  if(print == TRUE){
    cat("Het bestand", basename(bestand), "is gekozen voor analyse.\n")
    if(length(kwartaal) == 1){
      cat("Kwartaal", kwartaal, "is gekozen voor analyse.\n")
    } else {
      cat("Kwartalen", paste0(kwartaal, sep=","), "zijn gekozen voor analyse.\n")
    }
  }
  
  # inlezen incidenten tab
      data.incident <- uitlezen.incidenten(bestand = bestand, 
                                           kwartaal = kwartaal)  
      if(print == TRUE){
        cat("De incidenten van", kwartaal, "zijn ingelezen.\n")
      }
      
      # inlezen totaal blad
      data.totaal <- uitlezen.totalen(bestand = bestand) 
      if(print == TRUE){
        cat("Het totaal tabblad voor de analyse is ingelezen.\n")
      }
  
      #### met maken van een tabel voorl opslaag van het aantal incidenten per code ----
      som.incident <- som.incidenten(incident = data.incident,
                                     totaal = data.totaal)
      if(print == TRUE){
        cat("De incidenten codes zijn geteld.\n")
      }
      
      ### Kwartaal en totaal data wegschrijven ----
      schrijf.data(bestand = bestand, 
                   kwartaal = kwartaal, 
                   totaal = data.totaal,
                   incident =  som.incident, 
                   print = print,
                   test = test)
}
#### EINDE ----
