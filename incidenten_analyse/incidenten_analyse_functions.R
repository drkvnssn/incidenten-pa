vraag.klant <- function(){
  klant <- svDialogs::dlgInput(paste("Geef de klantnaam voor analyse. \n"))$res
  result <- toupper(klant)
  return(result)
}
lees.kwartaal.klant <- function(wb = NULL, kwartaal = NULL){
  if(class(wb)[1] != "workbook"){
    stop("Excel workbook is niet in het juiste format.\n")
  }
  if(kwartaal %in% c(1:4) == TRUE){
    kwartaal <- paste0("Q",kwartaal)
  }else{
    stop("Kwartaal niet in het juiste format.\n")
  }
         tabbladen <- XLConnect::getSheets(wb)
         q.data <- XLConnect::readWorksheet(wb, sheet = grep(pattern = kwartaal, x = tabbladen), header = TRUE)
         if(nrow(q.data) == 0){
          q.data <- as_tibble(matrix(rep(NA, 16), nrow = 1))
         }
         
           kolom.namen <- c("Verslag type","Jaar","Nr","AanvragendeAfdeling","AanvragendArts","Specialisme","incidentnr","Code","Unit",
                            "Omschrijving","Commentaar","Meldingsdatum","Door","Afhandelingsmelding","Afmeldingsdatum","Afgemeld door")
           colnames(q.data) <- kolom.namen
           q.data$AanvragendeAfdeling <- gsub(q.data$AanvragendeAfdeling, pattern = " ", replacement = "")
           q.data$Unit <- gsub(q.data$Unit, pattern = " ", replacement = "")
           q.data <- dplyr::as_tibble(q.data)
         return(q.data)
}
tellen.incidenten.klant <- function(q1, q2, q3, q4, klant = NULL){
  q.totaal <- rbind(q1, q2, q3, q4)
  q1 <- dplyr::filter(q1, AanvragendeAfdeling == klant)
  q2 <- dplyr::filter(q2, AanvragendeAfdeling == klant)
  q3 <- dplyr::filter(q3, AanvragendeAfdeling == klant)
  q4 <- dplyr::filter(q4, AanvragendeAfdeling == klant)
  q.totaal <- dplyr::filter(q.totaal, AanvragendeAfdeling == klant)
  
  # alle unieke codes voor deze klant opzoeken
  codes.klant <- q.totaal[!duplicated(q.totaal$Code),]
  codes.klant <- dplyr::select(codes.klant, AanvragendeAfdeling, Code, Unit, Omschrijving)
  
  resultaten <- dplyr::tibble(Q1 = rep(0, nrow(codes.klant)), 
                       Q2 = rep(0, nrow(codes.klant)),
                       Q3 = rep(0, nrow(codes.klant)),
                       Q4 = rep(0, nrow(codes.klant)),
                       Totaal = rep(0, nrow(codes.klant))
  )
  
  # hierna per kwartaal de telling doen voor deze klant
  for(i in 1:nrow(codes.klant)){
    if(nrow(q1) > 0){
      resultaten$Q1[i] <- sum(q1$Code %in% codes.klant$Code[i])
    }
    if(nrow(q2) > 0){
      resultaten$Q2[i] <- sum(q2$Code %in% codes.klant$Code[i])
    }
    if(nrow(q3) > 0){
      resultaten$Q3[i] <- sum(q3$Code %in% codes.klant$Code[i])
    }
    if(nrow(q4) > 0){
      resultaten$Q4[i] <- sum(q4$Code %in% codes.klant$Code[i])
    }
    if(nrow(q.totaal) > 0){
      resultaten$Totaal[i] <- sum(q.totaal$Code %in% codes.klant$Code[i])
    }
  }
  # ALLES UITZETTEN IN 1 BESTAND EN TOTALEN BEREKENEN
  resultaten.klant <- cbind(codes.klant, resultaten)
  resultaten.klant <- dplyr::arrange(resultaten.klant, Code)
  
  return(resultaten.klant)
}
totalen.klant <- function(x){
  x <- rbind(x, x[1,])
  x <- dplyr::as_tibble(x)
  x[nrow(x),2] <- NA
  x[nrow(x),3:4] <- c(NA,"TOTAAL:")
  x[nrow(x),5:9] <-  colSums(x[1:(nrow(x)-1),5:9])
  return(x)
}
bestand.backup <- function(bestand, folder = "backup", print = FALSE){
  naam <- basename(bestand)
  pad <- gsub(pattern = naam, replacement = "", x = bestand)
  datum.tijd <- mgsub(pattern = c("-"," ",":"), replacement = c("","_","-"),x = as.character(Sys.time()))
  datum.tijd <- substr(x = datum.tijd, start = 1, stop = 11)
  
  backup.folder <- file.path(pad, folder)
  if(!file.exists(backup.folder)){
    dir.create(backup.folder)
  }
  outputFile <- paste0(datum.tijd, "_", naam)
  outputFile <- file.path(backup.folder, outputFile)
  file.copy(from = bestand, to = outputFile, recursive = FALSE,
            copy.mode = TRUE, copy.date = FALSE)
  if(print == TRUE){
    cat("Een backup van het originele xlsx bestand is opgeslagen.\n")
  }
}
haal.jaar.kwartaal <- function(x = NULL){
  if(is.null(x) == TRUE){
    x <- Sys.Date()
  } else if (nchar(as.character(x)) != 10){
    stop("Datum format is niet goed. Het format 'YYYY-MM-DD' is nodig voor verder analyse.\n ")
  }
  datum <- strsplit(as.character(x), split = "-")[[1]][1:2]
  results <- matrix(data = NA, nrow = 2, ncol = 2,
                    dimnames = list(c("echt", "analyse"), 
                                    c("jaar", "kwartaal")))
  results[1,1] <- datum[1]
  if(as.numeric(datum[2]) %in% 1:3){
    results[2,1] <- as.character(as.numeric(datum[1])-1)
    results[2,2] <- "Q4"
    results[1,2] <- "Q1"
  } else if(as.numeric(datum[2]) %in% 4:6){
    results[2,1] <- datum[1]
    results[2,2] <- "Q1"
    results[1,2] <- "Q2"
  } else if(as.numeric(datum[2]) %in% 7:9){
    results[2,1] <- datum[1]
    results[2,2] <- "Q2"
    results[1,2] <- "Q3"
  } else if(as.numeric(datum[2]) %in% 10:12){
    results[2,1] <- datum[1]
    results[2,2] <- "Q3"
    results[1,2] <- "Q4"
  }
  return(results)
}
controle.jaar.kwartaal <- function(){
  jaar.kwartaal <- haal.jaar.kwartaal()
  vraag <- paste("Is",jaar.kwartaal[2,2],"van", jaar.kwartaal[2,1], "het juiste jaar en kwartaal voor analyse?     (j/n)\n\n")
  controle.jaar <- dlgInput(vraag)$res
  
  if(tolower(controle.jaar) %in% c("nee","n","no", "nope", "niet") == TRUE){
    jaar.kwartaal[2,1] <- dlgInput(paste("Geef het juiste jaar voor analyse:     (YYYY)\n"))$res
    kwartaal <- dlgInput(paste("Geef het juiste kwartaal voor analyse:     (1, 2, 3, 4)  \n"))$res
    if(nchar(as.character(kwartaal)) > 1){ 
      stop("Een kwartaal wordt aangegeven met een cijfer 1, 2, 3 of 4.\n")
    } else if(kwartaal < 1 | kwartaal > 4){
      stop("Een kwartaal wordt aangegeven met een cijfer 1, 2, 3 of 4.\n")
    }
    jaar.kwartaal[2,2] <- paste0("Q", kwartaal)
  }
  result <- list(jaar = jaar.kwartaal[2,1],
                 kwartaal = jaar.kwartaal[2,2])
  return(result)
}
controle.bestand <- function(bestand, jaar, print = TRUE){
  bestand <- basename(bestand)
  extension <- unlist(strsplit(x = bestand, ".", fixed = TRUE))
  extension <- extension[length(extension)]
  if(extension %in% c("xls", "xlsx") == FALSE){
    stop("Extensie hoort niet bij Excel bestand.\n")
  }
  controle.jaar <- grep(pattern = as.character(jaar),  x = bestand)
  if(controle.jaar != 1){
    stop("Het opgegeven jaartal is niet gevonden in de titel van het bestand.\n")
  }
  if(print == TRUE){
    cat("Opgeven bestand heeft juiste extensie en analyse jaar komt overeen.\n")
  }
}
vraag.kwartaal <- function(){
  kwartaal <- dlgInput(paste("Geef het juiste kwartaal voor analyse:     1, 2, 3, of 4 \n"))$res
  if(nchar(as.character(kwartaal)) > 1){ 
    stop("Een kwartaal wordt aangegeven met een cijfer 1, 2, 3, of 4.\n")
  } else if(as.numeric(kwartaal) < 1 | as.numeric(kwartaal) > 4){
    stop("Een kwartaal wordt aangegeven met een cijfer 1, 2, 3 of 4.\n")
  }
  result <- paste0("Q", kwartaal)
  return(result)
}
uitlezen.incidenten <- function(bestand = NULL, kwartaal = NULL){
  ### CONTROLEREN AANWEZIGHEID VAN INPUT ----
  if(is.null(bestand) == TRUE){
    stop("Er is geen bestandsnaam meegegeven voor de analyse.\n")
  } 
  if(is.null(kwartaal) == TRUE){
    stop("Er is geen kwartaal meegegeven voor de analyse of het format is niet juist.\n")
  } else if(is.numeric(kwartaal)){
    kwartaal <- paste0("Q", kwartaal)
  }
  
  if(file.exists(bestand)){
    tabbladen <- XLConnect::getSheets(loadWorkbook(bestand))
    tab.incidenten <- grep(pattern = tolower(kwartaal), x = tolower(tabbladen))
    incidenten.data <- readxl::read_excel(bestand, sheet = tab.incidenten)
  } else {
    cat("Bestand is niet gevonden op onderstaande locatie.\n    ", file)
  }
  return(incidenten.data)
}
uitlezen.totalen <- function(bestand = NULL){
  ### CONTROLEREN AANWEZIGHEID VAN INPUT ----
  if(is.null(bestand) == TRUE){
    stop("Er is geen bestandsnaam meegegeven voor de analyse.\n")
  } 
  if(file.exists(bestand)){
    tabbladen <- XLConnect::getSheets(loadWorkbook(bestand))
    tab.totaal <- grep(pattern = "totaal", x = tolower(tabbladen))
    totaal.data <- readxl::read_excel(bestand, sheet = tabbladen[tab.totaal])
    totaal.data <- rbind(totaal.data[1,],totaal.data)
    totaal.data[1,] <- NA 
  } else {
    cat("Bestand is niet gevonden op onderstaande locatie.\n    ", file)
  }
  return(totaal.data)
}
mgsub <- function(pattern, replacement, x, ...) {
  if (length(pattern)!=length(replacement)) {
    stop("pattern and replacement do not have the same length.")
  }
  result <- x
  for (i in 1:length(pattern)) {
    result <- gsub(pattern[i], replacement[i], result, ...)
  }
  result
}
schrijf.data <- function(bestand, kwartaal, totaal, incident, print = TRUE, test = FALSE){
  wb <- XLConnect::loadWorkbook(filename = bestand)
  XLConnect::setStyleAction(wb, XLC$"STYLE_ACTION.NONE")
  
  kolom.totaal <- grep(pattern = "q1", x = tolower(colnames(totaal)))+4
  kolom.kwartaal <- grep(pattern = tolower(kwartaal), x = tolower(colnames(totaal)))
  
  tabbladen <- XLConnect::getSheets(wb)
  blad.totaal <- grep(pattern = "totaal", tolower(tabbladen))
  
  
  for(q in 2:nrow(incident)){
    if(is.na(incident[q,3]) == FALSE){
      XLConnect::writeWorksheet(object = wb, 
                                data = incident[q, 3], 
                                sheet = blad.totaal, 
                                startRow = q, 
                                startCol = kolom.kwartaal, 
                                header=FALSE)
    }
  }
  
  if(test == TRUE){
    outputFile <- gsub(pattern = ".xlsx", replacement = "_test.xlsx", bestand)
  } else {
    outputFile <- bestand
  }
  
  for(i in 1:length(tabbladen)){
    setForceFormulaRecalculation(wb, sheet = i, TRUE)
  }
  saveWorkbook(object = wb, file = outputFile)
  if(print == TRUE){
    cat("\n  De incidenten analyse is afgerond.\n")
    pad <- gsub(pattern = as.character(basename(outputFile)), replacement = "", x = as.character(bestand))
    cat("\n  Het bewerkte bestand en lokatie:\n   ")
    cat(blue("\n    ", basename(outputFile),"\n"))
    cat(blue("    ", pad, "\n"))
  }
}
som.incidenten <- function(incident, totaal){
  #### met maken van een tabel voor opslag van het aantal incidenten per code ----
  aantal.incidenten <- data.frame(totaal$`Incident Code`,
                                  totaal$Eenheid,
                                  as.numeric(rep(NA, nrow(totaal))),
                                  stringsAsFactors = FALSE)
  colnames(aantal.incidenten) <- c(colnames(totaal)[1:2], paste0("Totaal"))
  
  ### INCIDENTEN TELLEN ----
  for(i in 1:nrow(totaal)){ # 
    if(is.na(as.numeric(totaal$'Incident Code'[i])) == FALSE){
      incidenten <- 0
      aantal.incidenten[i,1] <- totaal$'Incident Code'[i]
      incidenten <- sum(as.numeric(incident$Code) %in% as.numeric(totaal$'Incident Code'[i]))
      aantal.incidenten$Totaal[i] <- incidenten
    } else {
      aantal.incidenten[i,2] <- "TOTAAL"
      aantal.incidenten[i,3] <- NA
    } 
  }
  return(aantal.incidenten)
}
mySum <- function(x){
  x <- unlist(x)
  x[is.na(x)] <- 0
  result <- sum(as.numeric(as.character(x), na.rm = TRUE))
  return(result)
}
som.kwartalen <- function(somIncident, totaal, kwartaal, print = TRUE){
  
  kolom.q1 <- grep(pattern = "q1", x = tolower(colnames(totaal)))
  kolom.kwartaal <- grep(pattern = tolower(kwartaal), x = tolower(colnames(totaal)))
  totaal[ ,kolom.kwartaal] <- somIncident$Totaal
  
  som.q <- apply(totaal[ ,kolom.q1:(kolom.q1+3)], mySum, MARGIN = 1)
  totaal[,(kolom.q1+4)] <- som.q
  
  if(print == TRUE){
    cat("Totalen van alle kwartalen zijn berekend.\n")
  }
  return(totaal)
}
