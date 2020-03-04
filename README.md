### README bestand voor Incidenten analyse
201903, HFvanEssen

De aangemaakte scripts zijn de eerste versie van de master

# Incidenten kwartaal exporteren

Direct na het kwartaal worden alle incidenten uit SymPathy in desbetreffende kwartaal gezet.
Zoekvraag: DynamicSQL > Overzichten > Overzicht incidenten>alleen selectie op de datum, overige velden leeg

# Toevoegen van gegevens aan excel bestand

Vul het aantal indicenten per kwartaal in op het 'Totaal' tabblad (bv. 'Totaal 2018')
Indien handmatige analyse is gewenst dan kunnen incidenten gesorteerd of geselecteerd worden met de filter functie.
Aantallen per incident code moeten op het totaal blad worden ingevuld en worden automatisch overgenomen in de unit tabbladen.

# Voor een automatische analyse voer het onderstaande uit.

1. Open het bestand 'start_incidenten_analyse.R' mbv het programma RStudio.
2. Selecteer alle onderstaande tekst in het bestand (CTRL + A)


```
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
incidenten.analyse()

```

3. Druk op CTRL + Enter om de analyse te starten.
4. Volg de pop-up kaders die je verder op weg helpen.


