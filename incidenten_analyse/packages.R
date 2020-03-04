### PAKKETTEN  ----  
list_of_packages <- c("svDialogs", "XLConnect", "readxl", "dplyr")
## https://cran.r-project.org/web/packages/XLConnect/XLConnect.pdf

### PAKKETTEN controleren ----  
 new <- unique(!list_of_packages %in% installed.packages())
 if(length(new) == 1 & new[1] == TRUE){
   new.packages <- list_of_packages[!(list_of_packages %in% installed.packages()[,"Package"])]
   if(length(list_of_packages) > 0){
     install.packages(list_of_packages)
   }
 }


### PAKKETTEN OPSTARTEN ----
library(dplyr)
library(readxl)
library(XLConnect)
library(svDialogs)
 