# maak beoordelingsmatrixen op basis van een ingevuld .csv format

# stappen (zie code hieronder):
# 1 inlezen van .csv met cursusdoelen, indicatoren per cursusdoel en per (portfolio) onderdeel
# 2 samenstellen van lege beoordelingsmatrixen in Excel met tabbladen per (portfolio) onderdeel

# ouput: porfolio-TITEL.xlsx, te gebruiken in portfolio-manual.Rmd voor het genereren van een complete handleiding met rubrieken en beoordelignsformulieren


# libraries en setings
library(here)
library(dplyr)
library(flextable)
library(openxlsx)

filepad <- "cursus\ Themaveld\ Preken/toetsing/R-build-matrices/"
input <- "src/"
output <- "gen/"
matrixcsv <- "ilo_matrix_full.csv"

# 1 inlezen .csv
ilo_matrix_raw <- read.csv(paste0(filepad,matrixcsv), encoding = "UTF-8")
leerdoelen <- ilo_matrix_raw[ilo_matrix_raw$ilo != "",c(1:2)]
leerdoelen$X <- "0"

t <- data.frame(ilo_matrix_raw[,4:9] == "")
index <- rowSums(t) != 6 # alle zes kolommen leeg, dus geen indicatoren ingevuld
indicatoren <- ilo_matrix_raw[index & ilo_matrix_raw$ilo_nr>0,]
indicatoren <- indicatoren[!sapply(indicatoren, function(x) all(x== ""))] # lege kolommen verwijderen


# 2 creëren van matrices per (portfolio)onderdeel
for (kol in seq_along(names(indicatoren))[-c(1:2)]){ #per portfolio-onderdeel
#for (kol in 5){ # test voor eerste onderdeel
  # indicatoren selecteren
  d <- indicatoren[,c(1,2,kol)]
  d <- d[d[[3]] != "" & d$ilo_nr>0,]
  d[niveaus] <- ""
  file_d <- names(d)[3]
  # leerdoelen toevoegen
  names(d)[3] <- "ilo"
  d <- leerdoelen %>% 
    filter(ilo_nr %in% d$ilo_nr) %>% 
    bind_rows(d) %>% 
    arrange(ilo_nr,X) %>% 
    select(ilo_nr,X,ilo,c(4:7))
  d[niveaus] <- if_else(d$X==0,"X", "")
  assign(paste0("portfolio_",file_d), d, envir = .GlobalEnv)
  # export naar file
  #write.csv(d, file = paste0(filepad,"matrix-",file_d,".csv")) #wegschrijven naar excel met meerdere sheets
}

matrices <- paste0("portfolio_",names(indicatoren)[-c(1:2)])

# schrijf matrix-ontwerpen naar Excel met meerdere werkbladen (library openxlsx)

werkbladen <- vector(mode = "list", length = 0)

# maak list met alle werkbladen
for (w in seq_along(matrices)){
  werkbladen <- append(werkbladen, list(get(matrices[w])))  
}
names(werkbladen) <- matrices

#write.xlsx(werkbladen, file = paste0(filepad,"portfolio_matrices_leeg.xlsx"), asTable = TRUE)

## Create a new workbook
style_large <- createStyle(valign = "top", wrapText = TRUE) # wrap text / vertically top
style_ilo <- createStyle(valign = "top", wrapText = TRUE, textDecoration = "bold")
wb <- createWorkbook()

for (w in seq_along(werkbladen)){
  ## Add a worksheet
  ilos <- which(werkbladen[[w]]$X=="0")
  addWorksheet(wb, matrices[w])
  writeData(wb, sheet = w, 
            x = werkbladen[[w]], 
            colNames = TRUE)
  ## set col widths: 8.43 standaard (=2.3 cm)
  setColWidths(wb, sheet = w, 
               cols = c(1:7), widths = c(rep(8.43,2),
                                              8.43*5.3, 
                                              rep(8.43*3.3,4)))
  addStyle(wb, sheet = w, 
           style = style_large, 
           cols = 3:7, gridExpand = TRUE, rows = 1:100)
  addStyle(wb, sheet = w, 
           style = style_ilo, 
           cols = 3, gridExpand = TRUE, 
           rows = ilos+1)
  }

## Save workbook
if (TRUE) {
  saveWorkbook(wb, paste0(filepad,output,"portfolio_matrices_leeg.xlsx"), overwrite = TRUE)
}



