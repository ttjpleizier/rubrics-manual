# script om Excel sheet te maken van een portfolio opdracht
# sheet bestaat uit twee delen
# A. invul-formulier >
# B. score-tabel > 

library(tidyverse)
library(openxlsx)
library(here)

# settings ----------------------------------------------------------------

filexl <- "cursus Themaveld Preken/toetsing/R-build-matrices/portfolio_preken.xlsx"
sheets <- openxlsx::getSheetNames(file = here::here(filexl))[-1] # weglaten tabblad instellingen
instellingen <- as_tibble(openxlsx::read.xlsx(xlsxFile = here(filexl), sheet = "instellingen"))

niveaus <- c("onvold.", "zwak", "gemiddeld", "sterk")

sheetnr <- 6
naamfull <- instellingen[[1,sheetnr+1]]

opslaanals <- paste0(sheets[sheetnr],"_beoordeling_per_student.xlsx")

# functies ----------------------------------------------------------------

# twee nieuwe functies afgeleid uit functions_assessment.R

scoretablexlsx <- function(ingevuld, # ingevuld > xlsx sheet met rubric
                           cutoff = 0.55,
                           points = c(0,1,2,3),
                           sufficient = 5.5,
                           nivs = c("--","-","+", "++"),
                           ilo.long = FALSE){
  
  niveaus <- c(paste0("niv",rep(1:4)))
  names(ingevuld)[4:7] <- niveaus
  
  # tabel scores(punten) en cijfers
  ind_aantal <- nrow(ingevuld) - sum(ingevuld$X=="0")
  max_punten <- ind_aantal*max(points)
  cutoff_matrix <- round(cutoff*max_punten,0)
  score_o <- round(seq(1,cutoff_matrix),0)
  score_v <- seq(cutoff_matrix+1,max_punten)
  cijfer_o <- round(seq(0,5.4, length.out = length(score_o)),1)
  cijfer_v <- round(seq(5.5,10, length.out = length(score_v)),1)
  score_tabel <- data.frame(score_o,cijfer_o) %>% arrange(desc(score_o))
  names(score_tabel) <- c("score","cijfer")
  score_tabel <- data.frame(score = score_v, cijfer = cijfer_v) %>% arrange(desc(score_v)) %>% bind_rows(score_tabel)
  
  return(score_tabel)
}

formxlsx <- function(ingevuld, # ingevuld = xlsx sheet met rubric
                     cutoff = 0.55,
                     points = c(0,1,2,3),
                     sufficient = 5.5,
                     nivs = c("--","-","+", "++"),
                     ilo.long = FALSE){
  
  niveaus <- c(paste0("niv",rep(1:4)))
  names(ingevuld)[4:7] <- niveaus

  formulier <- ingevuld %>% select(c(1:3)) 
  formulier[niveaus] <- ""
  formulier$Opmerkingen <- NA
  formulier$l1 <- NA
  formulier$l2 <- NA
  
  puntentellingrij <- c(rep("",2),"Puntentelling",paste("*",points,""), rep("",2))
  totalenrij <- c(rep("",2),paste0("Totalen (",round(cutoff*100,2),"% = voldoende)"),
                  rep("___",4),
                  "___", "Cijfer:", "___")
  
  
  formulier <- rbind(formulier, puntentellingrij)
  formulier <- rbind(formulier, totalenrij)
  
  # in welke rij zitten ilo's?
  formulier$rownrs <- row_number(formulier$ilo_nr)
  rownrs <- formulier[formulier$X=="0",11]$rownrs
  formulier$rownrs <- NULL
  
  return(formulier)
}



# routine -----------------------------------------------------------------

naam <- sheets[sheetnr]
gevuld <- as_tibble(openxlsx::read.xlsx(xlsxFile = here(filexl), sheet = naam))

scores <- scoretablexlsx(gevuld)
form <- formxlsx(gevuld)
formfooter <- tail(form,2)
formbody <- head(form,-2)

eersterij <- 2
laatsterij <- nrow(formbody)+1

# combine form and scores
if(nrow(form) < nrow(scores)){
  a <- nrow(scores) - nrow(form)
  for (r in 1:a) formfooter <- rbind(formfooter, NA)
} else if(nrow(scores) < nrow(form)){
  a <- nrow(form) - nrow(scores)
  for (r in 1:a) scores <- rbind(scores, NA)
}

xlsxsheet <- bind_rows(formbody,formfooter)
xlsxsheet <- bind_cols(xlsxsheet,scores)
xlsxsheet

# velden voor naam student en naam docent invoegen

#xlsxsheet[(nrow(xlsxsheet)-3),6] <- "> Naam student"
#xlsxsheet[(nrow(xlsxsheet)-2),6] <- "> Beoordelend docent"

rijfooterinfo <- nrow(form) + 2
  
xlsxsheet[rijfooterinfo+1,6] <- paste("Onderdeel",naamfull)
xlsxsheet[(rijfooterinfo+3),6] <- "> Naam student"
xlsxsheet[(rijfooterinfo+4),6] <- "> Beoordelend docent"

# ilos en indicatoren
ilos <- which(xlsxsheet$X=="0")
indicators <- which(xlsxsheet$X[2:nrow(formbody)]!="0")+2
xlsxsheet[ilos,4:7] <- ""
kolomnamen <- c(rep("",2), "ilo/indicator", niveaus, "opmerkingen", rep("",2), rep(c("score", "cijfer"),1))

# formules toevoegen voor optellen van niv1:niv4
puntenperniveau <- formfooter[1,4:7]
puntenperniveau <- as.numeric(str_extract(puntenperniveau, "\\d"))
puntenperniveau

puntenrij <- nrow(formbody)+2

# construct formula
puntenformula <- paste0("COUNTIF(",
       paste0(LETTERS[4:7],eersterij,":",LETTERS[4:7],laatsterij),
       ";",
       '"X")',
       "*",
       puntenperniveau)

class(puntenformula) <- c(class(puntenformula),"formula")

# Create a new workbook with styling options

# opmaakstijlen voor excel
style_ilo <- createStyle(valign = "top", wrapText = TRUE, textDecoration = "bold")
style_ilonniv <- createStyle(locked = TRUE)
style_indicator <- createStyle(valign = 'top', wrapText = TRUE)
style_niv <- createStyle(border = "TopBottomLeftRight", 
                         halign = "center", valign = "center",
                         textDecoration = "bold", 
                         locked = FALSE)
style_formula <- createStyle(locked = TRUE, textDecoration = "bold", halign = "center")
style_score <- createStyle(locked = TRUE, halign = "center", border = "left")
style_opm <- createStyle(valign = "top", wrapText = TRUE, fontSize = 9, locked = FALSE)
style_letop <- createStyle(valign = 'center', halign = 'center', fontColour = "red", locked = TRUE)
style_invullen <- createStyle(fontColour = "#9C0006", bgFill = "#FFFF00")
style_naam <- createStyle(fgFill = "#FFFFFF") # grijs vlak

if(TRUE){
  wb <- createWorkbook()
  
  # Add a worksheet
  addWorksheet(wb, naam)
  
  writeData(wb, 
            sheet = naam, 
            x = xlsxsheet, 
            colNames = TRUE)
  
  writeData(wb,
            sheet = naam,
            t(kolomnamen),
            startCol=1,startRow=1, colNames = FALSE)
  
  formulerij <- puntenrij+1
  
  # is invoer correct? 
  # '?' als er op de hele rij nog niets is ingevoerd
  # '!'  als er meer waarden zijn ingevoerd 
  checkinvoer <- sprintf('IF(COUNTIF(D%1$s:G%1$s,"X")>1,"!",IF(COUNTIF(D%1$s:G%1$s,"X")<1,"?",""))', indicators)
  
  # FORMULES MOETEN MET ENGELSE NAAM WORDEN WEERGEGEVEN
  writeFormula(wb, sheet = naam, x = sprintf('COUNTIF(D2:D%s,"X")*%s', laatsterij, puntenperniveau[1]), startCol = 4, startRow = formulerij)
  writeFormula(wb, sheet = naam, x = sprintf('COUNTIF(E2:E%s,"X")*%s', laatsterij, puntenperniveau[2]), startCol = 5, startRow = formulerij)
  writeFormula(wb, sheet = naam, x = sprintf('COUNTIF(F2:F%s,"X")*%s', laatsterij, puntenperniveau[3]), startCol = 6, startRow = formulerij)
  writeFormula(wb, sheet = naam, x = sprintf('COUNTIF(G2:G%s,"X")*%s', laatsterij, puntenperniveau[4]), startCol = 7, startRow = formulerij)
  
  for (i in 1:length(indicators)){
    writeFormula(wb, sheet = naam, startCol = 9, startRow = indicators[i], x = checkinvoer[i])
  }
  
  writeFormula(wb, sheet = naam, x = sprintf('SUM(D%1$s:G%1$s)', formulerij), startCol = 8, startRow = formulerij)
  writeFormula(wb, sheet = naam, x = sprintf('VLOOKUP(H%s,K2:L50,2,0)',formulerij), startCol = 10, startRow = formulerij)
  ###

  
  # set col widths: 8.43 standaard (=2.3 cm)
  setColWidths(wb, sheet = naam, 
               cols = c(1:12), widths = c(rep(5.43,2),
                                         8.43*5.3, 
                                         rep(9.43,4),
                                         8.43*4,
                                         rep(6,2),
                                         rep(6.43,2)))
  
  addStyle(wb, sheet = naam, 
           style = style_ilo, 
           cols = 3, gridExpand = TRUE, 
           rows = ilos+1)
  
  addStyle(wb, sheet = naam,
           style = style_indicator,
           cols = 3, gridExpand = TRUE,
           rows = indicators)
  
  addStyle(wb, sheet = naam,
           style = style_opm,
           cols = 8,
           rows = indicators)
  
  addStyle(wb, sheet = naam,
           style = style_niv,
           cols = 4:7, gridExpand = TRUE,
           rows = indicators)
  
  addStyle(wb, sheet = naam,
           style = style_ilonniv,
           cols = 4:7, gridExpand = TRUE,
           rows = ilos+1)
  
  addStyle(wb, sheet = naam,
           style = style_formula,
           cols = 4:10, gridExpand = TRUE,
           rows = puntenrij+1)
  
  addStyle(wb, sheet = naam, 
           style = style_score,
           cols = 11:12,
           rows = 1:(nrow(xlsxsheet)+1), gridExpand = TRUE)
  
  addStyle(wb, sheet = naam,
           style = style_letop,
           cols = 9,
           rows = indicators,
           gridExpand = TRUE)
  
  addStyle(wb, sheet = naam,
           style = style_formula,
           cols = 6,
           row = rijfooterinfo + 2)
  
  # vlak met naam student/docent
  addStyle(wb, sheet = naam,
           style = style_naam,
           cols = 5:9, gridExpand = TRUE,
           rows = (rijfooterinfo + 2):(rijfooterinfo + 6))
  
  addStyle(wb, sheet = naam,
           style = style_niv,
           cols = 8, gridExpand = TRUE,
           rows = (rijfooterinfo + 4):(rijfooterinfo + 5), 
           )
  
  conditionalFormatting(wb, sheet = naam,
                        cols = 8,
                        rows = (rijfooterinfo + 4):(rijfooterinfo + 5), 
                        rule = "==0", style = style_invullen)
  
  # alleen in te vullen cellen kunnen bewerkt worden
  protectWorksheet(wb,
                   sheet = naam,
                   protect = TRUE, #TRUE
                   password = NULL,
                   lockSelectingLockedCells = TRUE,
                   lockSelectingUnlockedCells = FALSE,
                   lockFormattingCells = FALSE)
  
  
  
  saveWorkbook(wb, paste0("gen/",opslaanals), overwrite = TRUE)
}






