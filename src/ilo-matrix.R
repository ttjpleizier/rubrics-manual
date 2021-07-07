# maak beoordelingsmatrixen op basis van een ingevuld .csv format

# stappen (zie code hieronder):
# 1 inlezen van .csv met cursusdoelen, indicatoren per cursusdoel en per (portfolio) onderdeel
# 2 samenstellen van lege beoordelingsmatrixen in Excel met tabbladen per (portfolio) onderdeel

# de in stap 2 gegenereerde Excel gebruiken om de matrix verder in te vullen per niveau

# 3 inlezen van de ingevulde matrices in R en complete beoordelingsmatrices maken per sheet
# 4 score-cijfer tabel samenstellen
# 5 formulier voor invullen van individuele student gebruikt kan worden

# uitbouwen: bouw de code om naar functies


# libraries en setings
library(here)
library(dplyr)
library(flextable)
library(openxlsx)

filepad <- "cursus\ Themaveld\ Preken/toetsing/R-build-matrices/"
input <- "src/"
output <- "gen/"
matrixcsv <- "ilo_matrix_full.csv"

# instellingen score-tabel 
#niveaus <- c("onvoldoende","voldoende", "goed", "zeer goed") # 5.5 + 1.5 + 1.5 + 1.5
niveaus <- c("onvoldoende", "zwak", "gemiddeld", "sterk")
#niveaus <- c("--","-","+","++")

punten <- c(0,1,2,3) # toekennen punten per niveau
cutoff <- 0.55 #percentage 
sufficient <- 5.5


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



# 3 inlezen van de ingevulde matrices

#filexl <- "beoordelings-matrix.xlsx"
filexl <- "portfolio_matrices_ingevuld.xlsx"
sheets <- readxl::excel_sheets(path = paste0(filepad,filexl))

sheetnr <- 1

ingevuld <- readxl::read_xlsx(path = paste0(filepad,filexl), sheet = sheets[sheetnr])
names(ingevuld)[4:7] <- niveaus

## maken tabel met beoordelingsmatrix

# in welke rij zitten ilo's?
ingevuld$rownrs <- row_number(ingevuld$ilo_nr)
rownrs <- ingevuld[ingevuld$X=="0",]$rownrs
ingevuld$rownrs <- NULL
aantalrow <- nrow(ingevuld)

ingevuld_t <- ingevuld %>% 
  mutate(ilo_nr = ifelse(X=="0",ilo,ilo_nr)) %>% 
  flextable() %>% 
  valign(j=1:7, i=1:aantalrow, valign = "top", part = "body") %>% 
  merge_at(i=1,j=1:3, part = "header") %>% 
  set_header_labels(ilo_nr = "leerdoel / indicator") %>% 
  add_header_row(top = TRUE, values = paste0(sheets[sheetnr]), colwidths = 7)
  

for (ilo in 1:length(rownrs)){
  ingevuld_t <- merge_at(ingevuld_t, i=rownrs[ilo], j=1:3, part = "body")
}
  
ingevuld_t <- ingevuld_t %>% 
  fontsize(part = "all", size = 8) %>% 
  width(j=1:2, width = .25) %>%
  width(j=3, width = 1.5) %>% 
  width(j=4:7, width = 1.25)

ingevuld_t %>% 
  line_spacing(i=1:aantalrow, j=1:7, part = "body", space = 0.75) %>% 
  save_as_docx(path = paste0(filepad,output,"matrix_ingevuld.docx"))

ingevuld_t %>% # wordt te klein: twee aparte bestanden maken ilo 1 en ilo 2+4
  line_spacing(i=1:aantalrow, j=1:7, part = "body", space = 1) %>% 
  save_as_image(path = paste0(filepad,output,"matrix_ingevuld_1.png"))

# 4 score-cijfer-tabel
ind_aantal <- nrow(ingevuld) - sum(is.na(ingevuld$X))
max_punten <- ind_aantal*max(punten)
cutoff_matrix <- round(cutoff*max_punten,0)
score_o <- round(seq(1,cutoff_matrix),0)
score_v <- seq(cutoff_matrix+1,max_punten)
cijfer_o <- round(seq(0,5.4, length.out = length(score_o)),1)
cijfer_v <- round(seq(5.5,10, length.out = length(score_v)),1)
n <- max(length(score_o),length(score_v))
length(score_o) <- n; length(cijfer_o) <- n
length(score_v) <- n; length(cijfer_v) <- n
score_tabel <- data.frame(score_o,cijfer_o) %>% arrange(desc(score_o)) 
score_tabel <- data.frame(score_v, cijfer_v) %>% arrange(desc(score_v)) %>% bind_cols(score_tabel)
st <- as.data.frame(t(score_tabel))

st_t <- st %>% 
  tibble::rownames_to_column("score") %>%  
  mutate(score = c("Score voldoende", "Cijfer", "Score onvoldoende", "Cijfer")) %>% 
  flextable() %>% 
  #fontsize(part = "all", size = 4) %>% 
  delete_part(part = "header") %>% 
  hline_top(part = "body") %>% 
  hline_bottom(part = "body") %>% 
  hline(i=2,j=1:nrow(score_tabel)+1) %>% 
  colformat_double(i=1,j=1:nrow(score_tabel), digits=0) %>%
  colformat_double(i=3,j=1:nrow(score_tabel), digits=0) %>% 
  fit_to_width(max_width = 8)
  
st_t

st_t_ratio <- flextable_dim(st_t)$aspect_ratio

#save_as_docx(st_t, path = paste0(filepad,"score-cijfer.docx"))
save_as_image(st_t, path = paste0(filepad,output,"score-cijfer.png"))

# 5 formulier beoordeling per student

formulier <- ingevuld %>% select(c(1:3)) 
formulier[niveaus] <- ifelse(is.na(formulier$X), "", "\U25A2") # square rounded corners
formulier$Opmerkingen <- NA

puntentellingrij <- c(rep("",2),"Puntentelling",paste("x",punten),"")
totalenrij <- c(rep("",2),paste0("Totalen (",round(cutoff*100,2),"% = voldoende)"),rep("___",4),"______ \U27A1 Cijfer: _________")

# in welke rij zitten ilo's?
formulier$rownrs <- row_number(formulier$ilo_nr)
rownrs <- formulier[formulier$X=="0",9]$rownrs
formulier$rownrs <- NULL

form_t <- formulier %>% 
  mutate(Opmerkingen  = ifelse(is.na(X),"","_____________________________________")) %>% 
  mutate(ilo_nr = ifelse(X=="0",ilo,ilo_nr)) %>% 
  flextable() %>% 
  add_body(top = FALSE, 
           ilo = puntentellingrij[3],
           "onvoldoende" = puntentellingrij[4],
           "zwak" = puntentellingrij[5],
           "gemiddeld" = puntentellingrij[6],
           "sterk" = puntentellingrij[7]) %>% 
  add_body(top = FALSE, 
           ilo = totalenrij[3],
           "onvoldoende" = totalenrij[4],
           "zwak" = totalenrij[5],
           "gemiddeld" = totalenrij[6],
           "sterk" = totalenrij[7],
           "Opmerkingen" = totalenrij[8]) %>% 
  fontsize(part = "all", size = 8) %>% 
  merge_at(i=1,j=1:3, part = "header") %>% 
  set_header_labels(ilo_nr = "leerdoel / indicator") %>% 
  width(j=1:2, width = .25) %>%
  width(j=3, width = 2) %>% 
  width(j=4:7, width = .35) %>%
  width(j=8, width = 2.3) 

for (ilo in 1:length(rownrs)){
  form_t <- merge_at(form_t, i=rownrs[ilo], j=1:3, part = "body")
  }

# add picture with score-cijfer in footer of flextable

form_t <- form_t %>% 
  add_footer(ilo_nr = "test") %>% 
  merge_at(j=1:8, part = "footer") %>% 
  compose(part = "footer", 
          i = 1, j = 1, 
          value = as_paragraph(as_image(src = paste0(filepad,output,"score-cijfer.png"), 
                                        width = 8,
                                        height = 8*st_t_ratio)))

form_t

# opslaan als Word-document of Afbeeldingsbestand
#save_as_docx(form_t, path = paste0(filepad,"formulier.docx")) # werkt minder goed in Word
save_as_image(form_t, path = paste0(filepad,"formulier.png")) # import afbeelding in Word werkt beter











