# assessment form

assessment <- function(ingevuld,
                       cutoff = 0.55,
                       points = c(0,1,2,3),
                       sufficient = 5.5,
                       nivs = Niveaus){
  
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
    delete_part(part = "header") %>% 
    hline_top(part = "body") %>% 
    hline_bottom(part = "body") %>% 
    hline(i=2,j=1:nrow(score_tabel)+1) %>% 
    colformat_double(i=1,j=1:nrow(score_tabel), digits=0) %>%
  colformat_double(i=3,j=1:nrow(score_tabel), digits=0) %>% 
  fit_to_width(max_width = 8)
  
  st_t
  
  st_t_ratio <- flextable_dim(st_t)$aspect_ratio
  
  save_as_image(st_t, path = here(output,"score-cijfer.png"))
  
  # beoordelingsformulier per student
  
  formulier <- ingevuld %>% select(c(1:3)) 
  formulier[niveaus] <- ifelse(formulier$X=="0", "", "\U25A2") # square rounded corners
  formulier$Opmerkingen <- NA
  
  puntentellingrij <- c(rep("",2),"Puntentelling",paste("x",points,""))
  totalenrij <- c(rep("",2),paste0("Totalen (",round(cutoff*100,2),"% = voldoende)"),
                  rep("___",4),
                  "______ \U27A1 Cijfer: _________")

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
             "niv1" = puntentellingrij[4],
             "niv2" = puntentellingrij[5],
             "niv3" = puntentellingrij[6],
             "niv4" = puntentellingrij[7]) %>% 
    add_body(top = FALSE, 
             ilo = totalenrij[3],
             "niv1" = totalenrij[4],
             "niv2" = totalenrij[5],
             "niv3" = totalenrij[6],
             "niv4" = totalenrij[7],
             "Opmerkingen" = totalenrij[8]) %>% 
    set_header_labels("niv1" = nivs[1],
                      "niv2" = nivs[2],
                      "niv3" = nivs[3],
                      "niv4" = nivs[4]) %>% 
    
    fontsize(part = "all", size = 8) %>% 
    merge_at(i=1,j=1:3, part = "header") %>% 
    set_header_labels(ilo_nr = "leerdoel / indicator") %>% 
    width(j=1:2, width = .25) %>%
    width(j=3, width = 2) %>% 
    width(j=4:7, width = .35) %>%
    width(j=8, width = 2.3) %>% 
    align(j=4:7, align = "center")
  
  for (ilo in 1:length(rownrs)){
    form_t <- merge_at(form_t, i=rownrs[ilo], j=1:3, part = "body")
  }

  # toevoegen afbeelding met score-cijfer in footer 
  form_t <- form_t %>% 
    add_footer(ilo_nr = "test") %>% 
    merge_at(j=1:8, part = "footer") %>% 
    compose(part = "footer", 
            i = 1, j = 1, 
            value = as_paragraph(as_image(src = here(output,"score-cijfer.png"), 
                                          width = 8,
                                          height = 8*st_t_ratio)))
}