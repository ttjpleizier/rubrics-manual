# feedback formulier bij afzonderlijke portfolio onderdelen
# Theo Pleizier, juli 2021


feedbackform <- function(ingevuld, #feedback-formulier 
                         filter.pattern = "",
                         no.ilos = TRUE,
                         title = "Waar sta je nu?"){
  
  stopifnot(class(filter.pattern)=="character" ,
            class(no.ilos)=="logical") 
  
  if (filter.pattern != ""){
    ind <- str_detect(ingevuld$gemiddeld, filter.pattern)
    ingevuld <- ingevuld[!ind,]
  }
  
  if(no.ilos){
    ingevuld <- filter(ingevuld, X != "0")
  } else {
    ingevuld$X[ingevuld$X=="0"] <- "." 
    ingevuld$zwak[ingevuld$X=="."] <- ingevuld$ilo[ingevuld$X=="."]  
  }
  
  rijen_ingevuld <- row_number(ingevuld$ilo_nr)
  rijen_ingevuld_ilo <<- rijen_ingevuld[ingevuld$X=="."]
  
  test <<-ingevuld
  
  
  ff <- ingevuld %>% 
    mutate(ilo_nr = NA,
           ilo = NA) %>% 
    rubrics(skip.lowest = TRUE, 
            ilo.long = FALSE,
            title = title) %>% 
    set_header_labels(ilo_nr = "", 
                      zwak = "suggesties voor verbetering",
                      gemiddeld = "criteria",
                      sterk = "goed voorbeeld") %>% 
    compose(j = c("X", "zwak", "sterk"), value = as_paragraph(""), part = "body") %>% 
    fontsize(j=5, part = "body", size = 11) %>% 
    align(j=4:6, part = "header", align = "center")
  
  if(length(rijen_ingevuld_ilo)>0){
    ff <- ff %>% 
      italic(i=rijen_ingevuld_ilo, j=4:6, part = "body") 
    
    ff <- compose(ff, 
                  i = rijen_ingevuld_ilo, 
                  j=4, 
                  part = "body", 
                  value = as_paragraph(ingevuld$zwak[ingevuld$X=="."]))
    
    for (ilo in seq_along(rijen_ingevuld_ilo)){
      ff <- ff %>% 
        merge_at(i=rijen_ingevuld_ilo[ilo], j=4:6, part = "body") %>% 
        fontsize(i=rijen_ingevuld_ilo[ilo], j=4, part = "body", size = 8)
    } 
    
  }
  
  ff <- ff %>% 
    #add_footer_lines(top = FALSE, values = "Naam student: ") %>% 
    fit_to_width(max_width = 8) 
  
  return(ff)
}