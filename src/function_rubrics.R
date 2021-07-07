# functie om rubrics te printen
# Theo Pleizier, 5 juli 2021

rubrics <- function(ingevuld,
                    nivs = c("onvoldoende", "zwak", "gemiddeld", "sterk"),
                    skip.lowest = FALSE,
                    skip.rows = 0,
                    ilo.long = FALSE,
                    title = "Beoordelingsrubrieken") {
  
  stopifnot(class(skip.lowest)=="logical",
            class(ilo.long) == "logical",
            class(skip.rows) %in% c("integer", "numeric"))
  
  ingevuld$rownrs <- row_number(ingevuld$ilo_nr)
  ingevuld <- filter(ingevuld, !rownrs %in% skip.rows)
  
  ingevuld$rownrs <- row_number(ingevuld$ilo_nr)
  rnrs <- ingevuld[ingevuld$X=="0",]$rownrs
  ingevuld$rownrs <- NULL
  
  ingevuld[ingevuld$X=="0",4:7] <- "" 
  names(ingevuld)[4:7] <- nivs
  
  aantalrow <- nrow(ingevuld)
  aantalcol <- ncol(ingevuld)
  
  if (skip.lowest){
    kolomnamen <- names(ingevuld)[-4] 
    aantalcol <- aantalcol-1}
  else {
    kolomnamen <- names(ingevuld)
  }
  
  set_flextable_defaults(fonts_ignore=TRUE)
  
  ingevuld_t <- ingevuld %>% 
    mutate(ilo_nr = ifelse(X=="0",ilo,ilo_nr)) %>% 
    flextable(col_keys = kolomnamen) %>%
    valign(j=1:aantalcol, i=1:aantalrow, valign = "top", part = "body") %>% 
    merge_at(i=1,j=1:3, part = "header") %>% 
    set_header_labels(ilo_nr = "leerdoel / indicator") %>% 
    add_header_row(top = TRUE, values = title, colwidths = aantalcol)
  
  longilo <- if_else(ilo.long,aantalcol-1,3)
  for (ilo in seq_along(rnrs)){
    ingevuld_t <- merge_at(ingevuld_t, i=rnrs[ilo], j=1:longilo, part = "body")
  }
  
  if(skip.lowest){
    breedte_niv_cols <- 2 
  } else {
    breedte_niv_cols <- 1.5
  }

  ingevuld_t <- ingevuld_t %>% 
    fontsize(part = "all", size = 9) %>% 
    width(j=1:2, width = .25) %>%
    width(j=3, width = 1.5) %>% 
    width(j=4:aantalcol, width = breedte_niv_cols) 

  return(ingevuld_t)

}
