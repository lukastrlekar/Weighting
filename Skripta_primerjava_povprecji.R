# potrebni paketi
library(haven)
library(labelled)
library(weights)
library(openxlsx)

# TODO
# pri nominalnih tabelah daj iz funkcije wtd_t_test deleže in abs razlike

# pomožne funkcije

# funkcija ki prešteje št. relativnih razlik glede na intervale in št. stat. značilnih spremenljivk
count_rel_diff <- function(vec, p_vec) {
  vec <- abs(vec)
  
  frek <- c(sum(vec > 20), sum(vec > 10 & vec <= 20), sum(vec >= 5 & vec <= 10), sum(vec < 5))
  
  p_frek <- c(sum(vec > 20 & p_vec < 0.05), sum(vec > 10 & vec <= 20 & p_vec < 0.05), sum(vec >= 5 & vec <= 10 & p_vec < 0.05), sum(vec < 5 & p_vec < 0.05))
  
  list(sums = frek,
       cumsums = cumsum(frek),
       p_sums = p_frek,
       p_cumsums = cumsum(p_frek))
}

# funkcija za utežene frekvence
weighted_table = function(x, weights) {
  use <- !is.na(x)
  x <- x[use]
  weights <- weights[use]
  weights <- weights/mean(weights, na.rm = TRUE)
  
  vapply(split(weights, x), sum, 0, na.rm = TRUE)
}

# funkcija za izračun testne statistike (Welchev t-test)
wtd_t_test <- function(x, y, weights_x = NULL, weights_y = NULL){
  n1 <- sum(!is.na(x))
  n2 <- sum(!is.na(y))
  
  if(is.null(weights_x)) weights_x <- rep(1, length(x))
  
  if(is.null(weights_y)) weights_y <- rep(1, length(y))
  
  use_x <- !is.na(x)
  use_y <- !is.na(y)
  
  x <- x[use_x]
  y <- y[use_y]
  weights_x <- weights_x[use_x]
  weights_y <- weights_y[use_y]
  weights_x <- weights_x/mean(weights_x, na.rm = TRUE)
  weights_y <- weights_y/mean(weights_y, na.rm = TRUE)
  
  mu1 <- weighted.mean(x = x, w = weights_x, na.rm = TRUE)
  mu2 <- weighted.mean(x = y, w = weights_y, na.rm = TRUE)
  
  if(is.null(weights_x) && is.null(weights_y)) {
    se1 <- var(x, na.rm = TRUE)/n1
    se2 <- var(y, na.rm = TRUE)/n2
  } else {
    u1 <- x * weights_x
    se1 <- (var(u1, na.rm = TRUE) + mu1^2 * var(weights_x, na.rm = TRUE) - 2 * mu1 * cov(u1, weights_x))/(sum(weights_x, na.rm = TRUE))

    u2 <- y * weights_y
    se2 <- (var(u2, na.rm = TRUE) + mu2^2 * var(weights_y, na.rm = TRUE) - 2 * mu2 * cov(u2, weights_y))/(sum(weights_y, na.rm = TRUE))
    
    # se1 <- (var(x, na.rm = TRUE)/n1) * (1 + var(weights_x))
    # se2 <- (var(y, na.rm = TRUE)/n2) * (1 + var(weights_y))
  }
  
  t <- (mu2 - mu1)/(sqrt(se1 + se2))
  df <- (se1 + se2)^2/(se1^2/(n1 - 1) + se2^2/(n2 - 1))
  p <- pt(q = abs(t), df = df, lower.tail = FALSE) * 2
  
  c(povp1 = mu1, povp2 = mu2, diff = mu2 - mu1, t = t, p = p, df = df)
}


# # (1)
# (1/sum(y)^2) * sum(y^2 * (x - weighted.mean(x, y))^2) 
# 
# # (2)
# y1 = y/sum(y)
# 
# sum(y1^2 * (x- weighted.mean(x, y1))^2)
# # 1 = 2
# 
# # (3)
# n1 = sum(y)
# mu1 = weighted.mean(x,y)
# u1 = x*y
# weights_x = y
# (n1 * var(u1, na.rm = TRUE) + mu1^2 * n1 * var(weights_x, na.rm = TRUE) - 2 * mu1 * n1 * cov(u1, weights_x, use = "complete.obs"))/(sum(weights_x, na.rm = TRUE)^2)



# glavna funkcija za izvoz tabel v Excel
izvoz_excel_tabel <- function(baza1 = NULL,
                              baza2 = NULL,
                              ime_baza1 = "baza 1",
                              ime_baza2 = "baza 2",
                              utezi1 = NULL,
                              utezi2 = NULL,
                              stevilske_spremenljivke = NULL,
                              nominalne_spremenljivke = NULL,
                              file) {
  
  if(!is.data.frame(baza1) || !is.data.frame(baza2)){
    stop("Baza mora biti SPSS podatkovni okvir.")
  }
  
  if(is.null(utezi1) || is.null(utezi2)){
    stop("Določite uteži!")
  }
  
  if(!is.numeric(utezi1) || !is.numeric(utezi2)){
    stop("Uteži morajo biti v obliki numeričnega vektorja.")
  }
  
  # uteži povpr ni 1
  
  if(!is.null(stevilske_spremenljivke) && !is.character(stevilske_spremenljivke)){
    stop("Podane številske spremenljivke morajo biti v obliki character vector.")
  }
  
  if(!is.null(nominalne_spremenljivke) && !is.character(nominalne_spremenljivke)){
    stop("Podane nominalne spremenljivke morajo biti v obliki character vector.")
  }
  
  wb <- createWorkbook()
  
  addWorksheet(wb = wb, sheetName = "Opozorila", gridLines = FALSE)
  
  addWorksheet(wb = wb, sheetName = "Povzetek", gridLines = FALSE)
  
  warning_counter <- FALSE
  
  if(!is.null(stevilske_spremenljivke)){
    # Številske spremenljivke -------------------------------------------------
    
    imena_st_1 <- names(baza1)[names(baza1) %in% stevilske_spremenljivke]
    imena_st_2 <- names(baza2)[names(baza2) %in% stevilske_spremenljivke]
    
    non_st_spr <- union(setdiff(imena_st_1, imena_st_2), setdiff(imena_st_2, imena_st_1))
    
    if(length(imena_st_1) != length(imena_st_2)){
      warning(paste("Imena podanih številskih spremenljivk se ne ujemajo v obeh bazah. To so spremenljivke", paste(non_st_spr, collapse = ", ")))
      
      writeData(wb = wb, sheet = "Opozorila", xy = c(1,1),
                x = paste("Imena podanih številskih spremenljivk se ne ujemajo v obeh bazah. To so spremenljivke", paste(non_st_spr, collapse = ", "), "in so bile zato odstranjene iz analiz."))
      
      warning_counter <- TRUE
    }
    
    # izberemo samo spremenljivke, ki so prisotne v 1. in 2. bazi
    stevilske_spremenljivke <- imena_st_1[imena_st_1 %in% imena_st_2]
    
    # pretvorimo spss manjkajoče vrednosti v prave manjkajoče (NA) in uredimo bazi v isti vrstni red spremenljivk
    baza1_na <- user_na_to_na(baza1[,stevilske_spremenljivke, drop = FALSE])
    baza2_na <- user_na_to_na(baza2[,stevilske_spremenljivke, drop = FALSE])
    
    # preverimo, da je isto število kategorij v obeh bazah
    levels1 <- lapply(stevilske_spremenljivke, function(x) attr(baza1_na[[x]], "labels"))
    levels2 <- lapply(stevilske_spremenljivke, function(x) attr(baza2_na[[x]], "labels"))
    
    indeksi <- sapply(seq_along(stevilske_spremenljivke), function(i) all(names(levels1[[i]]) == names(levels2[[i]])))
    
    if(!all(indeksi)){
      warning(paste("Kategorije se pri številski/h spremenljivki/ah", paste(stevilske_spremenljivke[!indeksi == TRUE], collapse = ", "), "ne ujemajo. Te spremenljivke niso bile odstranjene iz analiz, vendar svetujemo previdnost pri analizi."))
      
      writeData(wb = wb, sheet = "Opozorila", xy = c(1,2),
                x = paste("Kategorije se pri številski/h spremenljivki/ah", paste(stevilske_spremenljivke[!indeksi == TRUE], collapse = ", "), "ne ujemajo. Te spremenljivke niso bile odstranjene iz analiz, vendar svetujemo previdnost pri analizi."))
      
      warning_counter <- TRUE
    }
    
    # tabela
    tabela_st <- data.frame("Spremenljivka" = stevilske_spremenljivke,
                            "Labela" = sapply(stevilske_spremenljivke, FUN = function(x) attr(baza1[[x]], which = "label")))
    
    # vrne min od tiste baze, kjer je manjša vrednost minimum
    tabela_st[["Min"]] <- pmin(sapply(baza1_na, min, na.rm = TRUE),
                               sapply(baza2_na, min, na.rm = TRUE))
    
    # vrne max od tiste baze, kjer je višja vrednost maksimum
    tabela_st[["Max"]] <- pmax(sapply(baza1_na, max, na.rm = TRUE),
                               sapply(baza2_na, max, na.rm = TRUE))
    
    ## Neutežene statistike ----------------------------------------------------
    
    # N neutežen baza 1
    tabela_st[[paste0("N - ", ime_baza1)]] <- colSums(!is.na(baza1_na))
    
    # N neutežen baza 2
    tabela_st[[paste0("N - ", ime_baza2)]] <- colSums(!is.na(baza2_na))
    
    # Neutežen Welchev t-test
    # var 0 ne bo delovalo!
    # statistike <- lapply(stevilske_spremenljivke, FUN = function(x){
    #   test <- t.test(x = baza1_na[[x]], y = baza2_na[[x]],
    #                  paired = FALSE, var.equal = FALSE)
    #   
    #   c("povp1" = test$estimate[["mean of x"]],
    #     "povp2" = test$estimate[["mean of y"]],
    #     "t"     = test$statistic[[1]],
    #     "p"     = test$p.value)
    # })
    
    statistike <- lapply(stevilske_spremenljivke, FUN = function(x) wtd_t_test(x = baza1_na[[x]], y = baza2_na[[x]]))
    
    # Neuteženo povprečje baza 1
    tabela_st[[paste0("Povprečje - ", ime_baza1)]] <- sapply(statistike, function(x) x[["povp1"]])
    
    # Neuteženo povprečje baza 2
    tabela_st[[paste0("Povprečje - ", ime_baza2)]] <- sapply(statistike, function(x) x[["povp2"]])
    
    # Absolutna razlika neuteženih povprečji
    tabela_st[["Razlika v povprečjih - absolutna"]] <- sapply(statistike, function(x) x[["diff"]])
    
    # Relativna razlika neuteženih povprečji
    tabela_st[["Razlika v povprečjih - relativna (%)"]] <- (tabela_st[["Razlika v povprečjih - absolutna"]]/tabela_st[[paste0("Povprečje - ", ime_baza1)]])*100
    
    # T-vrednost in signifikanca
    tabela_st[["t"]] <- sapply(statistike, function(x) x[["t"]])
    
    tabela_st[["p"]] <- sapply(statistike, function(x) x[["p"]])
    
    tabela_st[["Signifikanca"]] <- weights::starmaker(tabela_st[["p"]])
    
    ## Utežene statistike ------------------------------------------------------
    
    # Utežen Welchev t-test
    # utezene_statistike <- lapply(stevilske_spremenljivke, FUN = function(x){
    #   test <- weights::wtd.t.test(x = baza1_na[[x]], y = baza2_na[[x]],
    #                               weight = utezi1, weighty = utezi2, samedata = FALSE)
    #   
    #   c("povp1" = test$additional[["Mean.x"]],
    #     "povp2" = test$additional[["Mean.y"]],
    #     "t"     = test$coefficients[["t.value"]],
    #     "p"     = test$coefficients[["p.value"]])
    # })
    
    utezene_statistike <- lapply(stevilske_spremenljivke, FUN = function(x){
      wtd_t_test(x = baza1_na[[x]], y = baza2_na[[x]], weights_x = utezi1, weights_y = utezi2)
    })
    
    tabela_st_u <- data.frame(row.names = seq_along(utezene_statistike))
    
    # N utežen baza 1
    # tabela_st_u[[paste0("N - ", ime_baza1)]] <- sapply(stevilske_spremenljivke, function(x) sum(utezi1[!is.na(baza1_na[[x]])]))
    
    # N utežen baza 2
    # tabela_st_u[[paste0("N - ", ime_baza2)]] <- sapply(stevilske_spremenljivke, function(x) sum(utezi2[!is.na(baza2_na[[x]])]))
    
    # Uteženo povprečje baza 1
    tabela_st_u[[paste0("Povprečje - ", ime_baza1)]] <- sapply(utezene_statistike, function(x) x[["povp1"]])
    
    # Uteženo povprečje baza 2
    tabela_st_u[[paste0("Povprečje - ", ime_baza2)]] <- sapply(utezene_statistike, function(x) x[["povp2"]])
    
    # Absolutna razlika uteženih povprečji
    tabela_st_u[["Razlika v povprečjih - absolutna"]] <- sapply(utezene_statistike, function(x) x[["diff"]])
    
    # Relativna razlika uteženih povprečji
    tabela_st_u[["Razlika v povprečjih - relativna (%)"]] <- (tabela_st_u[["Razlika v povprečjih - absolutna"]]/tabela_st_u[[paste0("Povprečje - ", ime_baza1)]])*100
    
    # T-vrednost in signifikanca
    tabela_st_u[["t"]] <- sapply(utezene_statistike, function(x) x[["t"]])
    
    tabela_st_u[["p"]] <- sapply(utezene_statistike, function(x) x[["p"]])
    
    tabela_st_u[["Signifikanca"]] <- weights::starmaker(tabela_st_u[["p"]])
    
    ## Izvoz excel -------------------------------------------------------------
    
    # povzetek za številske spr.
    merge_tabela_st <- cbind(tabela_st, tabela_st_u)
    frekvence_rel_razlike_neutezene <- count_rel_diff(vec = merge_tabela_st[[10]], p_vec = merge_tabela_st[[12]]) # to spremeni, da je na tabela_st in tabela_st_u
    frekvence_rel_razlike_utezene <- count_rel_diff(vec = merge_tabela_st[[17]], p_vec = merge_tabela_st[[19]])
    
    writeData(wb = wb,
              sheet = "Povzetek",
              
              x = data.frame("Relativne razlike" = c("> 20%", "(10% - 20%]", "[5% - 10%]", "< 5%"),
                             # neutežene statistike
                             "f" = frekvence_rel_razlike_neutezene$sums,
                             "%" = frekvence_rel_razlike_neutezene$sums/sum(frekvence_rel_razlike_neutezene$sums),
                             "Kumul f" = frekvence_rel_razlike_neutezene$cumsums,
                             "Kumul %" = frekvence_rel_razlike_neutezene$cumsums/sum(frekvence_rel_razlike_neutezene$sums),
                             "f" = frekvence_rel_razlike_neutezene$p_sums,
                             "% od vseh spremenljivk" = frekvence_rel_razlike_neutezene$p_sums/sum(frekvence_rel_razlike_neutezene$sums),
                             "% od relativne razlike" = frekvence_rel_razlike_neutezene$p_sums/frekvence_rel_razlike_neutezene$sums,
                             "Kumul f" = frekvence_rel_razlike_neutezene$p_cumsums,
                             "Kumul %" = frekvence_rel_razlike_neutezene$p_cumsums/sum(frekvence_rel_razlike_neutezene$p_sums),
                             
                             # utežene statistike
                             "f" = frekvence_rel_razlike_utezene$sums,
                             "%" = frekvence_rel_razlike_utezene$sums/sum(frekvence_rel_razlike_utezene$sums),
                             "Kumul f" = frekvence_rel_razlike_utezene$cumsums,
                             "Kumul %" = frekvence_rel_razlike_utezene$cumsums/sum(frekvence_rel_razlike_utezene$sums),
                             "f" = frekvence_rel_razlike_utezene$p_sums,
                             "% od vseh spremenljivk" = frekvence_rel_razlike_utezene$p_sums/sum(frekvence_rel_razlike_utezene$sums),
                             "% od relativne razlike" = frekvence_rel_razlike_utezene$p_sums/frekvence_rel_razlike_utezene$sums,
                             "Kumul f" = frekvence_rel_razlike_utezene$p_cumsums,
                             "Kumul %" = frekvence_rel_razlike_utezene$p_cumsums/sum(frekvence_rel_razlike_utezene$p_sums),
                             check.names = FALSE),
              
              borders = "all", startRow = 4, 
              headerStyle = createStyle(textDecoration = "bold",
                                        border = c("top", "bottom", "left", "right"),
                                        halign = "center", valign = "center", wrapText = TRUE), keepNA = FALSE)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Številske spremenljivke", startCol = 2, startRow = 1)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:19, rows = 1)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Neutežene statistike", startCol = 2, startRow = 2)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:10, rows = 2)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Utežene statistike", startCol = 11, startRow = 2)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 11:19, rows = 2)
    
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Spremenljivke", startCol = 2, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:5, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Spremenljivke", startCol = 11, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 11:14, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne spremenljivke (p < 0.05)",
              startCol = 6, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 6:10, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne spremenljivke (p < 0.05)",
              startCol = 15, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 15:19, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh spremenljivk:", sum(frekvence_rel_razlike_neutezene$sums)),
                          paste("Št. statistično značilnih spremenljivk (p < 0.05) - neuteženi podatki:", sum(frekvence_rel_razlike_neutezene$p_sums)),
                          paste("Št. statistično značilnih spremenljivk (p < 0.05) - uteženi podatki:", sum(frekvence_rel_razlike_utezene$p_sums)))),
              startCol = 1, startRow = 10, rowNames = FALSE, colNames = FALSE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 1:2, cols = 1:19,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                border = c("top", "bottom", "left", "right")),
             rows = 3, cols = 1:19,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(numFmt = "0.0%"),
             rows = 5:8, cols = c(3, 5, 7, 8, 10, 12, 14, 16, 17, 19),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(fgFill = "#DAEEF3"), rows = 3:8, cols = 2:10,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(fgFill = "#FDE9D9"), rows = 3:8, cols = 11:19,
             gridExpand = TRUE, stack = TRUE)
    
    setColWidths(wb = wb, sheet = "Povzetek", cols = 1:19,
                 widths = c(16, rep(c(5, 6, 8, 8, 5, 21, 21, 8, 8), 2)))
    
    setRowHeights(wb = wb, sheet = "Povzetek", rows = 1:4, heights = c(20, 20, 25, 44))
    
    
    # statistike
    addWorksheet(wb = wb, sheetName = "Opisne statistike", gridLines = FALSE)
    
    writeData(wb = wb,
              sheet = "Opisne statistike",
              x = merge_tabela_st,
              borders = "all", startRow = 2,
              headerStyle = createStyle(textDecoration = "bold",
                                        wrapText = TRUE,
                                        border = c("top", "bottom", "left", "right"),
                                        halign = "center", valign = "center"))
    
    writeData(wb = wb, sheet = "Opisne statistike",
              x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001", xy = c(1, 1))
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(fontSize = 9),
             rows = 1, cols = 1)
    
    writeData(wb = wb, sheet = "Opisne statistike",
              x = "Neutežene statistike", xy = c(7, 1))
    
    mergeCells(wb = wb, sheet = "Opisne statistike", cols = 7:13, rows = 1)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 1, cols = 7)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(numFmt = "0.00"), rows = 3:(nrow(tabela_st)+2), cols = 7:12,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(fgFill = "#DAEEF3"), rows = 2:(nrow(tabela_st)+2), cols = 7:13,
             gridExpand = TRUE, stack = TRUE)
    
    writeData(wb = wb, sheet = "Opisne statistike",
              x = "Utežene statistike", xy = c(14, 1))
    
    mergeCells(wb = wb, sheet = "Opisne statistike", cols = 14:20, rows = 1)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 1, cols = 14)
    
    # addStyle(wb = wb, sheet = "Opisne statistike",
    #          style = createStyle(numFmt = "0"), rows = 3:(nrow(tabela_st)+2), cols = 14:15,
    #          gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(numFmt = "0.00"), rows = 3:(nrow(tabela_st)+2), cols = 14:19,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(fgFill = "#FDE9D9"), rows = 2:(nrow(tabela_st)+2), cols = 14:20,
             gridExpand = TRUE, stack = TRUE)
  }
  
  if(!is.null(nominalne_spremenljivke)){
    # Nominalne spremenljivke -------------------------------------------------

    imena_nom_1 <- names(baza1)[names(baza1) %in% nominalne_spremenljivke]
    imena_nom_2 <- names(baza2)[names(baza2) %in% nominalne_spremenljivke]
    
    non_spr <- union(setdiff(imena_nom_1, imena_nom_2), setdiff(imena_nom_2, imena_nom_1))
    
    if(length(non_spr) != 0){
      warning(paste("Imena podanih nominalnih spremenljivk se ne ujemajo v obeh bazah. To so spremenljivke", paste(non_spr, collapse = ", "), "in so bile zato odstranjene iz analiz."))
      
      writeData(wb = wb, sheet = "Opozorila", xy = c(1,4),
                x = paste("Imena podanih nominalnih spremenljivk se ne ujemajo v obeh bazah. To so spremenljivke", paste(non_spr, collapse = ", "), "in so bile zato odstranjene iz analiz."))
      
      warning_counter <- TRUE
    }
    
    # izberemo samo spremenljivke, ki so prisotne v 1. in 2. bazi
    nominalne_spremenljivke <- imena_nom_1[imena_nom_1 %in% imena_nom_2]
    
    # funkcija za neutežene in utežene frekvenčne statistike
    tabela_nominalne <- function(spr, baza1, baza2, utezi1, utezi2){
      
      data1 <- as_factor(user_na_to_na(baza1[[spr]]), levels = "both")
      data2 <- as_factor(user_na_to_na(baza2[[spr]]), levels = "both")
      
      if(!all(levels(data1) == levels(data2))){
        warning(paste("Kategorije v bazah se pri nominalni spremenljivki", spr, "ne ujemajo. Spremenljivka je bila izločena iz analiz."))
        
        return(paste("Kategorije v bazah se pri nominalni spremenljivki", spr, "ne ujemajo. Spremenljivka je bila izločena iz analiz."))
        
      } else {
        
        ## Neutežene statistike ----------------------------------------------------
        
        tabela_nom <- as.data.frame(table(data1))
        names(tabela_nom) <- c(paste0(spr, " - ", attr(x = baza1[[spr]], which = "label")),
                               paste0("N - ", ime_baza1))
        tabela_nom[,1] <- as.character(tabela_nom[,1])
        
        tabela_nom[[paste0("N - ", ime_baza2)]] <- as.data.frame(table(data2))[,2]
        
        tabela_nom[[paste0("Delež (%) - ", ime_baza1)]] <- (tabela_nom[[paste0("N - ", ime_baza1)]]/sum(tabela_nom[[paste0("N - ", ime_baza1)]]))*100 # popravi da bodo deleži iz statistik
        
        tabela_nom[[paste0("Delež (%) - ", ime_baza2)]] <- (tabela_nom[[paste0("N - ", ime_baza2)]]/sum(tabela_nom[[paste0("N - ", ime_baza2)]]))*100
        
        tabela_nom[["Razlika v deležih - absolutna"]] <- tabela_nom[[paste0("Delež (%) - ", ime_baza2)]] - tabela_nom[[paste0("Delež (%) - ", ime_baza1)]]
        
        tabela_nom[["Razlika v deležih - relativna (%)"]] <- (tabela_nom[["Razlika v deležih - absolutna"]]/tabela_nom[[paste0("Delež (%) - ", ime_baza1)]])*100
        
        dummies1 <- weights::dummify(data1, keep.na = TRUE)
        dummies2 <- weights::dummify(data2, keep.na = TRUE)
        
        # t-test
        # statistika_nom <- lapply(1:nrow(tabela_nom), FUN = function(i){ # popravi 1:nrow
        #   test <- t.test(x = dummies1[,i], y = dummies2[,i],
        #                  paired = FALSE, var.equal = FALSE)
        #   
        #   c("t" = test$statistic[[1]],
        #     "p" = test$p.value)
        # })
        
        statistika_nom <- lapply(seq_len(nrow(tabela_nom)), FUN = function(i) wtd_t_test(x = dummies1[,i], y = dummies2[,i]))
        
        tabela_nom[["t"]] <- sapply(statistika_nom, function(x) x[["t"]])
        
        tabela_nom[["p"]] <- sapply(statistika_nom, function(x) x[["p"]])
        
        tabela_nom[["Signifikanca"]] <- weights::starmaker(tabela_nom[["p"]])
        
        # Skupaj seštevek
        temp_df <- data.frame(t(c(NA, colSums(tabela_nom[,2:5]), rep(NA, ncol(tabela_nom)-5))))
        names(temp_df) <- names(tabela_nom)
        
        tabela_nom <- rbind(tabela_nom, temp_df)
        tabela_nom[nrow(tabela_nom),1] <- "Skupaj"
        
        ## Utežene statistike ------------------------------------------------------
        
        tabela_nom_u <- data.frame(row.names = seq_along(statistika_nom))
        
        # Utežena frekvenca baza 1
        tabela_nom_u[[paste0("N - ", ime_baza1)]] <- weighted_table(data1, utezi1)
        
        # Utežena frekvenca baza 2
        tabela_nom_u[[paste0("N - ", ime_baza2)]] <- weighted_table(data2, utezi2)
        
        # Utežen delež baza 1
        tabela_nom_u[[paste0("Delež (%) - ", ime_baza1)]] <- (tabela_nom_u[[paste0("N - ", ime_baza1)]]/sum(tabela_nom_u[[paste0("N - ", ime_baza1)]]))*100
        
        # Utežen delež baza 2
        tabela_nom_u[[paste0("Delež (%) - ", ime_baza2)]] <- (tabela_nom_u[[paste0("N - ", ime_baza2)]]/sum(tabela_nom_u[[paste0("N - ", ime_baza2)]]))*100
        
        # Absolutna razlika uteženih deležev
        tabela_nom_u[["Razlika v deležih - absolutna"]] <- tabela_nom_u[[paste0("Delež (%) - ", ime_baza2)]] - tabela_nom_u[[paste0("Delež (%) - ", ime_baza1)]]
        
        # Relativna razlika uteženih deležev
        tabela_nom_u[["Razlika v deležih - relativna (%)"]] <- (tabela_nom_u[["Razlika v deležih - absolutna"]]/tabela_nom_u[[paste0("Delež (%) - ", ime_baza1)]])*100
        
        # Utežen t-test
        # utezena_statistika_nom <- lapply(1:nrow(tabela_nom_u), FUN = function(i){
        #   test <- weights::wtd.t.test(x = dummies1[,i], y = dummies2[,i],
        #                               weight = utezi1, weighty = utezi2, samedata = FALSE)
        # 
        #   c("t" = test$coefficients[["t.value"]],
        #     "p" = test$coefficients[["p.value"]])
        # })
        
        utezena_statistika_nom <- lapply(seq_len(nrow(tabela_nom_u)), FUN = function(i){
          wtd_t_test(x = dummies1[,i], y = dummies2[,i], weights_x = utezi1, weights_y = utezi2)
        })
        
        tabela_nom_u[["t"]] <- sapply(utezena_statistika_nom, function(x) x[["t"]])
        
        tabela_nom_u[["p"]] <- sapply(utezena_statistika_nom, function(x) x[["p"]])
        
        tabela_nom_u[["Signifikanca"]] <- weights::starmaker(tabela_nom_u[["p"]])
        
        # Skupaj seštevek
        temp_df_u <- data.frame(t(c(colSums(tabela_nom_u[,1:4]), rep(NA, ncol(tabela_nom_u)-4))))
        names(temp_df_u) <- names(tabela_nom_u)
        
        tabela_nom_u <- rbind(tabela_nom_u, temp_df_u)
        
        # return function result
        return(cbind(tabela_nom, tabela_nom_u))
      }
    }
    
    ## Izvoz excel -------------------------------------------------------------
    
    factor_tables <- lapply(nominalne_spremenljivke, function(x) tabela_nominalne(spr = x, baza1 = baza1, baza2 = baza2, utezi1 = utezi1, utezi2 = utezi2))
    
    # dodamo še opozorila, če so prisotna
    vsota <- cumsum(sapply(factor_tables, is.character))
    
    for(i in seq_along(factor_tables)){
      if(is.character(factor_tables[[i]])){
        writeData(wb = wb, sheet = "Opozorila", startCol = 1, startRow = 4 + vsota[i],
                  x = factor_tables[[i]])
        
        warning_counter <- TRUE
      }
    }
    
    factor_tables <- factor_tables[!sapply(factor_tables, is.character)]
    
    # označi se vrstice, kjer je premajhen numerus za zanesljivo oceno p-vrednosti
    opozorilo_numerus <- lapply(factor_tables, FUN = function(x) {
      last_row <- length(x[[2]])
      # neuteženi
      n1 <- x[[2]][last_row]
      n2 <- x[[3]][last_row]
      rev1 <- n1 - x[[2]][-last_row]
      rev2 <- n2 - x[[3]][-last_row]
      
      # uteženi podatki
      n1u <- x[[11]][last_row]
      n2u <- x[[12]][last_row]
      rev1u <- n1u - x[[11]][-last_row]
      rev2u <- n2u - x[[12]][-last_row]
      
      neutez <- x[[2]][-last_row] <= 5 | x[[3]][-last_row] <= 5 | rev1 <= 5 | rev2 <= 5
      utez <- x[[11]][-last_row] <= 5 | x[[12]][-last_row] <= 5 | rev1u <= 5 | rev2u <= 5
      p_neutez <- x[[9]][-last_row] < 0.05
      p_utez <- x[[18]][-last_row] < 0.05
      
      list(neutez = neutez,
           utez = utez,
           p_neutez = p_neutez & neutez,
           p_utez = p_utez & utez)
    })
    
    p_numerus_neutezene_kategorije <- do.call(c, lapply(opozorilo_numerus, "[[", "p_neutez"))
    p_numerus_utezene_kategorije <- do.call(c, lapply(opozorilo_numerus, "[[", "p_utez"))
    
    # povzetek za nominalne spr.
    neutezene_kategorije <- abs(na.omit(do.call(rbind, lapply(factor_tables, "[", c(7, 9)))))
    utezene_kategorije <- abs(na.omit(do.call(rbind, lapply(factor_tables, "[", c(16, 18)))))
    
    frekvence_rel_razlike_neutezene_nom <- count_rel_diff(vec = neutezene_kategorije[[1]], p_vec = neutezene_kategorije[!p_numerus_neutezene_kategorije,][[2]])
    frekvence_rel_razlike_utezene_nom <- count_rel_diff(vec = utezene_kategorije[[1]], p_vec = utezene_kategorije[!p_numerus_utezene_kategorije,][[2]])
    
    writeData(wb = wb,
              sheet = "Povzetek",
              
              x = data.frame("Relativne razlike" = c("> 20%", "(10% - 20%]", "[5% - 10%]", "< 5%"),
                             # neutežene statistike
                             "f" = frekvence_rel_razlike_neutezene_nom$sums,
                             "%" = frekvence_rel_razlike_neutezene_nom$sums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                             "Kumul f" = frekvence_rel_razlike_neutezene_nom$cumsums,
                             "Kumul %" = frekvence_rel_razlike_neutezene_nom$cumsums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                             "f" = frekvence_rel_razlike_neutezene_nom$p_sums,
                             "% od vseh kategorij" = frekvence_rel_razlike_neutezene_nom$p_sums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                             "% od relativne razlike" = frekvence_rel_razlike_neutezene_nom$p_sums/frekvence_rel_razlike_neutezene_nom$sums,
                             "Kumul f" = frekvence_rel_razlike_neutezene_nom$p_cumsums,
                             "Kumul %" = frekvence_rel_razlike_neutezene_nom$p_cumsums/sum(frekvence_rel_razlike_neutezene_nom$p_sums),
                             
                             # utežene statistike
                             "f" = frekvence_rel_razlike_utezene_nom$sums,
                             "%" = frekvence_rel_razlike_utezene_nom$sums/sum(frekvence_rel_razlike_utezene_nom$sums),
                             "Kumul f" = frekvence_rel_razlike_utezene_nom$cumsums,
                             "Kumul %" = frekvence_rel_razlike_utezene_nom$cumsums/sum(frekvence_rel_razlike_utezene_nom$sums),
                             "f" = frekvence_rel_razlike_utezene_nom$p_sums,
                             "% od vseh kategorij" = frekvence_rel_razlike_utezene_nom$p_sums/sum(frekvence_rel_razlike_utezene_nom$sums),
                             "% od relativne razlike" = frekvence_rel_razlike_utezene_nom$p_sums/frekvence_rel_razlike_utezene_nom$sums,
                             "Kumul f" = frekvence_rel_razlike_utezene_nom$p_cumsums,
                             "Kumul %" = frekvence_rel_razlike_utezene_nom$p_cumsums/sum(frekvence_rel_razlike_utezene_nom$p_sums),
                             check.names = FALSE),
              
              borders = "all", startRow = 18,
              headerStyle = createStyle(textDecoration = "bold",
                                        wrapText = TRUE,
                                        border = c("top", "bottom", "left", "right"),
                                        halign = "center", valign = "center"))
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Nominalne spremenljivke", startCol = 2, startRow = 15)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:19, rows = 15)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Neutežene statistike", startCol = 2, startRow = 16)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:10, rows = 16)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Utežene statistike", startCol = 11, startRow = 16)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 11:19, rows = 16)
    
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Kategorije", startCol = 2, startRow = 17)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:5, rows = 17)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Kategorije", startCol = 11, startRow = 17)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 11:14, rows = 17)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne kategorije (p < 0.05)", startCol = 6, startRow = 17)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 6:10, rows = 17)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne kategorije (p < 0.05)", startCol = 15, startRow = 17)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 15:19, rows = 17)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh kategorij:", sum(frekvence_rel_razlike_neutezene_nom$sums)),
                          paste("Št. statistično značilnih kategorij (p < 0.05) - neuteženi podatki:", sum(frekvence_rel_razlike_neutezene_nom$p_sums)),
                          paste("Št. statistično značilnih kategorij (p < 0.05) - uteženi podatki:", sum(frekvence_rel_razlike_utezene_nom$p_sums)))),
              startCol = 1, startRow = 24, rowNames = FALSE, colNames = FALSE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 15:16, cols = 1:19,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 border = c("top", "bottom", "left", "right")),
             rows = 17, cols = 1:19,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(numFmt = "0.0%"),
             rows = 19:22, cols = c(3, 5, 7, 8, 10, 12, 14, 16, 17, 19),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(fgFill = "#DAEEF3"), rows = 17:22, cols = 2:10,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(fgFill = "#FDE9D9"), rows = 17:22, cols = 11:19,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = "Povzetek", rows = 15:18, heights = c(20, 20, 25, 44))
    
    # opozorilo če n <= 5
    if(any(p_numerus_neutezene_kategorije)) {
      writeData(wb = wb, sheet = "Povzetek",
                x = paste("Opomba: Kategorije (število:", sum(p_numerus_neutezene_kategorije),") so bile statistično značilne, vendar zaradi premajhnega št. enot niso bile upoštevane v povzetku."),
                startCol = 2, startRow = 23)
    }
    
    if(any(p_numerus_utezene_kategorije)) {
      writeData(wb = wb, sheet = "Povzetek",
                x = paste("Opomba: Kategorije (število:", sum(p_numerus_utezene_kategorije),") so bile statistično značilne, vendar zaradi premajhnega št. enot niso bile upoštevane v povzetku."),
                startCol = 11, startRow = 23)
    }
    
    # statistike
    addWorksheet(wb = wb,
                 sheetName = "Frekvencne tabele",
                 gridLines = FALSE)
    
    writeData(wb = wb, sheet = "Frekvencne tabele",
              x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001", xy = c(1, 2))
    
    addStyle(wb = wb, sheet = "Frekvencne tabele",
             style = createStyle(fontSize = 9),
             rows = 2, cols = 1)
    
    # starting row = number of rows of previous table
    start_rows <- c(0, cumsum(3 + sapply(factor_tables, nrow)[-length(factor_tables)])) + 3
    
    for(i in seq_along(factor_tables)){
      
      writeData(wb = wb, sheet = "Frekvencne tabele",
                x = factor_tables[[i]],
                borders = "all",
                startRow = start_rows[i],
                headerStyle = createStyle(textDecoration = "bold",
                                          wrapText = TRUE,
                                          border = c("top", "bottom", "left", "right"),
                                          halign = "center", valign = "center"))
      
      writeData(wb = wb, sheet = "Frekvencne tabele",
                x = "Neutežene statistike", startCol = 2, startRow = start_rows[i]-1)
      
      mergeCells(wb = wb, sheet = "Frekvencne tabele", cols = 2:10, rows = start_rows[i]-1)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(textDecoration = "bold",
                                   halign = "center", valign = "center",
                                   fontSize = 12),
               rows = start_rows[i]-1, cols = c(2, 11),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(fgFill = "#DAEEF3"), rows = start_rows[i]:(start_rows[i]+nrow(factor_tables[[i]])), cols = 2:10,
               gridExpand = TRUE, stack = TRUE)
      
      writeData(wb = wb, sheet = "Frekvencne tabele",
                x = "Utežene statistike", startCol = 11, startRow = start_rows[i]-1)
      
      mergeCells(wb = wb, sheet = "Frekvencne tabele", cols = 11:19, rows = start_rows[i]-1)
      
      # addStyle(wb = wb, sheet = "Frekvencne tabele",
      #          style = createStyle(textDecoration = "bold",
      #                              halign = "center", valign = "center",
      #                              fontSize = 12),
      #          rows = start_rows[i]-1, cols = 11,
      #          gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(fgFill = "#FDE9D9"), rows = start_rows[i]:(start_rows[i]+nrow(factor_tables[[i]])), cols = 11:19,
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(fgFill = "#FF5D5D"), rows = start_rows[i] + which(opozorilo_numerus[[i]]$neutez == TRUE), cols = 2:10,
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(fgFill = "#FF5D5D"), rows = start_rows[i] + which(opozorilo_numerus[[i]]$utez == TRUE), cols = 11:19,
               gridExpand = TRUE, stack = TRUE)
    }
    
    if(any(unlist(opozorilo_numerus))) {
      writeData(wb = wb, sheet = "Frekvencne tabele",
                x = "Opozorilo: Število enot v celici je premajhno za zanesljivo oceno p-vrednosti", startCol = 1, startRow = 1)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(fgFill = "#FF5D5D", wrapText = TRUE), rows = 1, cols = 1,
               gridExpand = TRUE, stack = TRUE)
    }
    
    addStyle(wb = wb, sheet = "Frekvencne tabele",
             style = createStyle(numFmt = "0.00"), rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])), cols = 4:9,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Frekvencne tabele",
             style = createStyle(numFmt = "0"), rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])), cols = 11:12,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Frekvencne tabele",
             style = createStyle(numFmt = "0.00"), rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])), cols = 13:18,
             gridExpand = TRUE, stack = TRUE)
    
    
    setColWidths(wb = wb, sheet = "Frekvencne tabele", cols = 1, widths = 60)
    
  }
  
  if(warning_counter == FALSE){
    writeData(wb = wb, sheet = "Opozorila", startCol = 1, startRow = 1,
              x = "Ni opozoril")
  }

  # shranimo excel datoteko
  saveWorkbook(wb = wb, file = file, overwrite = TRUE)
}




# baza1 <- haven::read_spss(file = "C:/Users/strle/OneDrive/Dokumenti/Delo/Weighting/Test files/2. CDI, PANDA, 21. val, koncano, uteži.sav",
#                              user_na = TRUE)
# utezi1 <- baza1$weights
# 
# baza2 <- haven::read_spss(file = "C:/Users/strle/OneDrive/Dokumenti/Delo/Weighting/Test files/3. Valicon, PANDA - 21. val, koncano, uteži.sav",
#                              user_na = TRUE)
# utezi2 <- baza2$weights
# 
# stevilske_spremenljivke <- c("VACC_EFFECT_v2",
#                              "VACC_NATURAL",
#                              "VACC_OBLIGATION",
#                              "TRUST_JOURNAL",
#                              "TRUST_HOSPITALS",
#                              "TRUST_NCDC",
#                              "TRUST_POLITICIANS",
#                              "TRUST_DOCTOR",
#                              "TRUST_NATIONAL_HEALTH",
#                              "TRUST_SCIENCE",
#                              "TRUST_POSSK19")
