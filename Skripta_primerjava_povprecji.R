# potrebni paketi
library(haven)
library(labelled)
library(weights)
library(openxlsx)


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
  
  if(!is.null(stevilske_spremenljivke) && !is.character(stevilske_spremenljivke)){
    stop("Podane številske spremenljivke morajo biti v obliki character vector.")
  }
  
  if(!is.null(nominalne_spremenljivke) && !is.character(nominalne_spremenljivke)){
    stop("Podane nominalne spremenljivke morajo biti v obliki character vector.")
  }
  
  wb <- createWorkbook()
  
  addWorksheet(wb = wb, sheetName = "Opozorila", gridLines = FALSE)
  
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
    tabela_st[[paste0("N - ", ime_baza1)]] <- colSums(!sapply(baza1_na, is.na))
    
    # N neutežen baza 2
    tabela_st[[paste0("N - ", ime_baza2)]] <- colSums(!sapply(baza2_na, is.na))
    
    # Neutežen Welchev t-test
    # var 0 ne bo delovalo!
    statistike <- lapply(stevilske_spremenljivke, FUN = function(x){
      test <- t.test(x = baza1_na[[x]], y = baza2_na[[x]],
                     paired = FALSE, var.equal = FALSE)
      
      c("povp1" = test$estimate[["mean of x"]],
        "povp2" = test$estimate[["mean of y"]],
        "t"     = test$statistic[[1]],
        "p"     = test$p.value)
    })
    
    # Neuteženo povprečje baza 1
    tabela_st[[paste0("Povprečje - ", ime_baza1)]] <- sapply(statistike, function(x) x[["povp1"]])
    
    # Neuteženo povprečje baza 2
    tabela_st[[paste0("Povprečje - ", ime_baza2)]] <- sapply(statistike, function(x) x[["povp2"]])
    
    # Absolutna razlika neuteženih povprečji
    tabela_st[["Razlika v povprečjih - absolutna"]] <- tabela_st[[paste0("Povprečje - ", ime_baza2)]] - tabela_st[[paste0("Povprečje - ", ime_baza1)]]
    
    # Relativna razlika neuteženih povprečji
    tabela_st[["Razlika v povprečjih - relativna (%)"]] <- (tabela_st[["Razlika v povprečjih - absolutna"]]/tabela_st[[paste0("Povprečje - ", ime_baza1)]])*100
    
    # T-vrednost in signifikanca
    tabela_st[["T-vrednost"]] <- sapply(statistike, function(x) x[["t"]])
    
    tabela_st[["P-vrednost"]] <- sapply(statistike, function(x) x[["p"]])
    
    tabela_st[["Signifikanca"]] <- weights::starmaker(tabela_st[["P-vrednost"]])
    
    ## Utežene statistike ------------------------------------------------------
    
    # Utežen Welchev t-test
    utezene_statistike <- lapply(stevilske_spremenljivke, FUN = function(x){
      test <- weights::wtd.t.test(x = baza1_na[[x]], y = baza2_na[[x]],
                                  weight = utezi1, weighty = utezi2, samedata = FALSE)
      
      c("povp1" = test$additional[["Mean.x"]],
        "povp2" = test$additional[["Mean.y"]],
        "t"     = test$coefficients[["t.value"]],
        "p"     = test$coefficients[["p.value"]])
    })
    
    tabela_st_u <- data.frame(row.names = seq_along(utezene_statistike))
    
    # N utežen baza 1
    tabela_st_u[[paste0("N - ", ime_baza1)]] <- sapply(stevilske_spremenljivke, function(x) sum(utezi1[!is.na(baza1_na[[x]])]))
    
    # N utežen baza 2
    tabela_st_u[[paste0("N - ", ime_baza2)]] <- sapply(stevilske_spremenljivke, function(x) sum(utezi2[!is.na(baza2_na[[x]])]))
    
    # Uteženo povprečje baza 1
    tabela_st_u[[paste0("Povprečje - ", ime_baza1)]] <- sapply(utezene_statistike, function(x) x[["povp1"]])
    
    # Uteženo povprečje baza 2
    tabela_st_u[[paste0("Povprečje - ", ime_baza2)]] <- sapply(utezene_statistike, function(x) x[["povp2"]])
    
    # Absolutna razlika uteženih povprečji
    tabela_st_u[["Razlika v povprečjih - absolutna"]] <- tabela_st_u[[paste0("Povprečje - ", ime_baza2)]] - tabela_st_u[[paste0("Povprečje - ", ime_baza1)]]
    
    # Relativna razlika uteženih povprečji
    tabela_st_u[["Razlika v povprečjih - relativna (%)"]] <- (tabela_st_u[["Razlika v povprečjih - absolutna"]]/tabela_st_u[[paste0("Povprečje - ", ime_baza1)]])*100
    
    # T-vrednost in signifikanca
    tabela_st_u[["T-vrednost"]] <- sapply(utezene_statistike, function(x) x[["t"]])
    
    tabela_st_u[["P-vrednost"]] <- sapply(utezene_statistike, function(x) x[["p"]])
    
    tabela_st_u[["Signifikanca"]] <- weights::starmaker(tabela_st_u[["P-vrednost"]])
    
    ## Izvoz excel -------------------------------------------------------------
    
    addWorksheet(wb = wb, sheetName = "Opisne statistike", gridLines = FALSE)
    
    writeData(wb = wb,
              sheet = "Opisne statistike",
              x = cbind(tabela_st, tabela_st_u),
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
              x = "Neutežene statistike", xy = c(5, 1))
    
    mergeCells(wb = wb, sheet = "Opisne statistike", cols = 5:13, rows = 1)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 1, cols = 5)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(numFmt = "0.00"), rows = 3:(nrow(tabela_st)+2), cols = 7:12,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(fgFill = "#DAEEF3"), rows = 2:(nrow(tabela_st)+2), cols = 5:13,
             gridExpand = TRUE, stack = TRUE)
    
    writeData(wb = wb, sheet = "Opisne statistike",
              x = "Utežene statistike", xy = c(14, 1))
    
    mergeCells(wb = wb, sheet = "Opisne statistike", cols = 14:22, rows = 1)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 1, cols = 14)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(numFmt = "0"), rows = 3:(nrow(tabela_st)+2), cols = 14:15,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(numFmt = "0.00"), rows = 3:(nrow(tabela_st)+2), cols = 16:21,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Opisne statistike",
             style = createStyle(fgFill = "#FDE9D9"), rows = 2:(nrow(tabela_st)+2), cols = 14:22,
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
    
    # funkcija za utežene frekvence
    weighted_table = function(x, weights) {
      vapply(split(weights, x), sum, 0, na.rm = TRUE)
    }
    
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
        
        tabela_nom[[paste0("Delež (%) - ", ime_baza1)]] <- (tabela_nom[[paste0("N - ", ime_baza1)]]/sum(tabela_nom[[paste0("N - ", ime_baza1)]]))*100
        
        tabela_nom[[paste0("Delež (%) - ", ime_baza2)]] <- (tabela_nom[[paste0("N - ", ime_baza2)]]/sum(tabela_nom[[paste0("N - ", ime_baza2)]]))*100
        
        tabela_nom[["Razlika v deležih - absolutna"]] <- tabela_nom[[paste0("Delež (%) - ", ime_baza2)]] - tabela_nom[[paste0("Delež (%) - ", ime_baza1)]]
        
        tabela_nom[["Razlika v deležih - relativna (%)"]] <- (tabela_nom[["Razlika v deležih - absolutna"]]/tabela_nom[[paste0("Delež (%) - ", ime_baza1)]])*100
        
        dummies1 <- weights::dummify(data1, keep.na = TRUE)
        dummies2 <- weights::dummify(data2, keep.na = TRUE)
        
        # t-test
        statistika_nom <- lapply(1:nrow(tabela_nom), FUN = function(i){
          test <- t.test(x = dummies1[,i], y = dummies2[,i],
                         paired = FALSE, var.equal = FALSE)
          
          c("t" = test$statistic[[1]],
            "p" = test$p.value)
        })
        
        tabela_nom[["T-vrednost"]] <- sapply(statistika_nom, function(x) x[["t"]])
        
        tabela_nom[["P-vrednost"]] <- sapply(statistika_nom, function(x) x[["p"]])
        
        tabela_nom[["Signifikanca"]] <- weights::starmaker(tabela_nom[["P-vrednost"]])
        
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
        utezena_statistika_nom <- lapply(1:nrow(tabela_nom_u), FUN = function(i){
          test <- weights::wtd.t.test(x = dummies1[,i], y = dummies2[,i],
                                      weight = utezi1, weighty = utezi2, samedata = FALSE)
          
          c("t" = test$coefficients[["t.value"]],
            "p" = test$coefficients[["p.value"]])
        })
        
        tabela_nom_u[["T-vrednost"]] <- sapply(utezena_statistika_nom, function(x) x[["t"]])
        
        tabela_nom_u[["P-vrednost"]] <- sapply(utezena_statistika_nom, function(x) x[["p"]])
        
        tabela_nom_u[["Signifikanca"]] <- weights::starmaker(tabela_nom_u[["P-vrednost"]])
        
        # Skupaj seštevek
        temp_df_u <- data.frame(t(c(colSums(tabela_nom_u[,1:4]), rep(NA, ncol(tabela_nom_u)-4))))
        names(temp_df_u) <- names(tabela_nom_u)
        
        tabela_nom_u <- rbind(tabela_nom_u, temp_df_u)
        
        # return function result
        return(cbind(tabela_nom, tabela_nom_u))
      }
    }
    
    
    ## Izvoz excel -------------------------------------------------------------
    
    factor_tables <- lapply(nominalne_spremenljivke, function(x) tabela_nominalne(spr = x, baza1, baza2, utezi1, utezi2))
    
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
    
    addWorksheet(wb = wb,
                 sheetName = "Frekvencne tabele",
                 gridLines = FALSE)
    
    writeData(wb = wb, sheet = "Frekvencne tabele",
              x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001", xy = c(1, 1))
    
    addStyle(wb = wb, sheet = "Frekvencne tabele",
             style = createStyle(fontSize = 9),
             rows = 1, cols = 1)
    
    # starting row = number of rows of previous table
    # + 2 (for the header and to add a empty row)
    # + 1 for the first table
    start_rows <- c(0, cumsum(3 + sapply(factor_tables, nrow)[-length(factor_tables)])) + 2
    
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
               rows = start_rows[i]-1, cols = 2)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(fgFill = "#DAEEF3"), rows = start_rows[i]:(start_rows[i]+nrow(factor_tables[[i]])), cols = 2:10,
               gridExpand = TRUE, stack = TRUE)
      
      writeData(wb = wb, sheet = "Frekvencne tabele",
                x = "Utežene statistike", startCol = 11, startRow = start_rows[i]-1)
      
      mergeCells(wb = wb, sheet = "Frekvencne tabele", cols = 11:19, rows = start_rows[i]-1)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(textDecoration = "bold",
                                   halign = "center", valign = "center",
                                   fontSize = 12),
               rows = start_rows[i]-1, cols = 11)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(fgFill = "#FDE9D9"), rows = start_rows[i]:(start_rows[i]+nrow(factor_tables[[i]])), cols = 11:19,
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




