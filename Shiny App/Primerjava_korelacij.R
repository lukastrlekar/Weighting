
izvoz_excel_korelacije <- function(baza1 = NULL,
                                   baza2 = NULL,
                                   ime_baza1 = "baza 1",
                                   ime_baza2 = "baza 2",
                                   utezi1 = NULL,
                                   utezi2 = NULL,
                                   stevilske_spremenljivke = NULL,
                                   se_calculation,
                                   survey_design1 = NULL,
                                   survey_design2 = NULL,
                                   file){
  
  if(se_calculation == "survey_se"){
    utezi1 <- weights(survey_design1)
    utezi2 <- weights(survey_design2)
  }
  
  wb <- createWorkbook()
  addWorksheet(wb = wb, sheetName = "Opozorila", gridLines = FALSE)
  addWorksheet(wb = wb, sheetName = "Povzetek", gridLines = FALSE)
  
  warning_counter <- FALSE
  
  # Številske spremenljivke -------------------------------------------------
  
  imena_st_1 <- names(baza1)[names(baza1) %in% stevilske_spremenljivke]
  imena_st_2 <- names(baza2)[names(baza2) %in% stevilske_spremenljivke]
  
  # preverimo ujemanje imen spremenljivk v obeh bazah
  non_st_spr <- union(setdiff(imena_st_1, imena_st_2), setdiff(imena_st_2, imena_st_1))
  
  if(length(imena_st_1) != length(imena_st_2)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,1),
              x = paste("Imena podanih številskih spremenljivk se ne ujemajo v obeh bazah. To so spremenljivke", paste(non_st_spr, collapse = ", "), "in so bile zato odstranjene iz analiz."))
    
    warning_counter <- TRUE
  }
  
  # izberemo samo spremenljivke, ki so prisotne v 1. in 2. bazi
  stevilske_spremenljivke <- imena_st_1[imena_st_1 %in% imena_st_2]
  
  # uredimo bazi v isti vrstni red spremenljivk
  baza1 <- baza1[,stevilske_spremenljivke, drop = FALSE]
  baza2 <- baza2[,stevilske_spremenljivke, drop = FALSE]
  
  # preverimo, da je isto število kategorij v obeh bazah
  levels1 <- lapply(stevilske_spremenljivke, function(x) attr(baza1[[x]], "labels", exact = TRUE))
  levels2 <- lapply(stevilske_spremenljivke, function(x) attr(baza2[[x]], "labels", exact = TRUE))
  
  indeksi <- vapply(seq_along(stevilske_spremenljivke), function(i) all(levels1[[i]] == levels2[[i]]), FUN.VALUE = logical(1))
  
  if(!all(indeksi)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,2),
              x = paste("Kategorije se pri številski/h spremenljivki/ah", paste(stevilske_spremenljivke[!indeksi], collapse = ", "), "ne ujemajo. Te spremenljivke niso bile odstranjene iz analiz, vendar svetujemo previdnost pri analizi."))
    
    warning_counter <- TRUE
  }
  
  # preverimo, da so spremenljivke res številske in neštevilske odstranimo
  numeric_variables <- vapply(stevilske_spremenljivke, function(x){
    t1 <- is.numeric(baza1[[x]]) | is.integer(baza1[[x]])
    t2 <- is.numeric(baza2[[x]]) | is.integer(baza2[[x]])
    all(t1, t2)
  }, FUN.VALUE = logical(1), USE.NAMES = FALSE)
  
  if(!all(numeric_variables)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,3),
              x = paste("Spremenljivke", paste(stevilske_spremenljivke[!numeric_variables], collapse = ", "), "so bile odstranjene iz analiz, saj vsebujejo neštevilske vrednosti oziroma so bile v SPSS definirane kot neštevilske v eni ali obeh bazah."))
    
    warning_counter <- TRUE
    
    stevilske_spremenljivke <- stevilske_spremenljivke[numeric_variables]
  }
  
  # preverimo in odstranimo spremenljivke z ničelno varianco (konstante) (tu se odstranijo tudi spremenljivke z vsemi NA)
  indeksi_variances <- vapply(stevilske_spremenljivke, function(x){
    v1 <- var(baza1[[x]], na.rm = TRUE)
    v1 <- ifelse(is.na(v1), 0, v1)
    v2 <- var(baza2[[x]], na.rm = TRUE)
    v2 <- ifelse(is.na(v2), 0, v2)
    any(v1 == 0, v2 == 0)
  }, FUN.VALUE = logical(1), USE.NAMES = FALSE)
  
  if(any(indeksi_variances)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,4),
              x = paste("Številske spremenljivke", paste(stevilske_spremenljivke[indeksi_variances], collapse = ", "), "so bile odstranjene, saj so konstante (imajo ničelno varianco) v eni ali obeh bazah."))
    
    warning_counter <- TRUE
    
    stevilske_spremenljivke <- stevilske_spremenljivke[!indeksi_variances]
  }
  
  baza1 <- baza1[,stevilske_spremenljivke, drop = FALSE]
  baza2 <- baza2[,stevilske_spremenljivke, drop = FALSE]
  
  if(length(stevilske_spremenljivke) > 1){
    
    addWorksheet(wb = wb, sheetName = "Korelacije", gridLines = FALSE)
    
    pairwise_n <- function(mat){
      cols <- colnames(mat)
      nn <- lower.tri(matrix(nrow = length(cols), ncol = length(cols)), diag = FALSE)
      rownames(nn) <- colnames(nn) <- cols
      mat_na <- is.na(mat)
      
      fun_pairwise_n <- function(x, y){
        sum(apply(mat_na[, c(x, y)], 1, function(z) !any(z)))
      }
      
      for (i in 1:nrow(nn)){
        for (j in 1:ncol(nn)){
          if(nn[i,j] == TRUE) nn[i,j] <- fun_pairwise_n(rownames(nn)[i], colnames(nn)[j]) else nn[i,j] <- NA
        }
      }
      return(nn)
    }
    
    # velikosti vzorca 1. baza
    n_baza1 <- pairwise_n(mat = baza1)
    n_baza1 <- n_baza1[lower.tri(n_baza1)]
    
    # velikosti vzorca 2. baza
    n_baza2 <- pairwise_n(mat = baza2)
    n_baza2 <- n_baza2[lower.tri(n_baza2)]
    
    # neutežene korelacije v 1. bazi
    cors_baza1 <- cor(x = baza1, use = "pairwise.complete.obs")
    cors_baza1 <- cors_baza1[lower.tri(cors_baza1)]
    
    # neutežene korelacije v 2. bazi
    cors_baza2 <- cor(x = baza2, use = "pairwise.complete.obs")
    cors_baza2 <- cors_baza2[lower.tri(cors_baza2)]
    
    # utežene korelacije v 1. bazi
    w_cors_baza1 <- weights::wtd.cors(x = baza1, weight = utezi1)
    w_cors_baza1 <- w_cors_baza1[lower.tri(w_cors_baza1)]
    
    # utežene korelacije v 2. bazi
    w_cors_baza2 <- weights::wtd.cors(x = baza2, weight = utezi2)
    w_cors_baza2 <- w_cors_baza2[lower.tri(w_cors_baza2)]
    
    # primerjava korelacij (p vrednosti, relativne razlike)
    
    p_values <- vapply(seq_along(cors_baza1), FUN = function(i){
      test <- cocor::cocor.indep.groups(r1.jk = cors_baza1[[i]],
                                        r2.hm = cors_baza2[[i]],
                                        n1 = n_baza1[[i]],
                                        n2 = n_baza2[[i]],
                                        test = "fisher1925")
      test@fisher1925$p.value
    }, FUN.VALUE = numeric(1))
    
    rel_changes <- ((cors_baza2 - cors_baza1)/abs(cors_baza1))*100
    
    w_p_values <- vapply(seq_along(w_cors_baza1), FUN = function(i){
      test <- cocor::cocor.indep.groups(r1.jk = w_cors_baza1[[i]],
                                        r2.hm = w_cors_baza2[[i]],
                                        n1 = n_baza1[[i]],
                                        n2 = n_baza2[[i]],
                                        test = "fisher1925")
      test@fisher1925$p.value
    }, FUN.VALUE = numeric(1))
    
    w_rel_changes <- ((w_cors_baza2 - w_cors_baza1)/abs(w_cors_baza1))*100
    
    write_cors_excel <- function(cors_baza1,
                                 cors_baza2,
                                 n_baza1,
                                 n_baza2,
                                 rel_changes,
                                 p_values,
                                 vars = stevilske_spremenljivke,
                                 row_start,
                                 labele){
      
      n_vars <- length(vars)
      
      m1 <- m2 <- m3 <- matrix(nrow = n_vars,
                               ncol = n_vars)
      
      # korelacije
      m1[lower.tri(m1)] <- paste(round(cors_baza1, 2),
                                 ",",
                                 round(cors_baza2,2))
      
      # absolutne, relativne razlike, p-vrednosti
      m2[lower.tri(m2)] <- paste(round(cors_baza2 - cors_baza1, 2),
                                 paste0("(",round(rel_changes, 0), "%)"),
                                 weights::starmaker(p_values))
      diag(m2) <- "—"
      
      # velikost vzorca
      m3[lower.tri(m3)] <- paste(n_baza1, ",", n_baza2)
      
      # matrika za prikaz korelacij
      cors_prikaz <- matrix(nrow = n_vars*3,
                            ncol = n_vars)
      
      vec_nrow <- seq_len(nrow(cors_prikaz))
      
      cors_prikaz[vec_nrow %% 3 == 1, ] <- m1
      cors_prikaz[vec_nrow %% 3 == 2, ] <- m2
      cors_prikaz[vec_nrow %% 3 == 0, ] <- m3
      
      colnames(cors_prikaz) <- seq_len(ncol(cors_prikaz))
      
      # dodamo še stolpca z imeni spremenljivk in legendo
      temp_mat <- matrix(nrow = n_vars*3,
                         ncol = 3)
      temp_mat[vec_nrow %% 3 == 1, 1] <- paste(paste0(seq_along(vars), "."),
                                               vars)
      temp_mat[vec_nrow %% 3 == 1, 2] <- labele
      temp_mat[vec_nrow %% 3 == 1, 3] <- paste0("r1 (", ime_baza1, ") , r2 (", ime_baza2, ")")
      temp_mat[vec_nrow %% 3 == 2, 3] <- "Abs. raz. (rel. raz.), sig."
      temp_mat[vec_nrow %% 3 == 0, 3] <- "N1 , N2"
      
      cors_prikaz <- cbind(temp_mat, cors_prikaz)
      colnames(cors_prikaz)[1:3] <- c("Spremenljivka", "Labela", " ")
      
      # izvoz Excel
      writeData(wb = wb,
                sheet = "Korelacije",
                x = cors_prikaz,
                borders = "columns",
                startRow = row_start,
                rowNames = FALSE,
                headerStyle = createStyle(textDecoration = "bold",
                                          border = c("top", "bottom", "left", "right"),
                                          halign = "center", valign = "center"))
      
      merge_min <- (vec_nrow + row_start)[vec_nrow %% 3 == 1]
      merge_max <- (vec_nrow + row_start)[vec_nrow %% 3 == 0]
      
      for(i in seq_along(vars)){
        mergeCells(wb = wb,
                   sheet = "Korelacije",
                   cols = 1, rows = c(merge_min[i], merge_max[i])) 
        
        mergeCells(wb = wb,
                   sheet = "Korelacije",
                   cols = 2, rows = c(merge_min[i], merge_max[i])) 
      }
      
      addStyle(wb = wb,
               sheet = "Korelacije",
               style = createStyle(border = "bottom"),
               rows = (vec_nrow + row_start)[vec_nrow %% 3 == 0], cols = 1:ncol(cors_prikaz),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb,
               sheet = "Korelacije",
               style = createStyle(halign = "center", valign = "center", wrapText = TRUE),
               rows = row_start:(nrow(cors_prikaz) + row_start), cols = c(1, 3:ncol(cors_prikaz)),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb,
               sheet = "Korelacije",
               style = createStyle(halign = "left", valign = "center", wrapText = TRUE),
               rows = row_start:(nrow(cors_prikaz) + row_start), cols = 2,
               gridExpand = TRUE, stack = TRUE)
      
      m4 <- m5 <- m6 <- matrix(data = FALSE,
                               nrow = n_vars,
                               ncol = n_vars)
      
      m4[lower.tri(m4)] <- abs(rel_changes) > 20 & p_values < 0.05
      m5[lower.tri(m5)] <- abs(rel_changes) > 10 & p_values < 0.05
      m6[lower.tri(m6)] <- abs(rel_changes) > 5 & p_values < 0.05
      
      row_highlight <- vec_nrow[vec_nrow %% 3 == 1] + row_start
      
      for (i in seq_len(nrow(m4))){
        for (j in seq_len(ncol(m4))){
          if(m4[i,j] == TRUE){
            addStyle(wb = wb,
                     sheet = "Korelacije",
                     style = createStyle(fgFill = "#EC8984"),
                     rows = (row_highlight[i]):(row_highlight[i] + 2), cols = j+3,
                     gridExpand = TRUE, stack = TRUE)
            
          } else if(m5[i,j] == TRUE){
            addStyle(wb = wb,
                     sheet = "Korelacije",
                     style = createStyle(fgFill = "#F5C0BD"),
                     rows = (row_highlight[i]):(row_highlight[i] + 2), cols = j+3,
                     gridExpand = TRUE, stack = TRUE)
            
          } else if(m6[i,j] == TRUE){
            addStyle(wb = wb,
                     sheet = "Korelacije",
                     style = createStyle(fgFill = "#FBE9E9"),
                     rows = (row_highlight[i]):(row_highlight[i] + 2), cols = j+3,
                     gridExpand = TRUE, stack = TRUE)
          }
        }
      }
    }
    
    # opis spremenljivk
    labele <- vapply(stevilske_spremenljivke,
                     FUN = function(x) {
                       labela1 <- attr(baza1[[x]], which = "label", exact = TRUE)
                       labela2 <- attr(baza2[[x]], which = "label", exact = TRUE)
                       labela2 <- ifelse(is.null(labela2), "", labela2)
                       ifelse(is.null(labela1), labela2, labela1)
                     },
                     FUN.VALUE = character(1))
    
    write_cors_excel(cors_baza1 = cors_baza1,
                     cors_baza2 = cors_baza2,
                     n_baza1 = n_baza1,
                     n_baza2 = n_baza2,
                     rel_changes = rel_changes,
                     p_values = p_values,
                     row_start = 8,
                     labele = labele)
    
    write_cors_excel(cors_baza1 = w_cors_baza1,
                     cors_baza2 = w_cors_baza2,
                     n_baza1 = n_baza1,
                     n_baza2 = n_baza2,
                     rel_changes = w_rel_changes,
                     p_values = w_p_values,
                     row_start = (length(stevilske_spremenljivke) * 3) + 11,
                     labele = labele)
    
    writeData(wb = wb,
              sheet = "Korelacije",
              x = "Primerjava Pearsonovih korelacijskih koeficientov (r) med 2 neodvisnima vzorcema",
              startCol = 1, startRow = 1)
    
    mergeCells(wb = wb, sheet = "Korelacije", cols = 1:(length(stevilske_spremenljivke) + 3), rows = 1)
    
    addStyle(wb = wb, sheet = "Korelacije",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 14),
             rows = 1, cols = 1,
             gridExpand = TRUE, stack = TRUE)
    
    writeData(wb = wb,
              sheet = "Korelacije",
              x = "Legenda:",
              startCol = 1, startRow = 2)
    writeData(wb = wb,
              sheet = "Korelacije",
              x = "Relativna razlika > 20% in p < 0.05",
              startCol = 1, startRow = 3)
    writeData(wb = wb,
              sheet = "Korelacije",
              x = "Relativna razlika > 10% in p < 0.05",
              startCol = 1, startRow = 4)
    writeData(wb = wb,
              sheet = "Korelacije",
              x = "Relativna razlika > 5% in p < 0.05",
              startCol = 1, startRow = 5)
    
    addStyle(wb = wb, sheet = "Korelacije",
             style = createStyle(fgFill = "#E8746E"),
             rows = 3, cols = 1,
             gridExpand = TRUE, stack = TRUE)
    addStyle(wb = wb, sheet = "Korelacije",
             style = createStyle(fgFill = "#F2AFAC"),
             rows = 4, cols = 1,
             gridExpand = TRUE, stack = TRUE)
    addStyle(wb = wb, sheet = "Korelacije",
             style = createStyle(fgFill = "#F9DCDB"),
             rows = 5, cols = 1,
             gridExpand = TRUE, stack = TRUE)
    
    writeData(wb = wb,
              sheet = "Korelacije",
              x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001",
              startCol = 1, startRow = 6)
    
    writeData(wb = wb,
              sheet = "Korelacije",
              x = "Neutežene korelacije",
              startCol = 1, startRow = 7)
    
    mergeCells(wb = wb, sheet = "Korelacije", cols = 1:(length(stevilske_spremenljivke) + 3), rows = 7)
    
    writeData(wb = wb,
              sheet = "Korelacije",
              x = "Utežene korelacije",
              startCol = 1, startRow = (length(stevilske_spremenljivke) * 3) + 10)
    
    mergeCells(wb = wb, sheet = "Korelacije",
               cols = 1:(length(stevilske_spremenljivke) + 3),
               rows = (length(stevilske_spremenljivke) * 3) + 10)
    
    addStyle(wb = wb, sheet = "Korelacije",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = c(7, ((length(stevilske_spremenljivke) * 3) + 10)), cols = 1,
             gridExpand = TRUE, stack = TRUE)
    
    setColWidths(wb = wb,
                 sheet = "Korelacije",
                 cols = 3,
                 widths = "auto")
    
    setColWidths(wb = wb,
                 sheet = "Korelacije",
                 cols = 1:2,
                 widths = c(28,15))
    
    setColWidths(wb = wb,
                 sheet = "Korelacije",
                 cols = 4:(length(stevilske_spremenljivke) + 3),
                 widths = 14)
    
    # povzetek
    frekvence_rel_razlike_neutezene <- count_rel_diff(vec = rel_changes, p_vec = p_values)
    frekvence_rel_razlike_utezene <- count_rel_diff(vec = w_rel_changes, p_vec = w_p_values)
    
    tbl_st <- data.frame("Relativne razlike" = c("> 20%", "(10% - 20%]", "[5% - 10%]", "< 5%"),
                         # neutežene korelacije
                         "f" = frekvence_rel_razlike_neutezene$sums,
                         "%" = frekvence_rel_razlike_neutezene$sums/sum(frekvence_rel_razlike_neutezene$sums),
                         "Kumul f" = frekvence_rel_razlike_neutezene$cumsums,
                         "Kumul %" = frekvence_rel_razlike_neutezene$cumsums/sum(frekvence_rel_razlike_neutezene$sums),
                         "f*" = frekvence_rel_razlike_neutezene$p_sums,
                         "% od vseh spremenljivk" = frekvence_rel_razlike_neutezene$p_sums/sum(frekvence_rel_razlike_neutezene$sums),
                         "% od relativne razlike" = frekvence_rel_razlike_neutezene$p_sums/frekvence_rel_razlike_neutezene$sums,
                         "Kumul f*" = frekvence_rel_razlike_neutezene$p_cumsums,
                         "Kumul %*" = frekvence_rel_razlike_neutezene$p_cumsums/sum(frekvence_rel_razlike_neutezene$sums),
                         # utežene korelacije
                         "f" = frekvence_rel_razlike_utezene$sums,
                         "%" = frekvence_rel_razlike_utezene$sums/sum(frekvence_rel_razlike_utezene$sums),
                         "Kumul f" = frekvence_rel_razlike_utezene$cumsums,
                         "Kumul %" = frekvence_rel_razlike_utezene$cumsums/sum(frekvence_rel_razlike_utezene$sums),
                         "f*" = frekvence_rel_razlike_utezene$p_sums,
                         "% od vseh spremenljivk" = frekvence_rel_razlike_utezene$p_sums/sum(frekvence_rel_razlike_utezene$sums),
                         "% od relativne razlike" = frekvence_rel_razlike_utezene$p_sums/frekvence_rel_razlike_utezene$sums,
                         "Kumul f*" = frekvence_rel_razlike_utezene$p_cumsums,
                         "Kumul %*" = frekvence_rel_razlike_utezene$p_cumsums/sum(frekvence_rel_razlike_utezene$sums),
                         check.names = FALSE)
    
    tbl_st[is.na(tbl_st)] <- 0
    
    writeData(wb = wb,
              sheet = "Povzetek",
              x = tbl_st,
              borders = "all", startRow = 4, 
              headerStyle = createStyle(textDecoration = "bold",
                                        border = c("top", "bottom", "left", "right"),
                                        halign = "center", valign = "center", wrapText = TRUE))
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Povzetek - primerjava Pearsonovih korelacijskih koeficientov med 2 neodvisnima vzorcema", 
              startCol = 2, startRow = 1)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:19, rows = 1)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Neutežene korelacije", startCol = 2, startRow = 2)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:10, rows = 2)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Utežene korelacije", startCol = 11, startRow = 2)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 11:19, rows = 2)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Primerjave korelacij", startCol = 2, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:5, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Primerjave korelacij", startCol = 11, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 11:14, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne razlike korelacij (p < 0.05)",
              startCol = 6, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 6:10, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne razlike korelacij (p < 0.05)",
              startCol = 15, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 15:19, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh primerjav korelacij:", sum(frekvence_rel_razlike_neutezene$sums)),
                          paste("Št. statistično značilnih razlik med korelacijami (p < 0.05) - neuteženi podatki:", sum(frekvence_rel_razlike_neutezene$p_sums)),
                          paste("Št. statistično značilnih razlik med korelacijami (p < 0.05) - uteženi podatki:", sum(frekvence_rel_razlike_utezene$p_sums)))),
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
             style = createStyle(numFmt = "0%"),
             rows = 5:8, cols = c(3, 5, 7, 8, 10, 12, 14, 16, 17, 19),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(halign = "center"),
             rows = 5:8, cols = 1:19,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(fgFill = "#DAEEF3"), rows = 3:8, cols = 2:10,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(fgFill = "#FDE9D9"), rows = 3:8, cols = 11:19,
             gridExpand = TRUE, stack = TRUE)
    
    setColWidths(wb = wb, sheet = "Povzetek", cols = 1:19,
                 widths = c(16, rep(c(5, 6, 8, 8, 5, 21, 21, 8, 9), 2)))
    
    setRowHeights(wb = wb, sheet = "Povzetek", rows = 1:4, heights = c(20, 20, 25, 44))
    
    if(warning_counter == FALSE){
      removeWorksheet(wb = wb, sheet = "Opozorila")
    }
    
    saveWorkbook(wb = wb, file = file, overwrite = TRUE) 
  }
}