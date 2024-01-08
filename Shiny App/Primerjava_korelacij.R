# funkcija ki prešteje št. korelacij glede na Cohenov učinek q (1988) in št. stat. značilnih spremenljivk
# The value q is the estimate of the effect size. The following intervals were proposed by Cohen to interpret these values: 
# q < 0.1, no effect; 0.1 ≤ q < 0.3, small effect; 0.3 ≤ q < 0.5, medium effect; q ≥ 0.5, large effect.
# https://stats.stackexchange.com/questions/99741/how-to-compare-the-strength-of-two-pearson-correlations#99747
count_cohens_q <- function(vec, p_vec) {
  vec <- vec[!is.na(vec)] # q values
  p_vec <- p_vec[!is.na(p_vec)]
  vec <- abs(vec)
  
  frek <- c(sum(vec >= 0.5), sum(vec >= 0.3 & vec < 0.5), sum(vec >= 0.1 & vec < 0.3), sum(vec < 0.1))
  
  p_frek <- c(sum(vec >= 0.5 & p_vec < 0.05), sum(vec >= 0.3 & vec < 0.5 & p_vec < 0.05), sum(vec >= 0.1 & vec < 0.3 & p_vec < 0.05), sum(vec < 0.1 & p_vec < 0.05))
  
  list(sums = frek,
       cumsums = cumsum(frek),
       p_sums = p_frek,
       p_cumsums = cumsum(p_frek))
}

# glavna funkcija za izvoz tabel v Excel
izvoz_excel_korelacije <- function(baza1 = NULL,
                                   baza2 = NULL,
                                   ime_baza1 = "baza 1",
                                   ime_baza2 = "baza 2",
                                   utezi1 = NULL,
                                   utezi2 = NULL,
                                   stevilske_spremenljivke = NULL,
                                   nominalne_spremenljivke = NULL,
                                   file){
  wb <- createWorkbook()
  addWorksheet(wb = wb, sheetName = "Opozorila", gridLines = FALSE)
  addWorksheet(wb = wb, sheetName = "Povzetek", gridLines = FALSE)
  
  use_weights <- FALSE
  if(!is.null(utezi1) && !is.null(utezi2)) use_weights <- TRUE
  
  warning_counter <- FALSE
  
  stevilske_spremenljivke <- c(stevilske_spremenljivke, nominalne_spremenljivke)
  
  imena_st_1 <- names(baza1)[names(baza1) %in% stevilske_spremenljivke]
  imena_st_2 <- names(baza2)[names(baza2) %in% stevilske_spremenljivke]
  
  # preverimo ujemanje imen spremenljivk v obeh bazah
  non_st_spr <- union(setdiff(imena_st_1, imena_st_2), setdiff(imena_st_2, imena_st_1))
  
  if(length(non_st_spr) > 0){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,1),
              x = paste("Imena podanih spremenljivk se ne ujemajo v obeh bazah. To so spremenljivke", paste(non_st_spr, collapse = ", "), "in so bile zato odstranjene iz analiz."))
    
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
              x = paste("Kategorije se pri spremenljivki/ah", paste(stevilske_spremenljivke[!indeksi], collapse = ", "), "ne ujemajo. Te spremenljivke niso bile odstranjene iz analiz, vendar svetujemo previdnost pri analizi."))
    
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
              x = paste("Spremenljivke", paste(stevilske_spremenljivke[indeksi_variances], collapse = ", "), "so bile odstranjene, saj so konstante (imajo ničelno varianco) v eni ali obeh bazah."))
    
    warning_counter <- TRUE
    
    stevilske_spremenljivke <- stevilske_spremenljivke[!indeksi_variances]
  }
  
  # obdržimo le dihotomne nominalne spremenljivke (ker se lahko izračuna korelacija le med razmernostno in dihotomno nominalno
  # in med dvema dihotomnima (Phi koeficient))
  nom_dihot <- vapply(nominalne_spremenljivke, FUN = function(x){
    (length(tabulate(baza1[[x]])) == 2) && (length(tabulate(baza2[[x]])) == 2)
  }, FUN.VALUE = logical(1), USE.NAMES = TRUE)
  
  nominalne_spremenljivke <- names(nom_dihot[nom_dihot == FALSE])
  
  if(any(!nom_dihot)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,5),
              x = paste("Izračun Pearsonove korelacije je mogoč le za dihotomne nominalne spremenljivke. Spremenljivke", paste(nominalne_spremenljivke, collapse = ", "), "niso dihotomne so bile zato odstranjene iz analiz."))
    
    warning_counter <- TRUE
  }
  
  stevilske_spremenljivke[stevilske_spremenljivke %in% nominalne_spremenljivke] <- NA
  stevilske_spremenljivke <- stevilske_spremenljivke[!is.na(stevilske_spremenljivke)]
  
  baza1 <- baza1[,stevilske_spremenljivke, drop = FALSE]
  baza2 <- baza2[,stevilske_spremenljivke, drop = FALSE]
  
  if(length(stevilske_spremenljivke) > 1){
    
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
    
    # primerjava korelacij (p vrednosti, Cohenov q - velikost učinka)
    p_values <- vapply(seq_along(cors_baza1), FUN = function(i){
      if(anyNA(c(cors_baza1[[i]], cors_baza2[[i]]))){
        NA
      } else {
        test <- cocor::cocor.indep.groups(r1.jk = cors_baza1[[i]],
                                          r2.hm = cors_baza2[[i]],
                                          n1 = n_baza1[[i]],
                                          n2 = n_baza2[[i]],
                                          test = "fisher1925")
        test@fisher1925$p.value
      }
      
    }, FUN.VALUE = numeric(1))
    
    # Fisherjeva transformacija: atanh == 1/2*(log(1+r/1-r))
    cohens_q <- abs(atanh(cors_baza1) - atanh(cors_baza2))
    
    write_cors_excel <- function(cors_baza1,
                                 cors_baza2,
                                 n_baza1,
                                 n_baza2,
                                 cohens_q,
                                 p_values,
                                 vars = stevilske_spremenljivke,
                                 row_start,
                                 labele,
                                 table_name,
                                 sheet_name){
      
      addWorksheet(wb = wb, sheetName = sheet_name, gridLines = FALSE)
      
      n_vars <- length(vars)
      
      m1 <- m2 <- m3 <- matrix(nrow = n_vars,
                               ncol = n_vars)
      
      # korelacije
      m1[lower.tri(m1)] <- paste(round(cors_baza1, 2),
                                 ",",
                                 round(cors_baza2,2))
      
      # absolutne razlike, Cohenov q - velikost učinka, p-vrednosti
      m2[lower.tri(m2)] <- paste(round(cors_baza2 - cors_baza1, 2),
                                 paste0("(", round(cohens_q, 2), ")"),
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
      temp_mat[vec_nrow %% 3 == 2, 3] <- "Abs. raz. (Cohenov q), sig."
      temp_mat[vec_nrow %% 3 == 0, 3] <- "N1 , N2"
      
      cors_prikaz <- cbind(temp_mat, cors_prikaz)
      colnames(cors_prikaz)[1:3] <- c("Spremenljivka", "Labela", " ")
      
      # izvoz Excel
      writeData(wb = wb,
                sheet = sheet_name,
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
                   sheet = sheet_name,
                   cols = 1, rows = c(merge_min[i], merge_max[i])) 
        
        mergeCells(wb = wb,
                   sheet = sheet_name,
                   cols = 2, rows = c(merge_min[i], merge_max[i])) 
      }
      
      addStyle(wb = wb,
               sheet = sheet_name,
               style = createStyle(border = "bottom"),
               rows = (vec_nrow + row_start)[vec_nrow %% 3 == 0], cols = 1:ncol(cors_prikaz),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb,
               sheet = sheet_name,
               style = createStyle(halign = "center", valign = "center", wrapText = TRUE),
               rows = row_start:(nrow(cors_prikaz) + row_start), cols = c(1, 3:ncol(cors_prikaz)),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb,
               sheet = sheet_name,
               style = createStyle(halign = "left", valign = "center", wrapText = TRUE),
               rows = row_start:(nrow(cors_prikaz) + row_start), cols = 2,
               gridExpand = TRUE, stack = TRUE)
      
      m4 <- m5 <- m6 <- matrix(data = FALSE,
                               nrow = n_vars,
                               ncol = n_vars)
      
      m4[lower.tri(m4)] <- cohens_q >= 0.5 & p_values < 0.05
      m4[is.na(m4)] <- FALSE
      m5[lower.tri(m5)] <- cohens_q >= 0.3 & cohens_q < 0.5 & p_values < 0.05
      m5[is.na(m5)] <- FALSE
      m6[lower.tri(m6)] <- cohens_q >= 0.1 & cohens_q < 0.3 & p_values < 0.05
      m6[is.na(m6)] <- FALSE
      
      row_highlight <- vec_nrow[vec_nrow %% 3 == 1] + row_start
      
      for (i in seq_len(nrow(m4))){
        for (j in seq_len(ncol(m4))){
          if(m4[i,j] == TRUE){
            addStyle(wb = wb,
                     sheet = sheet_name,
                     style = createStyle(fgFill = "#EC8984"),
                     rows = (row_highlight[i]):(row_highlight[i] + 2), cols = j+3,
                     gridExpand = TRUE, stack = TRUE)
            
          } else if(m5[i,j] == TRUE){
            addStyle(wb = wb,
                     sheet = sheet_name,
                     style = createStyle(fgFill = "#F5C0BD"),
                     rows = (row_highlight[i]):(row_highlight[i] + 2), cols = j+3,
                     gridExpand = TRUE, stack = TRUE)
            
          } else if(m6[i,j] == TRUE){
            addStyle(wb = wb,
                     sheet = sheet_name,
                     style = createStyle(fgFill = "#FBE9E9"),
                     rows = (row_highlight[i]):(row_highlight[i] + 2), cols = j+3,
                     gridExpand = TRUE, stack = TRUE)
          }
        }
      }
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Primerjava Pearsonovih korelacijskih koeficientov (r) med 2 neodvisnima vzorcema",
                startCol = 1, startRow = 1)
      
      # mergeCells(wb = wb, sheet = sheet_name, cols = 1:(length(stevilske_spremenljivke) + 3), rows = 1)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(textDecoration = "bold",
                                   fontSize = 14),
               rows = 1, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Legenda:",
                startCol = 1, startRow = 2)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Visok učinek (q ≥ 0.5) in p < 0.05",
                startCol = 1, startRow = 3)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Zmeren učinek (0.3 ≤ q < 0.5) in p < 0.05",
                startCol = 1, startRow = 4)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Šibek učinek (0.1 ≤ q < 0.3) in p < 0.05",
                startCol = 1, startRow = 5)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#E8746E"),
               rows = 3, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#F2AFAC"),
               rows = 4, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#F9DCDB"),
               rows = 5, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001",
                startCol = 1, startRow = 6)
      
      if(anyNA(c(cors_baza1, cors_baza2))){
        writeData(wb = wb,
                  sheet = sheet_name,
                  x = "NA: korelacije ni mogoče izračunati, ker je vsaj ena od spremenljivk konstanta.",
                  startCol = 1, startRow = 7)
      }
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = table_name,
                startCol = 4, startRow = 8)
      
      # mergeCells(wb = wb, sheet = sheet_name,
      #            cols = 1:(length(stevilske_spremenljivke) + 3), rows = 8)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(textDecoration = "bold",
                                   halign = "left", valign = "center",
                                   fontSize = 12),
               rows = 8, cols = 4,
               gridExpand = TRUE, stack = TRUE)
      
      setColWidths(wb = wb,
                   sheet = sheet_name,
                   cols = 1:2,
                   widths = c(32, 15))
      
      setColWidths(wb = wb,
                   sheet = sheet_name,
                   cols = 3,
                   widths = "auto")
      
      setColWidths(wb = wb,
                   sheet = sheet_name,
                   cols = 4:(length(stevilske_spremenljivke) + 3),
                   widths = 14)
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
                     cohens_q = cohens_q,
                     p_values = p_values,
                     row_start = 9,
                     labele = labele,
                     table_name = "Neutežene korelacije",
                     sheet_name = "Neutežene korelacije")
    
    # če so izbrane uteži, se izračunajo še utežene korelacije
    if(use_weights){
      # utežene korelacije v 1. bazi
      w_cors_baza1 <- weights::wtd.cors(x = baza1, weight = utezi1)
      w_cors_baza1 <- w_cors_baza1[lower.tri(w_cors_baza1)]
      
      # utežene korelacije v 2. bazi
      w_cors_baza2 <- weights::wtd.cors(x = baza2, weight = utezi2)
      w_cors_baza2 <- w_cors_baza2[lower.tri(w_cors_baza2)]
      
      w_p_values <- vapply(seq_along(w_cors_baza1), FUN = function(i){
        if(anyNA(c(w_cors_baza1[[i]], w_cors_baza2[[i]]))){
          NA
        } else {
          test <- cocor::cocor.indep.groups(r1.jk = w_cors_baza1[[i]],
                                            r2.hm = w_cors_baza2[[i]],
                                            n1 = n_baza1[[i]],
                                            n2 = n_baza2[[i]],
                                            test = "fisher1925")
          test@fisher1925$p.value
        }
      }, FUN.VALUE = numeric(1))
      
      w_cohens_q <- abs(atanh(w_cors_baza1) - atanh(w_cors_baza2))
      
      write_cors_excel(cors_baza1 = w_cors_baza1,
                       cors_baza2 = w_cors_baza2,
                       n_baza1 = n_baza1,
                       n_baza2 = n_baza2,
                       cohens_q = w_cohens_q,
                       p_values = w_p_values,
                       row_start = 9,
                       labele = labele,
                       table_name = "Utežene korelacije",
                       sheet_name = "Utežene korelacije")
    }
    
    # povzetek
    frekvence_rel_razlike_neutezene <- count_cohens_q(vec = cohens_q, p_vec = p_values)
    
    tbl_st <- data.frame("Velikost učinka razlike med korelacijama (Cohenov q)" = c("Visok učinek (q ≥ 0.5)",
                                                                                    "Zmeren učinek (0.3 ≤ q < 0.5)",
                                                                                    "Šibek učinek (0.1 ≤ q < 0.3)",
                                                                                    "Brez učinka (q < 0.1)"),
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
                         check.names = FALSE)
    
    if(use_weights){
      frekvence_rel_razlike_utezene <- count_cohens_q(vec = w_cohens_q, p_vec = w_p_values)
      
      w_tbl_st <- data.frame(
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
      
      tbl_st <- cbind(tbl_st, w_tbl_st)
    }
    
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
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:ncol(tbl_st), rows = 1)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Neutežene korelacije", startCol = 2, startRow = 2)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:10, rows = 2)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Primerjave korelacij", startCol = 2, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:5, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne razlike korelacij (p < 0.05)",
              startCol = 6, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 6:10, rows = 3)
    
    if(use_weights){
      writeData(wb = wb, sheet = "Povzetek",
                x = "Utežene korelacije", startCol = 11, startRow = 2)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 11:19, rows = 2)
      
      writeData(wb = wb, sheet = "Povzetek",
                x = "Primerjave korelacij", startCol = 11, startRow = 3)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 11:14, rows = 3)
      
      writeData(wb = wb, sheet = "Povzetek",
                x = "Od tega statistično značilne razlike korelacij (p < 0.05)",
                startCol = 15, startRow = 3)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 15:19, rows = 3)
    }
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh primerjav korelacij:", sum(frekvence_rel_razlike_neutezene$sums)),
                          paste("Št. statistično značilnih razlik med korelacijami (p < 0.05) - neuteženi podatki:", sum(frekvence_rel_razlike_neutezene$p_sums)))),
              startCol = 1, startRow = 10, rowNames = FALSE, colNames = FALSE)
    
    if(use_weights){
      writeData(wb = wb, sheet = "Povzetek",
                x = paste("Št. statistično značilnih razlik med korelacijami (p < 0.05) - uteženi podatki:", sum(frekvence_rel_razlike_utezene$p_sums)),
                startCol = 1, startRow = 12, rowNames = FALSE, colNames = FALSE)
    }
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 1:2, cols = 1:ncol(tbl_st),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 border = c("top", "bottom", "left", "right")),
             rows = 3, cols = 1:ncol(tbl_st),
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
    
    if(use_weights){
      addStyle(wb = wb, sheet = "Povzetek",
               style = createStyle(fgFill = "#FDE9D9"), rows = 3:8, cols = 11:19,
               gridExpand = TRUE, stack = TRUE)
    }
    
    setColWidths(wb = wb, sheet = "Povzetek", cols = 1:19,
                 widths = c(25, rep(c(5, 6, 8, 8, 5, 21, 21, 8, 9), 2)))
    
    setRowHeights(wb = wb, sheet = "Povzetek", rows = 1:4, heights = c(20, 20, 25, 44))
  }
  
  if(warning_counter == FALSE){
    writeData(wb = wb, sheet = "Opozorila", startCol = 1, startRow = 1,
              x = "Ni opozoril")
  }
  
  saveWorkbook(wb = wb, file = file, overwrite = TRUE) 
}