# TODO
# uteži naj bodo opcijske kot pri korelacijah
# korelacije primerjava vzorec podvzorec

# https://stackoverflow.com/questions/70626996/r-how-to-conduct-two-sample-t-test-with-two-different-survey-designs

# pomožne funkcije
load_to_environment <- function(RData, env = new.env()) {
  load(RData, env)
  return(env[[names(env)[1]]])
}

# funkcija ki prešteje št. relativnih razlik glede na intervale in št. stat. značilnih spremenljivk
count_rel_diff <- function(vec, p_vec) {
  p_vec[is.na(p_vec)] <- 1
  vec <- abs(vec)
  
  frek <- c(sum(vec > 20), sum(vec > 10 & vec <= 20), sum(vec >= 5 & vec <= 10), sum(vec < 5))
  
  p_frek <- c(sum(vec > 20 & p_vec < 0.05), sum(vec > 10 & vec <= 20 & p_vec < 0.05), sum(vec >= 5 & vec <= 10 & p_vec < 0.05), sum(vec < 5 & p_vec < 0.05))
  
  list(sums = frek,
       cumsums = cumsum(frek),
       p_sums = p_frek,
       p_cumsums = cumsum(p_frek))
}

# funkcija za utežene frekvence
weighted_table <- function(x, weights) {
  use <- !is.na(x)
  x <- x[use]
  weights <- weights[use]
  weights <- weights/mean(weights, na.rm = TRUE)
  
  vapply(split(weights, x), sum, 0, na.rm = TRUE)
}

# funkcija za izračun testne statistike (Welchev t-test)
wtd_t_test <- function(x,
                       y,
                       weights_x = NULL,
                       weights_y = NULL,
                       prop = FALSE,
                       se_calculation,
                       survey_design1,
                       survey_design2){

  if(is.null(weights_x)) weights_x <- rep(1, length(x))

  if(is.null(weights_y)) weights_y <- rep(1, length(y))
  
  n1 <- sum(!is.na(x))
  n2 <- sum(!is.na(y))
  
  mu1 <- weighted.mean(x = x, w = weights_x, na.rm = TRUE)
  mu2 <- weighted.mean(x = y, w = weights_y, na.rm = TRUE)
  
  if(se_calculation == "taylor_se"){
    use_x <- !is.na(x)
    use_y <- !is.na(y)
    
    x <- x[use_x]
    y <- y[use_y]
    weights_x <- weights_x[use_x]
    weights_y <- weights_y[use_y]
    
    # SE^2
    se1_2 <- (n1/((n1 - 1) * sum(weights_x)^2)) * sum(weights_x^2 * (x - mu1)^2)
    se2_2 <- (n2/((n2 - 1) * sum(weights_y)^2)) * sum(weights_y^2 * (y - mu2)^2)
  }
  
  if(se_calculation == "survey_se"){
    se1_2 <- SE(svymean(x = x, design = survey_design1, na.rm = TRUE))^2
    se2_2 <- SE(svymean(x = y, design = survey_design2, na.rm = TRUE))^2
  }

  t <- (mu2 - mu1)/(sqrt(se1_2 + se2_2))

  if(prop == TRUE){
    p <- pnorm(q = abs(t), lower.tail = FALSE) * 2
  } else {
    df <- (se1_2 + se2_2)^2/(se1_2^2/(n1 - 1) + se2_2^2/(n2 - 1))
    p <- pt(q = abs(t), df = df, lower.tail = FALSE) * 2
  }

  c(povp1 = mu1, povp2 = mu2, diff = mu2 - mu1, t = t, p = p)
}


# funkcija za izračun testne statistike za primerjavo vzorca s podvzorcem
wtd_t_test_sample_subsample <- function(sample,
                                        subsample, # podvzorec 1
                                        weights_sample = NULL,
                                        weights_subsample){
  
  subsample_index <- !is.na(weights_subsample) 
  subsample2 <- sample[!subsample_index] # podvzorec 2
  indicator <- FALSE
  
  if(is.null(weights_sample)){
    weights_sample <- rep(1, length(sample))
    weights_subsample <- rep(1, length(sample))
    indicator <- TRUE
  }
  
  weights_subsample <- weights_subsample[subsample_index]
  weights_sample <- weights_sample/mean(weights_sample)
  
  n_sample <- sum(!is.na(sample)) # n celega vzorca
  n_subsample <- sum(!is.na(subsample)) # n podvzorca
  
  # delež vsebovanosti prvega podvzorca v celotnem vzorcu
  w <- sum(weights_sample[subsample_index][!is.na(sample[subsample_index])])/sum(weights_sample[!is.na(sample)])
  # w <- n_subsample/n_sample 
  
  # povprečji
  mu_sample <- weighted.mean(x = sample, w = weights_sample, na.rm = TRUE)
  mu_subsample <- weighted.mean(x = subsample, w = weights_subsample, na.rm = TRUE)
  
  if(w > 0 && w <= 1){
    use_subsample <- !is.na(subsample)
    use_subsample2 <- !is.na(subsample2)
    
    subsample <- subsample[use_subsample]
    subsample2 <- subsample2[use_subsample2]
    
    weights_subsample1 <- weights_sample[subsample_index] # uteži celega vzorca, izbran podvzorec 1
    weights_subsample1 <- weights_subsample1[use_subsample]/mean(weights_subsample1[use_subsample])
    
    weights_subsample2 <- weights_sample[!subsample_index] # uteži celega vzorca, izbran podvzorec 2
    weights_subsample2 <- weights_subsample2[use_subsample2]/mean(weights_subsample2[use_subsample2])
    
    weights_subsample <- weights_subsample[use_subsample]/mean(weights_subsample[use_subsample]) # uteži, utežene posebej za podvzorec 1
    
    taylor_se <- function(n, weights, x){
      mu <- weighted.mean(x, weights)
      (n/((n - 1) * sum(weights)^2)) * sum(weights^2 * (x - mu)^2)
    }
    
    # SE^2
    # se_2 <- sum(w^2 * taylor_se(n_subsample,
    #                             weights_subsample1,
    #                             subsample) , (1 - w)^2 * taylor_se(n_sample - n_subsample,
    #                                                                weights_subsample2,
    #                                                                subsample2) , taylor_se(n_subsample,
    #                                                                                        weights_subsample,
    #                                                                                        subsample) , - (2 * w * (cov(weights_subsample1*subsample, weights_subsample*subsample)/n_subsample)), na.rm = TRUE)
    
    mu_x <- weighted.mean(subsample, weights_subsample)
    mu_y <- weighted.mean(subsample, weights_subsample1)
    if(indicator == TRUE){
      cov_xy <- sum((subsample - mu_x) * (subsample - mu_y) * weights_subsample * weights_subsample1) / ((n_subsample - 1) * n_subsample) # nepristranska cenilka (n-1)
    } else {
      cov_xy <- sum((subsample - mu_x) * (subsample - mu_y) * weights_subsample * weights_subsample1) / (sum(weights_subsample) * sum(weights_subsample1))
    }

    se_2 <- sum(w^2 * taylor_se(n_subsample,
                                weights_subsample1,
                                subsample) , (1 - w)^2 * taylor_se(n_sample - n_subsample,
                                                                   weights_subsample2,
                                                                   subsample2) , taylor_se(n_subsample,
                                                                                           weights_subsample,
                                                                                           subsample) , - (2 * w * cov_xy), na.rm = TRUE)
    
    z <- (mu_subsample - mu_sample)/(sqrt(se_2))
    
    p <- pnorm(q = abs(z), lower.tail = FALSE) * 2
    
    c(povp1 = mu_sample,
      povp2 = mu_subsample,
      diff = mu_subsample - mu_sample,
      w = w,
      z = z,
      p = p)
    
  } else {
    c(povp1 = mu_sample,
      povp2 = mu_subsample,
      diff = mu_subsample - mu_sample,
      w = w,
      z = NA,
      p = NA)
  }
}

# glavna funkcija za izvoz tabel v Excel
izvoz_excel_tabel <- function(baza1 = NULL,
                              baza2 = NULL,
                              ime_baza1 = "baza 1",
                              ime_baza2 = "baza 2",
                              utezi1 = NULL,
                              utezi2 = NULL,
                              stevilske_spremenljivke = NULL,
                              nominalne_spremenljivke = NULL,
                              se_calculation,
                              survey_design1 = NULL,
                              survey_design2 = NULL,
                              file,
                              compare_sample_subsample = FALSE) {
  
  if(se_calculation == "survey_se"){
    utezi1 <- weights(survey_design1)
    utezi2 <- weights(survey_design2)
  }
  
  if(compare_sample_subsample == TRUE){
    baza2 <- baza1[!is.na(utezi2), , drop = FALSE]
  }
  
  if(!is.data.frame(baza1) || !is.data.frame(baza2)){
    stop("Baza mora biti SPSS podatkovni okvir.")
  }
  
  if(is.null(utezi1) || is.null(utezi2)){
    stop("Določite uteži!")
  }
  
  if(!is.numeric(utezi1) || !is.numeric(utezi2)){
    stop("Uteži morajo biti v obliki numeričnega vektorja.")
  }
  
  if(anyNA(utezi1)){
    stop("Uteži v 1. bazi vsebujejo manjkajoče vrednosti.")
  }
  
  if(anyNA(utezi2) && compare_sample_subsample == FALSE){
    stop("Uteži v 2. bazi vsebujejo manjkajoče vrednosti.")
  }
  
  # if(!is.null(stevilske_spremenljivke) && !is.character(stevilske_spremenljivke)){
  #   stop("Podane številske spremenljivke morajo biti v obliki character vector.")
  # }
  # 
  # if(!is.null(nominalne_spremenljivke) && !is.character(nominalne_spremenljivke)){
  #   stop("Podane nominalne spremenljivke morajo biti v obliki character vector.")
  # }
  
  wb <- createWorkbook()
  
  addWorksheet(wb = wb, sheetName = "Opozorila", gridLines = FALSE)
  
  addWorksheet(wb = wb, sheetName = "Povzetek", gridLines = FALSE)
  
  warning_counter <- FALSE
  
  if(!is.null(stevilske_spremenljivke)){
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
    
    # pretvorimo spss manjkajoče vrednosti v prave manjkajoče (NA) in uredimo bazi v isti vrstni red spremenljivk
    baza1_na <- baza1[,stevilske_spremenljivke, drop = FALSE]
    baza2_na <- baza2[,stevilske_spremenljivke, drop = FALSE]
    
    # preverimo, da je isto število kategorij v obeh bazah
    levels1 <- lapply(stevilske_spremenljivke, function(x) attr(baza1_na[[x]], "labels", exact = TRUE))
    levels2 <- lapply(stevilske_spremenljivke, function(x) attr(baza2_na[[x]], "labels", exact = TRUE))
    
    indeksi <- vapply(seq_along(stevilske_spremenljivke), function(i) all(levels1[[i]] == levels2[[i]]), FUN.VALUE = logical(1))
    
    if(!all(indeksi)){
      writeData(wb = wb, sheet = "Opozorila", xy = c(1,2),
                x = paste("Kategorije se pri številski/h spremenljivki/ah", paste(stevilske_spremenljivke[!indeksi], collapse = ", "), "ne ujemajo. Te spremenljivke niso bile odstranjene iz analiz, vendar svetujemo previdnost pri analizi."))
      
      warning_counter <- TRUE
    }
    
    # preverimo, da so spremenljivke res številske in neštevilske odstranimo
    numeric_variables <- vapply(stevilske_spremenljivke, function(x){
      t1 <- is.numeric(baza1_na[[x]]) | is.integer(baza1_na[[x]])
      t2 <- is.numeric(baza2_na[[x]]) | is.integer(baza2_na[[x]])
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
      v1 <- var(baza1_na[[x]], na.rm = TRUE)
      v1 <- ifelse(is.na(v1), 0, v1)
      v2 <- var(baza2_na[[x]], na.rm = TRUE)
      v2 <- ifelse(is.na(v2), 0, v2)
      any(v1 == 0, v2 == 0)
    }, FUN.VALUE = logical(1), USE.NAMES = FALSE)
    
    if(any(indeksi_variances)){
      writeData(wb = wb, sheet = "Opozorila", xy = c(1,4),
                x = paste("Številske spremenljivke", paste(stevilske_spremenljivke[indeksi_variances], collapse = ", "), "so bile odstranjene, saj so konstante (imajo ničelno varianco) v eni ali obeh bazah."))
      
      warning_counter <- TRUE
      
      stevilske_spremenljivke <- stevilske_spremenljivke[!indeksi_variances]
    }
    
    baza1_na <- baza1_na[,stevilske_spremenljivke, drop = FALSE]
    baza2_na <- baza2_na[,stevilske_spremenljivke, drop = FALSE]
    
    # tabela
    tabela_st <- data.frame("Spremenljivka" = stevilske_spremenljivke,
                            "Labela" = vapply(stevilske_spremenljivke,
                                              FUN = function(x) {
                                                labela1 <- attr(baza1[[x]], which = "label", exact = TRUE)
                                                labela2 <- attr(baza2[[x]], which = "label", exact = TRUE)
                                                labela2 <- ifelse(is.null(labela2), "", labela2)
                                                ifelse(is.null(labela1), labela2, labela1)
                                                },
                                              FUN.VALUE = character(1)))
    if(nrow(tabela_st) > 0){
      
      # tabela_st[["Min"]] <- pmin(sapply(baza1_na, min, na.rm = TRUE),
      #                            sapply(baza2_na, min, na.rm = TRUE))

      tabela_st[["Min"]] <- sapply(stevilske_spremenljivke, function(x){
        labele <- attr(x = baza1[[x]], which = "labels", exact = TRUE)
        if(is.null(labele)) labele <- attr(x = baza2[[x]], which = "labels", exact = TRUE) else labele
        if(!is.null(labele)){
          string <- stringr::str_squish(paste(unname(labele[1]), "-", names(labele[1])))
          paste(stringr::str_unique(stringr::str_split_1(string, " ")), collapse = " ")
        } else {
          # vrne min od tiste baze, kjer je manjša vrednost minimum
          pmin(min(baza1_na[[x]], na.rm = TRUE), min(baza2_na[[x]], na.rm = TRUE))
        }
      }, USE.NAMES = FALSE)
      
      # tabela_st[["Maks"]] <- pmax(sapply(baza1_na, max, na.rm = TRUE),
      #                            sapply(baza2_na, max, na.rm = TRUE))
      
      tabela_st[["Maks"]] <- sapply(stevilske_spremenljivke, function(x){
        labele <- attr(x = baza1[[x]], which = "labels", exact = TRUE)
        if(is.null(labele)) labele <- attr(x = baza2[[x]], which = "labels", exact = TRUE) else labele
        if(!is.null(labele)){
          string <- stringr::str_squish(paste(unname(labele[length(labele)]), "-", names(labele[length(labele)])))
          paste(stringr::str_unique(stringr::str_split_1(string, " ")), collapse = " ")
        } else {
          # vrne max od tiste baze, kjer je višja vrednost maksimum
          pmax(max(baza1_na[[x]], na.rm = TRUE), max(baza2_na[[x]], na.rm = TRUE))
        }
      }, USE.NAMES = FALSE)
      
      ## Neutežene statistike ----------------------------------------------------
      
      # N neutežen baza 1
      tabela_st[[paste0("N - ", ime_baza1)]] <- colSums(!is.na(baza1_na))
      
      # N neutežen baza 2
      tabela_st[[paste0("N - ", ime_baza2)]] <- colSums(!is.na(baza2_na))
      
      # Neutežen Welchev t-test
      # statistike2 <- lapply(stevilske_spremenljivke, FUN = function(x){
      # test <- t.test(x = baza1_na[[x]], y = baza2_na[[x]],
      #                paired = FALSE, var.equal = FALSE)
      # 
      #   c("povp1" = test$estimate[["mean of x"]],
      #     "povp2" = test$estimate[["mean of y"]],
      #     "diff"  = test$estimate[["mean of y"]] - test$estimate[["mean of x"]],
      #     "t"     = -test$statistic[[1]],
      #     "p"     = test$p.value)
      # })
      # all.equal(statistike, statistike2)
      
      if(compare_sample_subsample == FALSE) {
        statistike <- lapply(stevilske_spremenljivke, FUN = function(x) wtd_t_test(x = baza1_na[[x]], y = baza2_na[[x]], se_calculation = "taylor_se"))
      } else if(compare_sample_subsample == TRUE) {
        statistike <- lapply(stevilske_spremenljivke, FUN = function(x) wtd_t_test_sample_subsample(sample = baza1_na[[x]], subsample = baza2_na[[x]], weights_subsample = utezi2))
      }
      
      # Neuteženo povprečje baza 1
      tabela_st[[paste0("Povprečje - ", ime_baza1)]] <- sapply(statistike, function(x) x[["povp1"]])
      
      # Neuteženo povprečje baza 2
      tabela_st[[paste0("Povprečje - ", ime_baza2)]] <- sapply(statistike, function(x) x[["povp2"]])
      
      # Absolutna razlika neuteženih povprečji
      tabela_st[["Razlika v povprečjih - absolutna"]] <- sapply(statistike, function(x) x[["diff"]])
      
      # Relativna razlika neuteženih povprečji
      tabela_st[["Razlika v povprečjih - relativna (%)"]] <- (tabela_st[["Razlika v povprečjih - absolutna"]]/tabela_st[[paste0("Povprečje - ", ime_baza1)]])*100
      
      # T/Z-vrednost in signifikanca
      if(compare_sample_subsample == FALSE) {
        tabela_st[["t"]] <- sapply(statistike, function(x) x[["t"]])
      } else if(compare_sample_subsample == TRUE) {
        tabela_st[["z"]] <- sapply(statistike, function(x) x[["z"]])
      }
      
      tabela_st[["p"]] <- sapply(statistike, function(x) x[["p"]])
      
      tabela_st[["Signifikanca"]] <- weights::starmaker(tabela_st[["p"]])
      
      ## Utežene statistike ------------------------------------------------------
      
      # Utežen Welchev t-test
      # utezene_statistike2 <- lapply(stevilske_spremenljivke, FUN = function(x){
      #    # napačen izračun SE
      #   test <- weights::wtd.t.test(x = baza1_na[[x]], y = baza2_na[[x]],
      #                               weight = utezi1, weighty = utezi2, samedata = FALSE)
      # 
      #   c("povp1" = test$additional[["Mean.x"]],
      #     "povp2" = test$additional[["Mean.y"]],
      #     "t"     = test$coefficients[["t.value"]],
      #     "p"     = test$coefficients[["p.value"]])
      # })
      
      if(compare_sample_subsample == FALSE) {
        utezene_statistike <- lapply(stevilske_spremenljivke, FUN = function(x){
          wtd_t_test(x = baza1_na[[x]], y = baza2_na[[x]],
                     weights_x = utezi1, weights_y = utezi2,
                     se_calculation = se_calculation,
                     survey_design1 = survey_design1, survey_design2 = survey_design2)
        })
      } else if(compare_sample_subsample == TRUE) {
        utezene_statistike <- lapply(stevilske_spremenljivke, FUN = function(x){
          wtd_t_test_sample_subsample(sample = baza1_na[[x]],
                                      subsample = baza2_na[[x]],
                                      weights_sample = utezi1,
                                      weights_subsample = utezi2)
        })
      }
      
      tabela_st_u <- data.frame(row.names = seq_along(utezene_statistike))
      
      # Uteženo povprečje baza 1
      tabela_st_u[[paste0("Povprečje - ", ime_baza1)]] <- sapply(utezene_statistike, function(x) x[["povp1"]])
      
      # Uteženo povprečje baza 2
      tabela_st_u[[paste0("Povprečje - ", ime_baza2)]] <- sapply(utezene_statistike, function(x) x[["povp2"]])
      
      # Absolutna razlika uteženih povprečji
      tabela_st_u[["Razlika v povprečjih - absolutna"]] <- sapply(utezene_statistike, function(x) x[["diff"]])
      
      # Relativna razlika uteženih povprečji
      tabela_st_u[["Razlika v povprečjih - relativna (%)"]] <- (tabela_st_u[["Razlika v povprečjih - absolutna"]]/tabela_st_u[[paste0("Povprečje - ", ime_baza1)]])*100
      
      # Z/T-vrednost in signifikanca
      if(compare_sample_subsample == FALSE) {
        tabela_st_u[["t"]] <- sapply(utezene_statistike, function(x) x[["t"]])
      } else if(compare_sample_subsample == TRUE){
        tabela_st_u[["z"]] <- sapply(utezene_statistike, function(x) x[["z"]])
      }
      
      tabela_st_u[["p"]] <- sapply(utezene_statistike, function(x) x[["p"]])
      
      tabela_st_u[["Signifikanca"]] <- weights::starmaker(tabela_st_u[["p"]])
      
      if(compare_sample_subsample == TRUE) {
        tabela_st_u[["Delež podvzorca v vzorcu - neuteženo"]] <- sapply(statistike, function(x) x[["w"]])
        tabela_st_u[["Delež podvzorca v vzorcu - uteženo"]] <- sapply(utezene_statistike, function(x) x[["w"]])
      }
      
      ## Izvoz excel -------------------------------------------------------------
      
      # povzetek za številske spr.
      merge_tabela_st <- cbind(tabela_st, tabela_st_u)
      frekvence_rel_razlike_neutezene <- count_rel_diff(vec = merge_tabela_st[[10]], p_vec = merge_tabela_st[[12]])
      frekvence_rel_razlike_utezene <- count_rel_diff(vec = merge_tabela_st[[17]], p_vec = merge_tabela_st[[19]])
      
      tbl_st <- data.frame("Relativne razlike" = c("> 20%", "(10% - 20%]", "[5% - 10%]", "< 5%"),
                           # neutežene statistike
                           "f" = frekvence_rel_razlike_neutezene$sums,
                           "%" = frekvence_rel_razlike_neutezene$sums/sum(frekvence_rel_razlike_neutezene$sums),
                           "Kumul f" = frekvence_rel_razlike_neutezene$cumsums,
                           "Kumul %" = frekvence_rel_razlike_neutezene$cumsums/sum(frekvence_rel_razlike_neutezene$sums),
                           "f*" = frekvence_rel_razlike_neutezene$p_sums,
                           "% od vseh spremenljivk" = frekvence_rel_razlike_neutezene$p_sums/sum(frekvence_rel_razlike_neutezene$sums),
                           "% od relativne razlike" = frekvence_rel_razlike_neutezene$p_sums/frekvence_rel_razlike_neutezene$sums,
                           "Kumul f*" = frekvence_rel_razlike_neutezene$p_cumsums,
                           "Kumul %*" = frekvence_rel_razlike_neutezene$p_cumsums/sum(frekvence_rel_razlike_neutezene$sums),
                           # utežene statistike
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
                x = "Številske (intervalne, razmernostne) spremenljivke", startCol = 2, startRow = 1)
      
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
                x = cbind(c(paste("Št. vseh številskih spremenljivk:", sum(frekvence_rel_razlike_neutezene$sums)),
                            paste("Št. statistično značilnih številskih spremenljivk (p < 0.05) - neuteženi podatki:", sum(frekvence_rel_razlike_neutezene$p_sums)),
                            paste("Št. statistično značilnih številskih spremenljivk (p < 0.05) - uteženi podatki:", sum(frekvence_rel_razlike_utezene$p_sums)))),
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
      
      addStyle(wb = wb, sheet = "Opisne statistike",
               style = createStyle(halign = "center"),
               rows = 3:(nrow(tabela_st)+2), cols = 5:ncol(merge_tabela_st),
               gridExpand = TRUE, stack = TRUE)
      
      writeData(wb = wb, sheet = "Opisne statistike",
                x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001", xy = c(1, 1))
      
      addStyle(wb = wb, sheet = "Opisne statistike",
               style = createStyle(fontSize = 9),
               rows = 1, cols = 1)
      
      writeData(wb = wb, sheet = "Opisne statistike",
                x = "Neutežene statistike", xy = c(7, 1))
      
      mergeCells(wb = wb, sheet = "Opisne statistike", cols = 7:13, rows = 1)
      
      writeData(wb = wb, sheet = "Opisne statistike",
                x = "Utežene statistike", xy = c(14, 1))
      
      mergeCells(wb = wb, sheet = "Opisne statistike", cols = 14:20, rows = 1)
      
      addStyle(wb = wb, sheet = "Opisne statistike",
               style = createStyle(textDecoration = "bold",
                                   halign = "center", valign = "center",
                                   fontSize = 12),
               rows = 1, cols = c(7,14),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Opisne statistike",
               style = createStyle(numFmt = "0.00"), rows = 3:(nrow(tabela_st)+2), cols = c(7:12, 14:19, 21:22),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Opisne statistike",
               style = createStyle(fgFill = "#DAEEF3"), rows = 2:(nrow(tabela_st)+2), cols = 7:13,
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Opisne statistike",
               style = createStyle(fgFill = "#FDE9D9"), rows = 2:(nrow(tabela_st)+2), cols = 14:20,
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Opisne statistike",
               style = createStyle(numFmt = "0"),
               rows = 3:(nrow(tabela_st)+2), cols = c(10, 17),
               gridExpand = TRUE, stack = TRUE)
    }
  }
  
  if(!is.null(nominalne_spremenljivke)){
    # Nominalne spremenljivke -------------------------------------------------

    imena_nom_1 <- names(baza1)[names(baza1) %in% nominalne_spremenljivke]
    imena_nom_2 <- names(baza2)[names(baza2) %in% nominalne_spremenljivke]
    
    non_spr <- union(setdiff(imena_nom_1, imena_nom_2), setdiff(imena_nom_2, imena_nom_1))
    
    if(length(non_spr) != 0){
      writeData(wb = wb, sheet = "Opozorila", xy = c(1,5),
                x = paste("Imena podanih opisnih spremenljivk se ne ujemajo v obeh bazah. To so spremenljivke", paste(non_spr, collapse = ", "), "in so bile zato odstranjene iz analiz."))
      
      warning_counter <- TRUE
    }
    
    # izberemo samo spremenljivke, ki so prisotne v 1. in 2. bazi
    nominalne_spremenljivke <- imena_nom_1[imena_nom_1 %in% imena_nom_2]
    
    # funkcija za neutežene in utežene frekvenčne statistike
    tabela_nominalne <- function(spr, baza1, baza2, utezi1, utezi2){
      
      data1 <- as_factor(baza1[[spr]])
      data2 <- as_factor(baza2[[spr]])
      
      if(!all(levels(data1) == levels(data2))){
        return(paste("Kategorije v bazah se pri opisni spremenljivki", spr, "ne ujemajo. Spremenljivka je bila izločena iz analiz."))
        
      } else if(all(is.na(data1)) | all(is.na(data2))) {
        return(paste("Opisna spremenljivka", spr, "ima vse vrednosti manjkajoče (v eni ali obeh bazah). Spremenljivka je bila izločena iz analiz."))
        
      } else if(length(unique(na.omit(data1))) == 1 | length(unique(na.omit(data2))) == 1) {
        return(paste("Opisna spremenljivka", spr, "je konstanta (v eni ali obeh bazah). Spremenljivka je bila izločena iz analiz."))

      } else {
        ## Neutežene statistike ----------------------------------------------------
        labela1 <- attr(baza1[[spr]], "label", exact = TRUE)
        labela2 <- attr(baza2[[spr]], "label", exact = TRUE)
        labela2 <- ifelse(is.null(labela2), "", paste0(" - ", labela2))
        
        tabela_nom <- as.data.frame(table(data1))
        names(tabela_nom) <- c(paste0(spr, ifelse(is.null(labela1), labela2, paste0(" - ", labela1))),
                               paste0("N - ", ime_baza1))
        tabela_nom[,1] <- as.character(tabela_nom[,1])
        
        tabela_nom[[paste0("N - ", ime_baza2)]] <- as.data.frame(table(data2))[,2]
        
        dummies1 <- weights::dummify(data1, keep.na = TRUE)
        dummies2 <- weights::dummify(data2, keep.na = TRUE)
        
        if(compare_sample_subsample == FALSE) {
          # z-test za neodvisna deleža
          n1 <- sum(!is.na(data1))
          n2 <- sum(!is.na(data2))
          
          statistika_nom <- lapply(seq_len(nrow(tabela_nom)), FUN = function(i){
            test <- prop.test(x = c(sum(dummies1[,i], na.rm = TRUE), sum(dummies2[,i], na.rm = TRUE)),
                              n = c(n1, n2), correct = FALSE)
            
            c("povp1" = test$estimate[[1]],
              "povp2" = test$estimate[[2]],
              "z" = unname(sqrt(test$statistic)),
              "p" = test$p.value)
          })
        } else if(compare_sample_subsample == TRUE) {
          statistika_nom <- lapply(seq_len(nrow(tabela_nom)), FUN = function(i) wtd_t_test_sample_subsample(sample = dummies1[,i], subsample = dummies2[,i], weights_subsample = utezi2))
        }
        
        tabela_nom[[paste0("Delež (%) - ", ime_baza1)]] <- sapply(statistika_nom, function(x) x[["povp1"]])*100 
        
        tabela_nom[[paste0("Delež (%) - ", ime_baza2)]] <- sapply(statistika_nom, function(x) x[["povp2"]])*100
        
        tabela_nom[["Razlika v deležih - absolutna"]] <- tabela_nom[[paste0("Delež (%) - ", ime_baza2)]] - tabela_nom[[paste0("Delež (%) - ", ime_baza1)]]
        
        tabela_nom[["Razlika v deležih - relativna (%)"]] <- (tabela_nom[["Razlika v deležih - absolutna"]]/tabela_nom[[paste0("Delež (%) - ", ime_baza1)]])*100
        
        tabela_nom[["z"]] <- sapply(statistika_nom, function(x) x[["z"]])
        
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
        tabela_nom_u[[paste0("N - ", ime_baza2)]] <- weighted_table(data2, utezi2[!is.na(utezi2)])
        
        # Utežen z-test za neodvisna deleža
        if(compare_sample_subsample == FALSE) {
          utezena_statistika_nom <- lapply(seq_len(nrow(tabela_nom_u)), FUN = function(i){
            wtd_t_test(x = dummies1[,i], y = dummies2[,i], weights_x = utezi1, weights_y = utezi2, prop = TRUE,
                       se_calculation = se_calculation, survey_design1 = survey_design1, survey_design2 = survey_design2)
          })
        } else if(compare_sample_subsample == TRUE) {
          utezena_statistika_nom <- lapply(seq_len(nrow(tabela_nom_u)), FUN = function(i){
            wtd_t_test_sample_subsample(sample = dummies1[,i], subsample = dummies2[,i],
                                        weights_sample = utezi1, weights_subsample = utezi2)
          })
        }

        # Utežen delež baza 1
        tabela_nom_u[[paste0("Delež (%) - ", ime_baza1)]] <- sapply(utezena_statistika_nom, function(x) x[["povp1"]]) * 100
        # tabela_nom_u[[paste0("Delež (%) - ", ime_baza1)]] <- (tabela_nom_u[[paste0("N - ", ime_baza1)]]/sum(tabela_nom_u[[paste0("N - ", ime_baza1)]]))*100
        
        # Utežen delež baza 2
        tabela_nom_u[[paste0("Delež (%) - ", ime_baza2)]] <- sapply(utezena_statistika_nom, function(x) x[["povp2"]]) * 100
        # tabela_nom_u[[paste0("Delež (%) - ", ime_baza2)]] <- (tabela_nom_u[[paste0("N - ", ime_baza2)]]/sum(tabela_nom_u[[paste0("N - ", ime_baza2)]]))*100
        
        # Absolutna razlika uteženih deležev
        tabela_nom_u[["Razlika v deležih - absolutna"]] <- sapply(utezena_statistika_nom, function(x) x[["diff"]]) * 100
        # tabela_nom_u[["Razlika v deležih - absolutna"]] <- tabela_nom_u[[paste0("Delež (%) - ", ime_baza2)]] - tabela_nom_u[[paste0("Delež (%) - ", ime_baza1)]]
        
        # Relativna razlika uteženih deležev
        tabela_nom_u[["Razlika v deležih - relativna (%)"]] <- (tabela_nom_u[["Razlika v deležih - absolutna"]]/tabela_nom_u[[paste0("Delež (%) - ", ime_baza1)]])*100
        
        if(compare_sample_subsample == FALSE) {
          tabela_nom_u[["z"]] <- sapply(utezena_statistika_nom, function(x) x[["t"]])
        } else if(compare_sample_subsample == TRUE) {
          tabela_nom_u[["z"]] <- sapply(utezena_statistika_nom, function(x) x[["z"]])
        }
        
        tabela_nom_u[["p"]] <- sapply(utezena_statistika_nom, function(x) x[["p"]])
        
        tabela_nom_u[["Signifikanca"]] <- weights::starmaker(tabela_nom_u[["p"]])
        
        if(compare_sample_subsample == TRUE) {
          tabela_nom_u[["Delež podvzorca v vzorcu - neuteženo"]] <- NA
          tabela_nom_u[["Delež podvzorca v vzorcu - uteženo"]] <- NA
        }
        
        # Skupaj seštevek
        temp_df_u <- data.frame(t(c(colSums(tabela_nom_u[,1:4]), rep(NA, ncol(tabela_nom_u)-4))))
        names(temp_df_u) <- names(tabela_nom_u)
        
        if(compare_sample_subsample == TRUE) {
          temp_df_u[["Delež podvzorca v vzorcu - neuteženo"]] <- statistika_nom[[1]][["w"]]
          temp_df_u[["Delež podvzorca v vzorcu - uteženo"]] <- utezena_statistika_nom[[1]][["w"]]
        }
        
        tabela_nom_u <- rbind(tabela_nom_u, temp_df_u)
        
        return(cbind(tabela_nom, tabela_nom_u))
      }
    }
    
    ## Izvoz excel -------------------------------------------------------------
    
    factor_tables <- lapply(nominalne_spremenljivke, function(x) tabela_nominalne(spr = x, baza1 = baza1, baza2 = baza2, utezi1 = utezi1, utezi2 = utezi2))
    
    # dodamo še opozorila, če so prisotna
    vsota <- cumsum(sapply(factor_tables, is.character))
    
    for(i in seq_along(factor_tables)){
      if(is.character(factor_tables[[i]])){
        writeData(wb = wb, sheet = "Opozorila", startCol = 1, startRow = 5 + vsota[i],
                  x = factor_tables[[i]])
        
        warning_counter <- TRUE
      }
    }
    
    factor_tables <- factor_tables[!vapply(factor_tables, is.character, FUN.VALUE = logical(1))]
    
    if(length(factor_tables) > 0){
      
      # skupno število kategorij
      st_kategorij <- sum(sapply(factor_tables, function(x) sum(x[[1]] != "Skupaj")))
      
      # INF ali NaN relativne razlike popravimo
      factor_tables <- lapply(factor_tables, function(x) {
        x[is.finite(x[,7]) == FALSE,7] <- x[is.finite(x[,7]) == FALSE,6]
        x[is.finite(x[,16]) == FALSE,16] <- x[is.finite(x[,16]) == FALSE,15]
        x[is.na(x)] <- NA
        x
      })
      
      # začasno damo NaN vrednosti p kot 1 (za štetje stat značilnih)
      temp_factor_tables <- lapply(factor_tables, function(x) {
        x[is.finite(x[,9]) == FALSE,9] <- 1
        x[is.finite(x[,18]) == FALSE,18] <- 1
        x
      })
      
      # označi se vrstice, kjer je premajhen numerus za zanesljivo oceno p-vrednosti
      opozorilo_numerus <- lapply(temp_factor_tables, FUN = function(x) {
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
      neutezene_kategorije <- abs(na.omit(do.call(rbind, lapply(temp_factor_tables, "[", c(7, 9)))))
      # nadomestimo p vrednosti s premajhnim numerusom z 1 (da se jih ne šteje pod stat. značilne)
      if(any(p_numerus_neutezene_kategorije)) neutezene_kategorije[p_numerus_neutezene_kategorije,][[2]] <- 1
      
      utezene_kategorije <- abs(na.omit(do.call(rbind, lapply(temp_factor_tables, "[", c(16, 18)))))
      if(any(p_numerus_utezene_kategorije)) utezene_kategorije[p_numerus_utezene_kategorije,][[2]] <- 1
      
      frekvence_rel_razlike_neutezene_nom <- count_rel_diff(vec = neutezene_kategorije[[1]], p_vec = neutezene_kategorije[[2]])
      frekvence_rel_razlike_utezene_nom <- count_rel_diff(vec = utezene_kategorije[[1]], p_vec = utezene_kategorije[[2]])
      
      tbl_nom <- data.frame("Relativne razlike" = c("> 20%", "(10% - 20%]", "[5% - 10%]", "< 5%"),
                            # neutežene statistike
                            "f" = frekvence_rel_razlike_neutezene_nom$sums,
                            "%" = frekvence_rel_razlike_neutezene_nom$sums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                            "Kumul f" = frekvence_rel_razlike_neutezene_nom$cumsums,
                            "Kumul %" = frekvence_rel_razlike_neutezene_nom$cumsums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                            "f*" = frekvence_rel_razlike_neutezene_nom$p_sums,
                            "% od vseh kategorij" = frekvence_rel_razlike_neutezene_nom$p_sums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                            "% od relativne razlike" = frekvence_rel_razlike_neutezene_nom$p_sums/frekvence_rel_razlike_neutezene_nom$sums,
                            "Kumul f*" = frekvence_rel_razlike_neutezene_nom$p_cumsums,
                            "Kumul %*" = frekvence_rel_razlike_neutezene_nom$p_cumsums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                            # utežene statistike
                            "f" = frekvence_rel_razlike_utezene_nom$sums,
                            "%" = frekvence_rel_razlike_utezene_nom$sums/sum(frekvence_rel_razlike_utezene_nom$sums),
                            "Kumul f" = frekvence_rel_razlike_utezene_nom$cumsums,
                            "Kumul %" = frekvence_rel_razlike_utezene_nom$cumsums/sum(frekvence_rel_razlike_utezene_nom$sums),
                            "f*" = frekvence_rel_razlike_utezene_nom$p_sums,
                            "% od vseh kategorij" = frekvence_rel_razlike_utezene_nom$p_sums/sum(frekvence_rel_razlike_utezene_nom$sums),
                            "% od relativne razlike" = frekvence_rel_razlike_utezene_nom$p_sums/frekvence_rel_razlike_utezene_nom$sums,
                            "Kumul f*" = frekvence_rel_razlike_utezene_nom$p_cumsums,
                            "Kumul %*" = frekvence_rel_razlike_utezene_nom$p_cumsums/sum(frekvence_rel_razlike_utezene_nom$sums),
                            check.names = FALSE)
      
      tbl_nom[is.na(tbl_nom)] <- 0
      
      writeData(wb = wb,
                sheet = "Povzetek",
                x = tbl_nom,
                borders = "all", startRow = 18,
                headerStyle = createStyle(textDecoration = "bold",
                                          wrapText = TRUE,
                                          border = c("top", "bottom", "left", "right"),
                                          halign = "center", valign = "center"))
      
      writeData(wb = wb, sheet = "Povzetek",
                x = "Opisne (nominalne, ordinalne) spremenljivke", startCol = 2, startRow = 15)
      
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
                x = cbind(c(paste("Št. vseh kategorij:", st_kategorij),
                            paste("Št. statistično značilnih kategorij (p < 0.05) - neuteženi podatki:", sum(frekvence_rel_razlike_neutezene_nom$p_sums)),
                            paste("Št. statistično značilnih kategorij (p < 0.05) - uteženi podatki:", sum(frekvence_rel_razlike_utezene_nom$p_sums)))),
                startCol = 1, startRow = 24, rowNames = FALSE, colNames = FALSE)
      
      writeData(wb = wb, sheet = "Povzetek",
                x = cbind(c(paste("Št. vseh opisnih spremenljivk:", length(factor_tables)),
                            paste("Št. statistično značilnih opisnih spremenljivk (p < 0.05) - neuteženi podatki:",
                                  sum(vapply(seq_along(temp_factor_tables), function(i){
                                    any(temp_factor_tables[[i]][[9]][!opozorilo_numerus[[i]][["p_neutez"]]] < 0.05)
                                  }, FUN.VALUE = logical(1)))),
                            paste("Št. statistično značilnih opisnih spremenljivk (p < 0.05) - uteženi podatki:", 
                                  sum(vapply(seq_along(temp_factor_tables), function(i){
                                    any(temp_factor_tables[[i]][[18]][!opozorilo_numerus[[i]][["p_utez"]]] < 0.05)
                                  }, FUN.VALUE = logical(1)))))),
                startCol = 1, startRow = 28, rowNames = FALSE, colNames = FALSE)
      
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
               style = createStyle(numFmt = "0%"),
               rows = 19:22, cols = c(3, 5, 7, 8, 10, 12, 14, 16, 17, 19),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Povzetek",
               style = createStyle(halign = "center"),
               rows = 19:22, cols = 1:19,
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Povzetek",
               style = createStyle(fgFill = "#DAEEF3"), rows = 17:22, cols = 2:10,
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Povzetek",
               style = createStyle(fgFill = "#FDE9D9"), rows = 17:22, cols = 11:19,
               gridExpand = TRUE, stack = TRUE)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 2:10, rows = 23)
      mergeCells(wb = wb, sheet = "Povzetek", cols = 11:19, rows = 23)
      
      addStyle(wb = wb, sheet = "Povzetek",
               style = createStyle(wrapText = TRUE),
               rows = 23, cols = c(2, 11),
               gridExpand = TRUE, stack = TRUE)
      
      setColWidths(wb = wb, sheet = "Povzetek", cols = 1:19,
                   widths = c(16, rep(c(5, 6, 8, 8, 5, 21, 21, 8, 9), 2)))
      
      setRowHeights(wb = wb, sheet = "Povzetek", rows = c(15:18, 23), heights = c(20, 20, 25, 44, 28))
      
      # opozorilo če n <= 5
      if(any(p_numerus_neutezene_kategorije)) {
        writeData(wb = wb, sheet = "Povzetek",
                  x = paste0("Opomba: Nekatere izmed kategorij (št. takih kategorij: ", sum(p_numerus_neutezene_kategorije),") so bile statistično značilne, vendar zaradi premajhnega št. enot niso bile upoštevane v povzetku."),
                  startCol = 2, startRow = 23)
      }
      
      if(any(p_numerus_utezene_kategorije)) {
        writeData(wb = wb, sheet = "Povzetek",
                  x = paste0("Opomba: Nekatere izmed kategorij (št. takih kategorij: ", sum(p_numerus_utezene_kategorije),") so bile statistično značilne, vendar zaradi premajhnega št. enot niso bile upoštevane v povzetku."),
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
      start_rows <- c(0, cumsum(2 + sapply(factor_tables, nrow)[-length(factor_tables)])) + 3
      
      for(i in seq_along(factor_tables)){
        
        writeData(wb = wb, sheet = "Frekvencne tabele",
                  x = factor_tables[[i]],
                  borders = "all",
                  startRow = start_rows[i],
                  headerStyle = createStyle(textDecoration = "bold",
                                            wrapText = TRUE,
                                            border = c("top", "bottom", "left", "right"),
                                            halign = "center", valign = "center"))
        
        addStyle(wb = wb, sheet = "Frekvencne tabele",
                 style = createStyle(fgFill = "#DAEEF3"), rows = start_rows[i]:(start_rows[i]+nrow(factor_tables[[i]])), cols = 2:10,
                 gridExpand = TRUE, stack = TRUE)
        
        addStyle(wb = wb, sheet = "Frekvencne tabele",
                 style = createStyle(fgFill = "#FDE9D9"), rows = start_rows[i]:(start_rows[i]+nrow(factor_tables[[i]])), cols = 11:19,
                 gridExpand = TRUE, stack = TRUE)
        
        addStyle(wb = wb, sheet = "Frekvencne tabele",
                 style = createStyle(fgFill = "#D9D9D9"), rows = start_rows[i] + which(opozorilo_numerus[[i]]$neutez == TRUE), cols = 2:10,
                 gridExpand = TRUE, stack = TRUE)
        
        addStyle(wb = wb, sheet = "Frekvencne tabele",
                 style = createStyle(fgFill = "#D9D9D9"), rows = start_rows[i] + which(opozorilo_numerus[[i]]$utez == TRUE), cols = 11:19,
                 gridExpand = TRUE, stack = TRUE)
      }
      
      if(any(unlist(opozorilo_numerus))) {
        writeData(wb = wb, sheet = "Frekvencne tabele",
                  x = "Opozorilo: Število enot v celici je premajhno za zanesljivo oceno p-vrednosti", startCol = 1, startRow = 1)
        
        addStyle(wb = wb, sheet = "Frekvencne tabele",
                 style = createStyle(fgFill = "#D9D9D9", wrapText = TRUE), rows = 1, cols = 1,
                 gridExpand = TRUE, stack = TRUE)
      }
      
      writeData(wb = wb, sheet = "Frekvencne tabele",
                x = "Neutežene statistike", startCol = 2, startRow = 2)
      
      mergeCells(wb = wb, sheet = "Frekvencne tabele", cols = 2:10, rows = 2)
      
      writeData(wb = wb, sheet = "Frekvencne tabele",
                x = "Utežene statistike", startCol = 11, startRow = 2)
      
      mergeCells(wb = wb, sheet = "Frekvencne tabele", cols = 11:19, rows = 2)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(textDecoration = "bold",
                                   halign = "center", valign = "center",
                                   fontSize = 12),
               rows = 2, cols = c(2, 11),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(numFmt = "0.00"), rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])), cols = c(8,9,17,18,20:21),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(numFmt = "0"), rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])), cols = c(4:7, 11:16),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = "Frekvencne tabele",
               style = createStyle(halign = "center"), rows = 1:(start_rows[length(start_rows)]+nrow(factor_tables[[length(factor_tables)]])), cols = 2:21,
               gridExpand = TRUE, stack = TRUE)
      
      setColWidths(wb = wb, sheet = "Frekvencne tabele", cols = 1, widths = 60)
    }
  }
  
  if(warning_counter == FALSE){
    writeData(wb = wb, sheet = "Opozorila", startCol = 1, startRow = 1,
              x = "Ni opozoril")
  }

  # shranimo excel datoteko
  saveWorkbook(wb = wb, file = file, overwrite = TRUE)
}