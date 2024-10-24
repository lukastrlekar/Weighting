# TODO preveri, da se pri neuteženih ne uporablja survey funkcij, ker avtomatsko upošteva uteži

# any(class(survey_design) %in% "survey.design") - dodaj survey design in možnost izračuna SE 

# TODO Q164 Regija_vzorec Welcj ANOVA NaN in results may be dubious pr primerjavi ind. povprečij - določene skupine imajo ničelno var
# potrebno poseben simbol pri teh
#Robust tests of equality of means cannot be performed for Q164 because at least one group has 0 variance.	
# ALI Skupno povprečje veljavnih skupin (N > 1 & var > 0)
# in se naredi . pri vseh skupinah kjer var = 0



# svyglm že privzeto izračuna heteroscedasticity robust standard errors (okvirno HC0)
# https://stats.stackexchange.com/questions/57107/use-of-weights-in-svyglm-vs-glm
# https://stats.stackexchange.com/questions/333930/equivalence-of-svyglm-and-glm-for-simple-random-surveys

# razbitje
izvoz_excel_razbitje <- function(baza1 = NULL,
                                 utezi_spr = NULL,
                                 stevilske_spremenljivke = NULL,
                                 nominalne_spremenljivke = NULL,
                                 razbitje_spremenljivke = NULL,
                                 file,
                                 color_sig){
  
  wb <- createWorkbook()
  addWorksheet(wb = wb, sheetName = "Opozorila", gridLines = FALSE)
  addWorksheet(wb = wb, sheetName = "Povzetek", gridLines = FALSE)
  
  warning_counter <- FALSE
  
  # preverimo, da so spremenljivke res številske in neštevilske odstranimo
  numeric_variables <- vapply(stevilske_spremenljivke, function(x){
    is.numeric(baza1[[x]]) | is.integer(baza1[[x]])
  }, FUN.VALUE = logical(1), USE.NAMES = FALSE)
  
  if(!all(numeric_variables)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,1),
              x = paste("Spremenljivke", paste(stevilske_spremenljivke[!numeric_variables], collapse = ", "), "so bile odstranjene iz analiz, saj vsebujejo neštevilske vrednosti oziroma so bile v SPSS definirane kot neštevilske."))
    
    warning_counter <- TRUE
    
    stevilske_spremenljivke <- stevilske_spremenljivke[numeric_variables]
  }
  
  # preverimo in odstranimo spremenljivke z ničelno varianco (konstante) (tu se odstranijo tudi spremenljivke z vsemi NA)
  indeksi_variances <- vapply(stevilske_spremenljivke, function(x){
    v1 <- var(baza1[[x]], na.rm = TRUE)
    v1 <- ifelse(is.na(v1), 0, v1)
    v1 == 0
  }, FUN.VALUE = logical(1), USE.NAMES = FALSE)
  
  if(any(indeksi_variances)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,2),
              x = paste("Številske spremenljivke", paste(stevilske_spremenljivke[indeksi_variances], collapse = ", "), "so bile odstranjene, saj so konstante (imajo ničelno varianco)."))
    
    warning_counter <- TRUE
    
    stevilske_spremenljivke <- stevilske_spremenljivke[!indeksi_variances]
  }
  
  # preverimo in odstranimo spremenljivke razbitja, ki so konstante ali imajo le NA vrednosti
  indeksi_razbitje <- vapply(razbitje_spremenljivke, function(x){
    length(unique(baza1[[x]][!is.na(baza1[[x]])])) > 1
  }, FUN.VALUE = logical(1), USE.NAMES = FALSE)
  
  if(any(!indeksi_razbitje)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,3),
              x = paste("Spremenljivke razbitja", paste(razbitje_spremenljivke[!indeksi_razbitje], collapse = ", "),
                        "se pri analizi niso upoštevale, saj so konstante."))
    
    warning_counter <- TRUE
    
    razbitje_spremenljivke <- razbitje_spremenljivke[indeksi_razbitje]
  }
  
  # preverimo in odstranimo spremenljivke razbitja, ki se pojavijo med ostalimi spr.
  spr <- c(stevilske_spremenljivke, nominalne_spremenljivke)
  indeksi_spr <- razbitje_spremenljivke %in% spr
  
  if(any(indeksi_spr)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,4),
              x = paste("Spremenljivke razbitja", paste(razbitje_spremenljivke[indeksi_spr], collapse = ", "),
                        "se pri analizi niso upoštevale, saj so bile izbrane že med ostalimi spremenljivkami."))
    
    warning_counter <- TRUE
    
    razbitje_spremenljivke <- razbitje_spremenljivke[!indeksi_spr]
  }
  
  ### ŠTEVILSKE SPREMENLJIVKE
  if(length(stevilske_spremenljivke) >= 1 && length(razbitje_spremenljivke) >= 1){
    
    razbitje <- function(stevilske_spremenljivke,
                         razbitje_spremenljivke,
                         baza1,
                         use_weights,
                         utezi_spr = NULL){
      
      stolpci <- seq(from = 2, length.out = length(stevilske_spremenljivke), by = 2)
      list_razbitje <- vector("list", length(razbitje_spremenljivke))
      names(list_razbitje) <- razbitje_spremenljivke
      
      p_vrednosti <- vector("list", length(razbitje_spremenljivke))
      names(p_vrednosti) <- razbitje_spremenljivke
      
      f_vrednosti <- vector("list", length(razbitje_spremenljivke))
      names(f_vrednosti) <- razbitje_spremenljivke
      
      anova_p_vrednosti <- vector("list", length(razbitje_spremenljivke))
      names(anova_p_vrednosti) <- razbitje_spremenljivke
      
      for(name in razbitje_spremenljivke) {
        temp_baza1 <- baza1[, c(stevilske_spremenljivke, name, utezi_spr), drop = FALSE]
        temp_baza1[[name]] <- as_factor(temp_baza1[[name]])
        
        kat <- levels(temp_baza1[[name]])
        
        temp_df <- data.frame(matrix(NA, nrow = length(kat) + 1, ncol = (length(stevilske_spremenljivke) * 2) + 1))
        
        temp_df[[1]] <- c(kat, "Skupno povprečje veljavnih skupin (N > 1)")
        
        # statistična značilnost
        temp_p <- matrix(1, nrow = nrow(temp_df) - 1, ncol = length(stolpci))
        temp_f_vrednosti <- vector("list", length = length(stevilske_spremenljivke))
        temp_anova_p_vrednosti <- vector("numeric", length = length(stevilske_spremenljivke))
        
        for(i in seq_along(stevilske_spremenljivke)){
          temp_baza <- temp_baza1[, c(stevilske_spremenljivke[i], name, utezi_spr), drop = FALSE]
          
          # velikosti skupin
          n_sk <- tapply(temp_baza[[stevilske_spremenljivke[i]]], temp_baza[[name]], function(x) sum(!is.na(x)), simplify = TRUE)
          temp_df[-nrow(temp_df), stolpci[i] + 1] <- as.vector(n_sk)
          
          rows <- names(n_sk)[which(n_sk > 1)]
          rows <- match(rows, kat)
          
          # obdržimo samo skupine, kjer sta vsaj 2 enoti
          temp_baza <- temp_baza[!temp_baza[[name]] %in% names(n_sk)[which(n_sk <= 1)], ]
          temp_baza[[name]] <- droplevels(temp_baza[[name]])
          temp_baza <- as.data.frame(temp_baza)
          
          if(use_weights){ # uteženi podatki
            
            # povprečja skupin
            temp_df[rows, stolpci[i]] <- vapply(levels(temp_baza[[name]]), 
                                                function(whichpart) {
                                                  weighted.mean(x = temp_baza[[stevilske_spremenljivke[i]]][temp_baza[[name]] == whichpart], 
                                                                w = temp_baza[[utezi_spr]][temp_baza[[name]] == whichpart], na.rm = TRUE)
                                                }, FUN.VALUE = numeric(1), USE.NAMES = FALSE)
            # skupno povprečje
            temp_df[nrow(temp_df), stolpci[i]] <- weighted.mean(x = temp_baza[!is.na(temp_baza[[name]]), stevilske_spremenljivke[i], drop = TRUE],
                                                                w = temp_baza[!is.na(temp_baza[[name]]), utezi_spr, drop = TRUE],
                                                                na.rm = TRUE)
            
            if(length(rows) > 1){
              # primerjave povprečij s skupnim povprečjem
              # https://stackoverflow.com/questions/72843411/one-way-anova-using-the-survey-package-in-r
              survey_design <- svydesign(id = ~1, weights = as.formula(paste("~", utezi_spr)), data = temp_baza)
              lm_model <- svyglm(as.formula(paste(stevilske_spremenljivke[i], "~", name)), design = survey_design)
              
              # glht_model_zac <- glht(lm_model, linfct = do.call(mcp, setNames(list("GrandMean"), name)))
              # ta način ni OK, ker je matrika kontrastov na neuteženih N in ne pride pravilno
              # zato treba ročno pripraviti matriko kontrastov, kjer se upošteva uteži
              n_contr <- vapply(levels(temp_baza[[name]]), 
                                function(whichpart) {
                                  tmp_df <- temp_baza[!is.na(temp_baza[[stevilske_spremenljivke[i]]]), ]
                                  sum(tmp_df[[utezi_spr]][tmp_df[[name]] == whichpart], na.rm = TRUE)
                                }, FUN.VALUE = numeric(1), USE.NAMES = TRUE)
              
              contr_mat <- contrMat(n_contr, type = "GrandMean")
              contr_mat[,1] <- 0
              glht_model_zac <- glht(lm_model, linfct = contr_mat)
              
              glht_model <- summary(glht_model_zac)
              
              f_test <- summary(glht_model_zac, test = Ftest()) # v tem primeru (GrandMean contrasts) je to enako rezultati pri regTermTest
              
              temp_p[rows, i] <- as.vector(glht_model$test$pvalues)
              
              temp_f_vrednosti[[i]] <- paste0("F(", f_test$test$df[1L],
                                              ", ", f_test$test$df[2L], ") = ",
                                              round(f_test$test$fstat[1L], 2),
                                              weights::starmaker(f_test$test$pvalue[1L]))
              
              temp_anova_p_vrednosti[[i]] <- f_test$test$pvalue[1L]
            } else {
              temp_f_vrednosti[[i]] <- NA
              temp_anova_p_vrednosti[[i]] <- NA
            }

          } else { # neuteženi podatki
            
            # povprečja skupin
            temp_df[rows, stolpci[i]] <- as.vector(tapply(temp_baza[[stevilske_spremenljivke[i]]], temp_baza[[name]], mean, na.rm = TRUE, simplify = TRUE))
            # skupno povprečje
            temp_df[nrow(temp_df), stolpci[i]] <- mean(temp_baza[!is.na(temp_baza[[name]]), stevilske_spremenljivke[i], drop = TRUE], na.rm = TRUE)
            
            if(length(rows) > 1){
              # oneway ANOVA
              f_test <- oneway.test(as.formula(paste0(stevilske_spremenljivke[i], "~", name)), data = temp_baza, var.equal = FALSE)
              
              # primerjave povprečij s skupnim povprečjem
              # uporabimo pristop 'computation of group-specific variance estimates - approach is more robust in the presence of small sample sizes compared to sandwich variance estimation'
              model <- SimComp::SimTestDiff(data = temp_baza[!is.na(temp_baza[[stevilske_spremenljivke[i]]]),],
                                            grp = name,
                                            resp = stevilske_spremenljivke[i],
                                            type = "GrandMean",
                                            covar.equal = FALSE)
              
              temp_p[rows, i] <- as.vector(model$p.val.adj)
              
              temp_f_vrednosti[[i]] <- ifelse(is.na(f_test$p.value),
                                              NA,
                                              paste0("F(", round(f_test$parameter[1L]),
                                                     ", ", round(f_test$parameter[2L]), ") = ",
                                                     round(f_test$statistic[1L], 2),
                                                     weights::starmaker(f_test$p.value)))
              
              temp_anova_p_vrednosti[[i]] <- f_test$p.value
            } else {
              temp_f_vrednosti[[i]] <- NA
              temp_anova_p_vrednosti[[i]] <- NA
            }
            
            # lm_model <- lm(as.formula(paste0(stevilske_spremenljivke[i], "~", name)), data = temp_baza)
            # glht_model <- summary(glht(lm_model, linfct = do.call(mcp, setNames(list("GrandMean"), name))))
          }
          
          temp_df[nrow(temp_df), stolpci[i] + 1] <- sum(!is.na(temp_baza[!is.na(temp_baza[[name]]), stevilske_spremenljivke[i], drop = TRUE]))
        }
        
        p_vrednosti[[name]] <- temp_p
        f_vrednosti[[name]] <- unlist(temp_f_vrednosti, use.names = FALSE)
        anova_p_vrednosti[[name]] <- temp_anova_p_vrednosti
        list_razbitje[[name]] <- temp_df
      }
      
      # relativne razlike
      rel_razlike <- lapply(list_razbitje, function(x){
        df <- x[, stolpci, drop = FALSE]
        
        df[] <- lapply(df, function(s){
          ((s - s[length(s)])/s[length(s)]) * 100
        })
        
        df[nrow(df) + 1, ] <- 0
        df
      })
      
      list(stolpci = stolpci,
           list_razbitje = list_razbitje,
           p_vrednosti = p_vrednosti,
           f_vrednosti = f_vrednosti,
           anova_p_vrednosti = anova_p_vrednosti,
           rel_razlike = rel_razlike)
    }
    
    razbitje_write_excel <- function(sheet_name,
                                     list_razbitje,
                                     p_vrednosti,
                                     f_vrednosti,
                                     anova_p_vrednosti,
                                     rel_razlike,
                                     wb,
                                     utezi, 
                                     use_weights,
                                     stolpci,
                                     color_sig){
      
      addWorksheet(wb = wb, sheetName = sheet_name, gridLines = FALSE)
      
      list_razbitje_excel <- list_razbitje
      
      for(i in seq_along(list_razbitje_excel)){
        list_razbitje_excel[[i]][, stolpci] <- round(list_razbitje_excel[[i]][, stolpci], 2)
        list_razbitje_excel[[i]][, stolpci][is.na(list_razbitje_excel[[i]][, stolpci])] <- "•"
        
        for(s in seq_len(ncol(p_vrednosti[[i]]))){
          list_razbitje_excel[[i]][-nrow(list_razbitje_excel[[i]]), stolpci[s]] <-
            paste(list_razbitje_excel[[i]][-nrow(list_razbitje_excel[[i]]), stolpci[s]], weights::starmaker(p_vrednosti[[i]][,s]))
        }
        
        list_razbitje_excel[[i]][nrow(list_razbitje_excel[[i]]) + 1, stolpci] <- f_vrednosti[[i]]
        list_razbitje_excel[[i]][nrow(list_razbitje_excel[[i]]), 1] <- paste("Test ANOVA", ifelse(use_weights, "(Wald)", "(Welch)"))
      }
      
      vsota <- c(0, cumsum(sapply(list_razbitje_excel, nrow) + 1))
      
      for(i in seq_along(list_razbitje)){
        writeData(wb = wb, sheet = sheet_name, startCol = 1, startRow = 17 + vsota[i],
                  x = list_razbitje_excel[[i]], colNames = FALSE, rowNames = FALSE,
                  borders = "surrounding")
        
        writeData(wb = wb, sheet = sheet_name,
                  x = paste(names(list_razbitje)[i], 
                            ifelse(is.null(var_label(baza1[[names(list_razbitje[i])]])),
                                   "",
                                   paste("-", var_label(baza1[[names(list_razbitje[i])]])))),
                  startCol = 1, startRow = 16 + vsota[i],
                  borders = "surrounding")
        
        mergeCells(wb = wb, sheet = sheet_name,
                   cols = 1:ncol(list_razbitje[[1]]), rows = 16 + vsota[i])
        
        addStyle(wb = wb, sheet = sheet_name,
                 style = createStyle(textDecoration = "bold", valign = "center", fgFill = "#D0CECE"),
                 cols = 1:ncol(list_razbitje[[1]]), rows = 16 + vsota[i],
                 gridExpand = TRUE, stack = TRUE)
        
        addStyle(wb = wb, sheet = sheet_name,
                 style = createStyle(textDecoration = "italic"),
                 cols = 1:ncol(list_razbitje[[1]]),
                 rows = (15 + vsota[i] + nrow(list_razbitje_excel[[i]])) : ((16 + vsota[i] + nrow(list_razbitje_excel[[i]]))),
                 gridExpand = TRUE, stack = TRUE)
      }
      
      writeData(wb = wb, sheet = sheet_name,
                x = t(as.vector(rbind(rep("Povprečje", length(stevilske_spremenljivke)), rep("N", length(stevilske_spremenljivke))))),
                startCol = 2, startRow = 14, colNames = FALSE, borders = "surrounding")
      
      writeData(wb = wb, sheet = sheet_name,
                x = t(as.vector(rbind(apply(baza1[, stevilske_spremenljivke, drop = FALSE], 2, weighted.mean, w = utezi, na.rm = TRUE),
                                      colSums(!is.na(baza1[, stevilske_spremenljivke, drop = FALSE]))))),
                startCol = 2, startRow = 15, colNames = FALSE, borders = "surrounding")
      
      writeData(wb = wb, sheet = sheet_name,
                x = "Skupno povprečje na celotnem vzorcu",
                startCol = 1, startRow = 15, colNames = FALSE, borders = "surrounding")
      
      
      for(i in seq_along(stevilske_spremenljivke)){
        writeData(wb = wb, sheet = sheet_name,
                  x = paste(stevilske_spremenljivke[i], 
                            ifelse(is.null(var_label(baza1[[stevilske_spremenljivke[i]]])),
                                   "",
                                   paste("-", var_label(baza1[[stevilske_spremenljivke[i]]])))),
                  startCol = stolpci[i], startRow = 12, borders = "surrounding")
        
        labele <- attr(x = baza1[[stevilske_spremenljivke[i]]], which = "labels", exact = TRUE)
        if(!is.null(labele)){
          string_min <- stringr::str_squish(paste(unname(labele[1]), "-", names(labele[1])))
          string_min <- paste(stringr::str_unique(stringr::str_split_1(string_min, " ")), collapse = " ")
          
          string_max <- stringr::str_squish(paste(unname(labele[length(labele)]), "-", names(labele[length(labele)])))
          string_max <- paste(stringr::str_unique(stringr::str_split_1(string_max, " ")), collapse = " ")
          
          string <- paste0(string_min, " ; ", string_max)
        } else {
          string <- paste0(min(baza1[[stevilske_spremenljivke[i]]], na.rm = TRUE), " - ",
                           max(baza1[[stevilske_spremenljivke[i]]], na.rm = TRUE))
        }
        
        writeData(wb = wb, sheet = sheet_name,
                  x = string,
                  startCol = stolpci[i], startRow = 13, borders = "surrounding")
        
        mergeCells(wb = wb, sheet = sheet_name,
                   cols = stolpci[i]:(stolpci[i] + 1), rows = 12)
        
        mergeCells(wb = wb, sheet = sheet_name,
                   cols = stolpci[i]:(stolpci[i] + 1), rows = 13)
        
        for(j in seq_along(list_razbitje)){
          mergeCells(wb = wb, sheet = sheet_name,
                     cols = stolpci[i]:(stolpci[i] + 1), rows = 16 + vsota[j] + nrow(list_razbitje_excel[[j]]))
        }
        
        addStyle(wb = wb, sheet = sheet_name,
                 style = createStyle(border = "right"),
                 rows = 12:(sum(sapply(list_razbitje_excel, nrow) + 1) + 15),
                 cols = c(1, stolpci[i] + 1),
                 gridExpand = TRUE, stack = TRUE)
      }
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(border = c("top")),
               rows = 12,
               cols = 1:ncol(list_razbitje[[1]]),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(border = c("bottom")),
               rows = 12,
               cols = 2:ncol(list_razbitje[[1]]),
               gridExpand = TRUE, stack = TRUE)

      # obarvanje rel. razlik le, če p < 0.1
      if(color_sig){
        p_vrednosti_col <- lapply(p_vrednosti, function(m){
          rbind(m, matrix(data = 1, ncol = ncol(m), nrow = 2))
        })
        
        for(s in seq_along(razbitje_spremenljivke)) {
          for(i in seq_len(nrow(rel_razlike[[s]]))) {
            for(j in seq_len(ncol(rel_razlike[[s]]))) {
              if(isTRUE(rel_razlike[[s]][i,j] > 20 & p_vrednosti_col[[s]][i,j] < 0.1)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#EC8984"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] > 10 & p_vrednosti_col[[s]][i,j] < 0.1)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#F5C0BD"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] > 5 & p_vrednosti_col[[s]][i,j] < 0.1)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#FBE9E9"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] < -20 & p_vrednosti_col[[s]][i,j] < 0.1)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#538DD5"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] < -10 & p_vrednosti_col[[s]][i,j] < 0.1)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#8DB4E2"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] < -5 & p_vrednosti_col[[s]][i,j] < 0.1)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#C5D9F1"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              }
            }
          }
        }
      } else { # obarvanje vseh rel. razlik
        for(s in seq_along(razbitje_spremenljivke)) {
          for(i in seq_len(nrow(rel_razlike[[s]]))) {
            for(j in seq_len(ncol(rel_razlike[[s]]))) {
              if(isTRUE(rel_razlike[[s]][i,j] > 20)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#EC8984"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] > 10)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#F5C0BD"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] > 5)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#FBE9E9"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] < -20)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#538DD5"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] < -10)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#8DB4E2"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              } else if(isTRUE(rel_razlike[[s]][i,j] < -5)){
                addStyle(wb = wb,
                         sheet = sheet_name,
                         style = createStyle(fgFill = "#C5D9F1"),
                         rows = 16 + vsota[s] + i, cols = stolpci[j],
                         gridExpand = TRUE, stack = TRUE)
              }
            }
          }
        }
      }

      writeData(wb = wb,
                sheet = sheet_name,
                x = paste("Primerjava", ifelse(use_weights, "uteženih", "neuteženih"), "povprečij po skupinah (primerjava povprečja posamezne skupine s skupnim povprečjem skupin)"),
                startCol = 1, startRow = 1)
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = paste0("Relativna razlika glede na skupno povprečje",
                           ifelse(color_sig, " & p < 0.1", ""),
                           ":"),
                startCol = 1, startRow = 2)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Visoka pozitivna (> 20 %)",
                startCol = 1, startRow = 3)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Zmerna pozitivna (10 % – 20 %)",
                startCol = 1, startRow = 4)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Šibka pozitivna (5 % – 10 %)",
                startCol = 1, startRow = 5)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Šibka negativna (-5 % – -10 %) ",
                startCol = 1, startRow = 6)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Zmerna negativna (-10 % – -20 %)",
                startCol = 1, startRow = 7)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Visoka negativna (< -20 %)",
                startCol = 1, startRow = 8)
      
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
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#C5D9F1"),
               rows = 6, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#8DB4E2"),
               rows = 7, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#538DD5"),
               rows = 8, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Signifikanca: + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001",
                startCol = 1, startRow = 9)
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = paste0("Test ANOVA ", ifelse(use_weights, "(Wald)", "(Welch)"), " s predpostavko neenakih varianc po skupinah: poročana je F-statistika, testira se ničelna domneva, da so vsa povprečja po skupinah enaka"),
                startCol = 1, startRow = 10)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(textDecoration = "bold",
                                   fontSize = 14),
               rows = 1, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(textDecoration = "bold", valign = "top", halign = "center", wrapText = TRUE, fgFill = "#E7E6E6"),
               rows = 12, cols = 1:ncol(list_razbitje[[1]]),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(valign = "center", halign = "center", wrapText = TRUE, fgFill = "#E7E6E6"),
               rows = 13, cols = 1:ncol(list_razbitje[[1]]),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(textDecoration = c("bold", "italic"), halign = "center", valign = "center", fgFill = "#E7E6E6"),
               rows = 14, cols = 1:ncol(list_razbitje[[1]]),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(halign = "center"),
               rows = 15:(sum(sapply(list_razbitje_excel, nrow) + 1) + 15), cols = 2:ncol(list_razbitje[[1]]),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(textDecoration = "bold"),
               rows = (sapply(list_razbitje_excel, nrow)) + vsota[-length(vsota)] + 15,
               cols = 1:ncol(list_razbitje[[1]]),
               gridExpand = TRUE, stack = TRUE)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(numFmt = "0.00"),
               rows = 15, cols = stolpci,
               gridExpand = TRUE, stack = TRUE)
      
      setColWidths(wb = wb,
                   sheet = sheet_name,
                   cols = 1,
                   widths = 47)
      
      setRowHeights(wb = wb,
                    sheet = sheet_name,
                    rows = c(12, 13, 14, 16 + vsota[-length(vsota)]),
                    heights = c(45, 30, 18, rep(18, length(vsota) - 1)))
      
      freezePane(wb = wb, sheet = sheet_name, firstCol = TRUE)
    }
    
    razbitje_list_neutezeno <- razbitje(stevilske_spremenljivke = stevilske_spremenljivke,
                                        razbitje_spremenljivke = razbitje_spremenljivke,
                                        baza1 = baza1,
                                        use_weights = FALSE)
    
    razbitje_write_excel(sheet_name = "Povprečja - neutežena",
                         list_razbitje = razbitje_list_neutezeno$list_razbitje,
                         p_vrednosti = razbitje_list_neutezeno$p_vrednosti,
                         f_vrednosti = razbitje_list_neutezeno$f_vrednosti,
                         anova_p_vrednosti = razbitje_list_neutezeno$anova_p_vrednosti,
                         rel_razlike = razbitje_list_neutezeno$rel_razlike,
                         wb = wb,
                         utezi = rep(1, nrow(baza1)),
                         use_weights = FALSE,
                         stolpci = razbitje_list_neutezeno$stolpci,
                         color_sig = color_sig)
    
    if(!is.null(utezi_spr)){

      razbitje_list_utezeno <- razbitje(stevilske_spremenljivke = stevilske_spremenljivke,
                                        razbitje_spremenljivke = razbitje_spremenljivke,
                                        baza1 = baza1,
                                        use_weights = TRUE,
                                        utezi_spr = utezi_spr)
      
      razbitje_write_excel(sheet_name = "Povprečja - utežena",
                           list_razbitje = razbitje_list_utezeno$list_razbitje,
                           p_vrednosti = razbitje_list_utezeno$p_vrednosti,
                           f_vrednosti = razbitje_list_utezeno$f_vrednosti,
                           anova_p_vrednosti = razbitje_list_utezeno$anova_p_vrednosti,
                           rel_razlike = razbitje_list_utezeno$rel_razlike,
                           wb = wb,
                           utezi = baza1[[utezi_spr]],
                           use_weights = TRUE,
                           stolpci = razbitje_list_utezeno$stolpci,
                           color_sig = color_sig)
    } 
    
    # povzetek za številske spr.
    rel_razlike_povzetek_neut <- abs(unlist(lapply(razbitje_list_neutezeno$rel_razlike, function(x) x[-c(nrow(x), nrow(x) - 1), ]), use.names = FALSE))
    p_vrednosti_povzetek_neut <- unlist(razbitje_list_neutezeno$p_vrednosti, use.names = FALSE)
    frekvence_rel_razlike_neutezene <- count_rel_diff(vec = rel_razlike_povzetek_neut, p_vec = p_vrednosti_povzetek_neut)
    
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
                         check.names = FALSE)
    
    if(!is.null(utezi_spr)){
      rel_razlike_povzetek_ut <- abs(unlist(lapply(razbitje_list_utezeno$rel_razlike, function(x) x[-c(nrow(x), nrow(x) - 1), ]), use.names = FALSE))
      p_vrednosti_povzetek_ut <- unlist(razbitje_list_utezeno$p_vrednosti, use.names = FALSE)
      frekvence_rel_razlike_utezene <- count_rel_diff(vec = rel_razlike_povzetek_ut, p_vec = p_vrednosti_povzetek_ut)
      
      w_tbl_st <- data.frame(
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
              x = "Primerjava po skupinah za številske (intervalne, razmernostne) spremenljivke", startCol = 2, startRow = 1)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:ncol(tbl_st), rows = 1)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Neutežene statistike", startCol = 2, startRow = 2)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:10, rows = 2)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Skupine", startCol = 2, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:5, rows = 3)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilne primerjave po skupinah (p < 0.05)",
              startCol = 6, startRow = 3)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 6:10, rows = 3)
    
    if(!is.null(utezi_spr)){
      writeData(wb = wb, sheet = "Povzetek",
                x = "Utežene statistike", startCol = 11, startRow = 2)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 11:19, rows = 2)
      
      writeData(wb = wb, sheet = "Povzetek",
                x = "Skupine", startCol = 11, startRow = 3)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 11:14, rows = 3)
      
      writeData(wb = wb, sheet = "Povzetek",
                x = "Od tega statistično značilne primerjave po skupinah (p < 0.05)",
                startCol = 15, startRow = 3)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 15:19, rows = 3)
    }
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. skupnih primerjav povprečij (ANOVA): neuteženi podatki:",
                                sum(!is.na(unlist(razbitje_list_neutezeno$anova_p_vrednosti, use.names = FALSE))),
                                ifelse(!is.null(utezi_spr), paste0("; uteženi podatki: ", sum(!is.na(unlist(razbitje_list_utezeno$anova_p_vrednosti, use.names = FALSE)))))),
                          paste("Od tega statistično značilne primerjave (p < 0.05, vsaj eno povprečje je stat. značilno različno od drugih): neuteženi podatki:",
                                sum(unlist(razbitje_list_neutezeno$anova_p_vrednosti, use.names = FALSE) < 0.05, na.rm = TRUE),
                                ifelse(!is.null(utezi_spr), paste0("; uteženi podatki: ", sum(unlist(razbitje_list_utezeno$anova_p_vrednosti, use.names = FALSE) < 0.05, na.rm = TRUE)), "")))),
              startCol = 1, startRow = 10, rowNames = FALSE, colNames = FALSE)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh posameznih primerjav povprečij skupin s skupnim povprečjem:", sum(frekvence_rel_razlike_neutezene$sums)),
                          paste("Od tega št. statistično značilnih primerjav (p < 0.05) - neuteženi podatki:", sum(frekvence_rel_razlike_neutezene$p_sums)),
                          ifelse(!is.null(utezi_spr), paste("Od tega št. statistično značilnih primerjav (p < 0.05) - uteženi podatki:", sum(frekvence_rel_razlike_utezene$p_sums)), ""))),
              startCol = 1, startRow = 13, rowNames = FALSE, colNames = FALSE)
    
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
    
    if(!is.null(utezi_spr)){
      addStyle(wb = wb, sheet = "Povzetek",
               style = createStyle(fgFill = "#FDE9D9"), rows = 3:8, cols = 11:19,
               gridExpand = TRUE, stack = TRUE)
    }
    
    setColWidths(wb = wb, sheet = "Povzetek", cols = 1:19,
                 widths = c(16, rep(c(5, 6, 8, 8, 5, 21, 21, 8, 9), 2)))
    
    setRowHeights(wb = wb, sheet = "Povzetek", rows = 1:4, heights = c(20, 20, 25, 44))
  }
  
  ### NOMINALNE SPREMENLJIVKE
  
  # preverimo in odstranimo nominalne spremenljivke, ki so konstante ali imajo le NA vrednosti
  indeksi_nom <- vapply(nominalne_spremenljivke, function(x){
    length(unique(baza1[[x]][!is.na(baza1[[x]])])) > 1
  }, FUN.VALUE = logical(1), USE.NAMES = FALSE)
  
  if(any(!indeksi_nom)){
    writeData(wb = wb, sheet = "Opozorila", xy = c(1,5),
              x = paste("Nominalne spremenljivke", paste(nominalne_spremenljivke[!indeksi_nom], collapse = ", "),
                        "se pri analizi niso upoštevale, saj so konstante."))
    
    warning_counter <- TRUE
    
    nominalne_spremenljivke <- nominalne_spremenljivke[indeksi_nom]
  }
  
  if(length(nominalne_spremenljivke) >= 1 && length(razbitje_spremenljivke) >= 1){
    
    # funkcija za izračun Cramer V koeficienta
    # cramer V je ustrezno računati tudi na uteženih podatkih
    # https://stats.stackexchange.com/questions/69661/cram%C3%A9rs-v-on-rao-scott-adjusted-pearson-chi2

    cramersv <- function(x) {
      stat <- as.numeric(x[["statistic"]])
      n <- sum(x[["observed"]])
      k <- min(dim(x[["observed"]]))
      sqrt(stat/(n * (k - 1)))
    }
    
    # funkcija ki prešteje št. relativnih razlik glede na intervale in št. stat. značilnih celic glede na std. reziduale
    count_rel_diff_rez <- function(vec, p_vec) {
      p_vec[is.na(p_vec)] <- 1
      
      frek <- c(sum(vec > 20, na.rm = TRUE), sum(vec > 10 & vec <= 20, na.rm = TRUE), sum(vec >= 5 & vec <= 10, na.rm = TRUE), sum(vec < 5, na.rm = TRUE))
      
      p_frek <- c(sum(vec > 20 & p_vec >= 1.96, na.rm = TRUE), sum(vec > 10 & vec <= 20 & p_vec >= 1.96, na.rm = TRUE), sum(vec >= 5 & vec <= 10 & p_vec >= 1.96, na.rm = TRUE), sum(vec < 5 & p_vec >= 1.96, na.rm = TRUE))
      
      list(sums = frek,
           cumsums = cumsum(frek),
           p_sums = p_frek,
           p_cumsums = cumsum(p_frek))
    }
    
    crosstab <- function(utezi_spr = NULL,
                         use_weights = FALSE,
                         spr_razbitje,
                         nominalna_spr,
                         baza1){
      temp_baza <- baza1[, c(utezi_spr, spr_razbitje, nominalna_spr), drop = FALSE]
      
      cases <- complete.cases(temp_baza[[spr_razbitje]], temp_baza[[nominalna_spr]])
      temp_baza <- temp_baza[cases,]
      
      temp_baza[[spr_razbitje]] <- droplevels(as_factor(temp_baza[[spr_razbitje]]))
      temp_baza[[nominalna_spr]] <- droplevels(as_factor(temp_baza[[nominalna_spr]]))
      
      if ((nlevels(temp_baza[[spr_razbitje]]) >= 2L) && (nlevels(temp_baza[[nominalna_spr]]) >= 2L)){
        
        if (use_weights){
          temp_baza[[utezi_spr]] <- temp_baza[[utezi_spr]]/mean(temp_baza[[utezi_spr]])
          
          survey_design <- svydesign(id = ~1, weights = as.formula(paste0("~", utezi_spr)), data = temp_baza)
          temp_table <- svytable(as.formula(paste("~", spr_razbitje, "+", nominalna_spr)),
                                 design = survey_design)
          
          # ker so uteži normirane so vse komponente enake, le p-vrednost je drugačna
          hi_test <- chisq.test(temp_table, correct = FALSE)
          
          p_vrednost <- tryCatch({
            # Rao-Scott hi-kvadrat
            # The Rao-Scott tests have been shown to have better control of Type I error than the Wald and adjWald tests when the number of design degrees of freedom is small
            # but they may have lower power.
            # https://stats.stackexchange.com/questions/610528/svychisq-with-statistic-chisq-vs-statistic-adjwald
            summary(temp_table, statistic = "Chisq")$statistic$p.value
          }, error = function(e){
            p.value = NA
          })
          
          hi_test$p.value <- p_vrednost
        } else {
          # hi kvadrat test
          hi_test <- chisq.test(temp_baza[[spr_razbitje]], temp_baza[[nominalna_spr]], correct = FALSE)
          
          temp_table <- hi_test$observed
        }
      } else {
        temp_table <- table(temp_baza[[spr_razbitje]], temp_baza[[nominalna_spr]])
      }
      
      temp_df <- as.data.frame(matrix(NA, nrow = (nrow(temp_table) * 3) + 2, ncol = ncol(temp_table) + 3))
      
      # kategorije spremenljivke razbitje
      temp_df[c(seq(1, by = 3, length.out = nrow(temp_table)), nrow(temp_df)-1), 1] <- c(rownames(temp_table), "Skupaj")
      
      # imena vrstic
      temp_df[, 2] <- c(rep(c("N", "%", "Std. rez."), floor(nrow(temp_df)/3)), c("N", "%"))
      
      # kategorije odvisne nominalne spremenljivke
      colnames(temp_df) <- c(paste(spr_razbitje, "-", var_label(baza1[[spr_razbitje]])),
                             "",
                             colnames(temp_table),
                             "Skupaj")
      # frekvence
      temp_df[c(seq(1, by = 3, length.out = nrow(temp_table)), nrow(temp_df) - 1), -c(1,2)] <- round(addmargins(temp_table))
      
      # odstotki po vrsticah
      temp_df[c(seq(2, by = 3, length.out = nrow(temp_table))), -c(1,2)] <- addmargins(prop.table(temp_table, margin = 1), margin = 2)
      temp_df[nrow(temp_df), -c(1,2)] <- c(colSums(temp_table)/sum(temp_table), 1)
      
      # prilagojeni standardizirani reziduali
      res_rows <- c(seq(3, by = 3, length.out = nrow(temp_table)))
      res_cols <- c(1, 2, ncol(temp_df))
      temp_df[res_rows, -res_cols] <- tryCatch({
        round(hi_test$stdres, 2)},
        error = function(e){
          NA
        })
      
      temp_df2 <- temp_df[res_rows, -res_cols]
      
      # relativne razlike deležev
      temp_df_del <- temp_df[c(seq(2, by = 3, length.out = nrow(temp_table)), nrow(temp_df)), -c(1,2)]
      # temp_df_del[temp_df_del == 0] <- NA
      
      for(i in seq_len(ncol(temp_df_del))){
        # temp_df_del[,i] <- (temp_df_del[,i] - temp_df_del[nrow(temp_df_del),i]) * 100 # odst. točke
        temp_df_del[,i] <- ((temp_df_del[,i] - temp_df_del[nrow(temp_df_del),i])/temp_df_del[nrow(temp_df_del),i]) * 100
        
        temp_df[c(seq(2, by = 3, length.out = nrow(temp_table)), nrow(temp_df)), -c(1,2)][[i]] <-
          paste0(round(temp_df[c(seq(2, by = 3, length.out = nrow(temp_table)), nrow(temp_df)), -c(1,2)][[i]] * 100, 1), "%")
      }
      
      temp_df_del <- temp_df_del[-nrow(temp_df_del),-ncol(temp_df_del)]
      
      # približna stat. značilnost rezidualov
      # p < 0.1
      temp_df[res_rows, -res_cols][abs(temp_df2) >= 1.64 & abs(temp_df2) < 1.96] <-
        paste(temp_df2[abs(temp_df2) >= 1.64 & abs(temp_df2) < 1.96], "+")
      
      # p < 0.05
      temp_df[res_rows, -res_cols][abs(temp_df2) >= 1.96 & abs(temp_df2) < 2.58] <-
        paste(temp_df2[abs(temp_df2) >= 1.96 & abs(temp_df2) < 2.58], "*")
      
      # p < 0.01
      temp_df[res_rows, -res_cols][abs(temp_df2) >= 2.58 & abs(temp_df2) < 3.29] <-
        paste(temp_df2[abs(temp_df2) >= 2.58 & abs(temp_df2) < 3.29], "**")
      
      # p < 0.001
      temp_df[res_rows, -res_cols][abs(temp_df2) >= 3.29] <-
        paste(temp_df2[abs(temp_df2) >= 3.29], "***")
      
      # Cramerjev koeficient asociiranosti V
      cramerV <- tryCatch({cramersv(hi_test)},
                          error = function(e){
                            NA})
      
      # predpostavka X2: vsaj 80% teoretičnih frekvenc >= 5 in nobena celica nima < 1
      opozorilo <- tryCatch({
        (mean(hi_test$expected < 5, na.rm = TRUE) >= 0.2) || (min(hi_test$expected, na.rm = TRUE) < 1)},
        error = function(e){
          FALSE})
      
      list(temp_df = temp_df,
           temp_df_del = temp_df_del, # relativne razlike
           temp_df_res = temp_df2, # std. reziduali
           chi_p_value = tryCatch({hi_test$p.value}, error = function(e){NA}),
           cramerV = cramerV,
           opozorilo = opozorilo,
           besedilo = tryCatch({paste0(ifelse(use_weights, "X2 (Rao-Scott) = ", "X2 = "), round(hi_test$statistic[[1]], 2),
                                       ", p = ", paste0(round(hi_test$p.value, 3)),
                                       ifelse(is.na(hi_test$p.value), "", weights::starmaker(hi_test$p.value)),
                                       ifelse(opozorilo && !is.na(hi_test$p.value), " (!)", ""),
                                       ", Cramer V = ", round(cramerV, 3))},
                               error = function(e){""}))
    }
    
    razbitje_write_excel_nom <- function(list_razbitje_nom,
                                         nominalne_spremenljivke,
                                         razbitje_spremenljivke,
                                         baza1,
                                         sheet_name,
                                         wb,
                                         use_weights,
                                         color_sig){
      
      addWorksheet(wb = wb, sheetName = sheet_name, gridLines = FALSE)
      
      n_col_all <- vector("list", length = length(nominalne_spremenljivke))
      opozorilo_v <- vector("list", length = length(nominalne_spremenljivke))
      
      vsota_razbitje <- vector("list", length = length(nominalne_spremenljivke))
      for(i in seq_along(nominalne_spremenljivke)){
        vsota_razbitje[[i]] <- c(0, cumsum(sapply(lapply(list_razbitje_nom[[i]], "[[", "temp_df"), nrow) + 3))
        
        if(i > 1) vsota_razbitje[[i]] <- vsota_razbitje[[i]] + vsota_razbitje[[i-1]][length(vsota_razbitje[[i-1]])]
      }
      
      for(i in seq_along(nominalne_spremenljivke)){
        n_col_i <- c()
        opozorilo_i <- c()
        
        for(j in seq_along(razbitje_spremenljivke)){
          n_col <- ncol(list_razbitje_nom[[i]][[j]]$temp_df)
          n_col_i[j] <- n_col
          opozorilo_i[j] <- list_razbitje_nom[[i]][[j]]$opozorilo
          
          writeData(wb = wb, sheet = sheet_name, startCol = 1, startRow = 12 + vsota_razbitje[[i]][j],
                    x = list_razbitje_nom[[i]][[j]]$temp_df, colNames = TRUE, rowNames = FALSE,
                    borders = "surrounding")
          
          writeData(wb = wb, sheet = sheet_name, startCol = 3, startRow = 11 + vsota_razbitje[[i]][j],
                    x = paste(nominalne_spremenljivke[i], "-", var_label(baza1[[nominalne_spremenljivke[i]]])),
                    colNames = FALSE, rowNames = FALSE,
                    borders = "surrounding")
          
          writeData(wb = wb, sheet = sheet_name, startCol = 1, startRow = 11 + vsota_razbitje[[i]][j],
                    x = list_razbitje_nom[[i]][[j]]$besedilo,
                    colNames = FALSE, rowNames = FALSE,
                    borders = "none")
          
          mergeCells(wb = wb, sheet = sheet_name,
                     cols = 1:2, rows = 11 + vsota_razbitje[[i]][j])
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(textDecoration = "bold"),
                   rows = 11 + vsota_razbitje[[i]][j],
                   cols = 1:2,
                   gridExpand = TRUE, stack = TRUE)
          
          
          mergeCells(wb = wb, sheet = sheet_name,
                     cols = 3:n_col, rows = 11 + vsota_razbitje[[i]][j])
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(fgFill = "#D0CECE"),
                   rows = 11 + vsota_razbitje[[i]][j],
                   cols = 3:n_col,
                   gridExpand = TRUE, stack = TRUE)
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(fgFill = "#E7E6E6"),
                   rows = 12 + vsota_razbitje[[i]][j],
                   cols = 3:n_col,
                   gridExpand = TRUE, stack = TRUE)
          
          
          mergeCells(wb = wb, sheet = sheet_name,
                     cols = 1:2, rows = 12 + vsota_razbitje[[i]][j])
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(fgFill = "#D0CECE"),
                   rows = 12 + vsota_razbitje[[i]][j],
                   cols = 1:2,
                   gridExpand = TRUE, stack = TRUE)
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(fgFill = "#E7E6E6"),
                   rows = (13 + vsota_razbitje[[i]][j]) : (12 + vsota_razbitje[[i]][j] + nrow(list_razbitje_nom[[i]][[j]]$temp_df)),
                   cols = 1:2,
                   gridExpand = TRUE, stack = TRUE)
          
          # obrobe
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(border = c("top")),
                   rows = 11 + vsota_razbitje[[i]][j],
                   cols = 3:n_col,
                   gridExpand = TRUE, stack = TRUE)
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(border = c("top")),
                   rows = 12 + vsota_razbitje[[i]][j],
                   cols = 1:n_col,
                   gridExpand = TRUE, stack = TRUE)
          
          # združi se celice za kategorije
          st_v <- nrow(list_razbitje_nom[[i]][[j]]$temp_df)
          merge_cells_rows <- seq(1, st_v - 2, by = 3)
          
          for(k in seq_along(merge_cells_rows)){
            r_b <- 12 + vsota_razbitje[[i]][j] + merge_cells_rows[k]
            
            mergeCells(wb = wb, sheet = sheet_name,
                       cols = 1,
                       rows = r_b : (r_b + 2))
            
            # obarvanje rel. razlik deležev
            temp_del <- list_razbitje_nom[[i]][[j]]$temp_df_del[k,]
            temp_res <- list_razbitje_nom[[i]][[j]]$temp_df_res[k,]
            
            if(color_sig){ # obarvanje rel. razlik le, če je p < 0.1
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#FBE9E9"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del > 5 & abs(temp_res) >= 1.64) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#F5C0BD"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del > 10 & abs(temp_res) >= 1.64) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#EC8984"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del > 20 & abs(temp_res) >= 1.64) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#C5D9F1"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del < -5 & abs(temp_res) >= 1.64) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#8DB4E2"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del < -10 & abs(temp_res) >= 1.64) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#538DD5"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del < -20 & abs(temp_res) >= 1.64) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
            } else { # obarvanje vseh rel. razlik
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#FBE9E9"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del > 5) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#F5C0BD"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del > 10) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#EC8984"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del > 20) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#C5D9F1"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del < -5) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#8DB4E2"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del < -10) + 2,
                       gridExpand = TRUE, stack = TRUE)
              
              addStyle(wb = wb,
                       sheet = sheet_name,
                       style = createStyle(fgFill = "#538DD5"),
                       rows = 12 + vsota_razbitje[[i]][j] + (merge_cells_rows[k]+1),
                       cols = which(temp_del < -20) + 2,
                       gridExpand = TRUE, stack = TRUE)
            }
          }
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(border = c("bottom")),
                   rows = 11 + vsota_razbitje[[i]][j] + merge_cells_rows,
                   cols = 1:n_col,
                   gridExpand = TRUE, stack = TRUE)
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(border = c("right")),
                   rows = (11 + vsota_razbitje[[i]][j]) : (12 + vsota_razbitje[[i]][j] + nrow(list_razbitje_nom[[i]][[j]]$temp_df)),
                   cols = 1:n_col,
                   gridExpand = TRUE, stack = TRUE)
          
          # združi se celico "Skupaj"
          r_b_s <- 12 + vsota_razbitje[[i]][j] + (st_v - 1)
          
          mergeCells(wb = wb, sheet = sheet_name,
                     cols = 1,
                     rows = r_b_s : (r_b_s + 1))
          
          addStyle(wb = wb, sheet = sheet_name,
                   style = createStyle(border = c("top")),
                   rows = r_b_s,
                   cols = 1:n_col,
                   gridExpand = TRUE, stack = TRUE)
        }
        
        n_col_all[[i]] <- n_col_i
        opozorilo_v[[i]] <- opozorilo_i
      }
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(wrapText = TRUE, valign = "center", halign = "center"),
               rows = 11:(unlist(vsota_razbitje)[length(unlist(vsota_razbitje))] + 50),
               cols = 1:max(unlist(n_col_all), na.rm = TRUE),
               gridExpand = TRUE, stack = TRUE)
      
      writeData(wb = wb, sheet = sheet_name,
                x = ifelse(use_weights,
                           "Utežene kontingenčne tabele",
                           "Neutežene kontingenčne tabele"),
                startCol = 1, startRow = 1)
      
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(textDecoration = "bold",
                                   fontSize = 14),
               rows = 1, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = paste0("Relativna razlika glede na skupni delež stolpca",
                           ifelse(color_sig, " & std. rez. p < 0.1", ""), ":"),
                startCol = 1, startRow = 2)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Visoka pozitivna (> 20 %)",
                startCol = 1, startRow = 3)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Zmerna pozitivna (10 % – 20 %)",
                startCol = 1, startRow = 4)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Šibka pozitivna (5 % – 10 %)",
                startCol = 1, startRow = 5)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Šibka negativna (-5 % – -10 %) ",
                startCol = 1, startRow = 6)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Zmerna negativna (-10 % – -20 %)",
                startCol = 1, startRow = 7)
      writeData(wb = wb,
                sheet = sheet_name,
                x = "Visoka negativna (< -20 %)",
                startCol = 1, startRow = 8)
      
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
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#C5D9F1"),
               rows = 6, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#8DB4E2"),
               rows = 7, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      addStyle(wb = wb, sheet = sheet_name,
               style = createStyle(fgFill = "#538DD5"),
               rows = 8, cols = 1,
               gridExpand = TRUE, stack = TRUE)
      
      writeData(wb = wb,
                sheet = sheet_name,
                x = paste0("Signifikanca prilagojenih standardiziranih rezidualov (std. rez.): + p < 0.1, * p < 0.05, ** p < 0.01, *** p < 0.001",
                           ifelse(any(unlist(opozorilo_v)), "; (!): Več kot 20 % celic ima majhne teoretične frekvence (N<5) ali njihove vrednosti < 1, zato je hi-kvadrat test lahko nezanesljiv.", "")),
                startCol = 1, startRow = 9)
      
      setColWidths(wb = wb, sheet = sheet_name, cols = 1, widths = 38)
    }
    
    # neutežene statistike
    list_razbitje_nom <- vector("list", length(nominalne_spremenljivke))
    names(list_razbitje_nom) <- nominalne_spremenljivke
    
    for(name in nominalne_spremenljivke) {
      list_razbitje_nom[[name]] <- lapply(razbitje_spremenljivke, crosstab,
                                          utezi_spr = NULL, use_weights = FALSE,
                                          nominalna_spr = name, baza1 = baza1)
    }
    
    razbitje_write_excel_nom(list_razbitje_nom = list_razbitje_nom,
                             nominalne_spremenljivke = nominalne_spremenljivke,
                             razbitje_spremenljivke = razbitje_spremenljivke,
                             baza1 = baza1,
                             sheet_name = "Kontingenčne tabele - neutežene",
                             wb = wb,
                             use_weights = FALSE,
                             color_sig = color_sig)
    
    # utežene statistike
    if(!is.null(utezi_spr)){
      list_razbitje_nom_w <- vector("list", length(nominalne_spremenljivke))
      names(list_razbitje_nom_w) <- nominalne_spremenljivke
      
      for(name in nominalne_spremenljivke) {
        list_razbitje_nom_w[[name]] <- lapply(razbitje_spremenljivke, crosstab,
                                              utezi_spr = utezi_spr, use_weights = TRUE,
                                              nominalna_spr = name, baza1 = baza1)
      }
      
      razbitje_write_excel_nom(list_razbitje_nom = list_razbitje_nom_w,
                               nominalne_spremenljivke = nominalne_spremenljivke,
                               razbitje_spremenljivke = razbitje_spremenljivke,
                               baza1 = baza1,
                               sheet_name = "Kontingenčne tabele - utežene",
                               wb = wb,
                               use_weights = TRUE,
                               color_sig = color_sig)
    }
    
    # povzetek za nominalne spr.
    rel_razlike_povzetek_neut_nom <- abs(unlist(lapply(list_razbitje_nom, function(x){
      unlist(lapply(x, "[[", "temp_df_del"), use.names = FALSE)
    }), use.names = FALSE))
    
    p_vrednosti_povzetek_neut_nom <- abs(unlist(lapply(list_razbitje_nom, function(x){
      unlist(lapply(x, "[[", "temp_df_res"), use.names = FALSE)
    }), use.names = FALSE))
    
    frekvence_rel_razlike_neutezene_nom <- count_rel_diff_rez(vec = rel_razlike_povzetek_neut_nom,
                                                              p_vec = p_vrednosti_povzetek_neut_nom)
    
    tbl_nom <- data.frame("Relativne razlike" = c("> 20%", "(10% - 20%]", "[5% - 10%]", "< 5%"),
                          # neutežene statistike
                          "f" = frekvence_rel_razlike_neutezene_nom$sums,
                          "%" = frekvence_rel_razlike_neutezene_nom$sums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                          "Kumul f" = frekvence_rel_razlike_neutezene_nom$cumsums,
                          "Kumul %" = frekvence_rel_razlike_neutezene_nom$cumsums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                          "f*" = frekvence_rel_razlike_neutezene_nom$p_sums,
                          "% od vseh spremenljivk" = frekvence_rel_razlike_neutezene_nom$p_sums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                          "% od relativne razlike" = frekvence_rel_razlike_neutezene_nom$p_sums/frekvence_rel_razlike_neutezene_nom$sums,
                          "Kumul f*" = frekvence_rel_razlike_neutezene_nom$p_cumsums,
                          "Kumul %*" = frekvence_rel_razlike_neutezene_nom$p_cumsums/sum(frekvence_rel_razlike_neutezene_nom$sums),
                          check.names = FALSE)
    
    if(!is.null(utezi_spr)){
      rel_razlike_povzetek_ut_nom <- abs(unlist(lapply(list_razbitje_nom_w, function(x){
        unlist(lapply(x, "[[", "temp_df_del"), use.names = FALSE)
      }), use.names = FALSE))
      
      p_vrednosti_povzetek_ut_nom <- abs(unlist(lapply(list_razbitje_nom_w, function(x){
        unlist(lapply(x, "[[", "temp_df_res"), use.names = FALSE)
      }), use.names = FALSE))
      
      frekvence_rel_razlike_utezene_nom <- count_rel_diff_rez(vec = rel_razlike_povzetek_ut_nom,
                                                              p_vec = p_vrednosti_povzetek_ut_nom)
      
      w_tbl_nom <- data.frame(
        # utežene statistike
        "f" = frekvence_rel_razlike_utezene_nom$sums,
        "%" = frekvence_rel_razlike_utezene_nom$sums/sum(frekvence_rel_razlike_utezene_nom$sums),
        "Kumul f" = frekvence_rel_razlike_utezene_nom$cumsums,
        "Kumul %" = frekvence_rel_razlike_utezene_nom$cumsums/sum(frekvence_rel_razlike_utezene_nom$sums),
        "f*" = frekvence_rel_razlike_utezene_nom$p_sums,
        "% od vseh spremenljivk" = frekvence_rel_razlike_utezene_nom$p_sums/sum(frekvence_rel_razlike_utezene_nom$sums),
        "% od relativne razlike" = frekvence_rel_razlike_utezene_nom$p_sums/frekvence_rel_razlike_utezene_nom$sums,
        "Kumul f*" = frekvence_rel_razlike_utezene_nom$p_cumsums,
        "Kumul %*" = frekvence_rel_razlike_utezene_nom$p_cumsums/sum(frekvence_rel_razlike_utezene_nom$sums),
        check.names = FALSE)
      
      tbl_nom <- cbind(tbl_nom, w_tbl_nom)
    }
    
    tbl_nom[is.na(tbl_nom)] <- 0
    
    writeData(wb = wb,
              sheet = "Povzetek",
              x = tbl_nom,
              borders = "all", startRow = 20, 
              headerStyle = createStyle(textDecoration = "bold",
                                        border = c("top", "bottom", "left", "right"),
                                        halign = "center", valign = "center", wrapText = TRUE))
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Kontingenčne tabele za nominalne spremenljivke", startCol = 2, startRow = 17)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:ncol(tbl_nom), rows = 17)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Neutežene statistike", startCol = 2, startRow = 18)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:10, rows = 18)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Celice", startCol = 2, startRow = 19)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 2:5, rows = 19)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = "Od tega statistično značilni standardizirani reziduali (p < 0.05)",
              startCol = 6, startRow = 19)
    
    mergeCells(wb = wb, sheet = "Povzetek", cols = 6:10, rows = 19)
    
    if(!is.null(utezi_spr)){
      writeData(wb = wb, sheet = "Povzetek",
                x = "Utežene statistike", startCol = 11, startRow = 18)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 11:19, rows = 18)
      
      writeData(wb = wb, sheet = "Povzetek",
                x = "Celice", startCol = 11, startRow = 19)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 11:14, rows = 19)
      
      writeData(wb = wb, sheet = "Povzetek",
                x = "Od tega statistično značilni standardizirani reziduali (p < 0.05)",
                startCol = 15, startRow = 19)
      
      mergeCells(wb = wb, sheet = "Povzetek", cols = 15:19, rows = 19)
    }
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh celic:", length(rel_razlike_povzetek_neut_nom)),
                          paste("Št. statistično značilnih celic (p < 0.05) - neuteženi podatki:", sum(frekvence_rel_razlike_neutezene_nom$p_sums)),
                          ifelse(!is.null(utezi_spr), paste("Št. statistično značilnih celic (p < 0.05) - uteženi podatki:", sum(frekvence_rel_razlike_utezene_nom$p_sums)), ""))),
              startCol = 1, startRow = 26, rowNames = FALSE, colNames = FALSE)
    
    chi_p_values_povzetek_neut <- unlist(lapply(list_razbitje_nom, function(x){
      unlist(lapply(x, "[[", "chi_p_value"), use.names = FALSE)
    }), use.names = FALSE)
    
    writeData(wb = wb, sheet = "Povzetek",
              x = cbind(c(paste("Št. vseh kontingenčnih tabel:", length(chi_p_values_povzetek_neut)),
                          paste("Od tega št. statistično značilnih povezanosti spremenljivk (hi-kvadrat test, p < 0.05) - neuteženi podatki:",
                                sum(chi_p_values_povzetek_neut < 0.05, na.rm = TRUE)))),
              startCol = 1, startRow = 30, rowNames = FALSE, colNames = FALSE)
    
    if(!is.null(utezi_spr)){
      chi_p_values_povzetek_ut <- unlist(lapply(list_razbitje_nom_w, function(x){
        unlist(lapply(x, "[[", "chi_p_value"), use.names = FALSE)
      }), use.names = FALSE)
      
      writeData(wb = wb, sheet = "Povzetek",
                x = paste("Od tega št. statistično značilnih povezanosti spremenljivk (hi-kvadrat test, p < 0.05) - uteženi podatki:",
                          sum(chi_p_values_povzetek_ut < 0.05, na.rm = TRUE)),
                startCol = 1, startRow = 32, rowNames = FALSE, colNames = FALSE)
    }
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 fontSize = 12),
             rows = 17:18, cols = 1:19,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(textDecoration = "bold",
                                 halign = "center", valign = "center",
                                 border = c("top", "bottom", "left", "right")),
             rows = 19, cols = 1:ncol(tbl_nom),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(numFmt = "0%"),
             rows = 21:24, cols = c(3, 5, 7, 8, 10, 12, 14, 16, 17, 19),
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(halign = "center"),
             rows = 21:24, cols = 1:19,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = "Povzetek",
             style = createStyle(fgFill = "#DAEEF3"), rows = 19:24, cols = 2:10,
             gridExpand = TRUE, stack = TRUE)
    
    if(!is.null(utezi_spr)){
      addStyle(wb = wb, sheet = "Povzetek",
               style = createStyle(fgFill = "#FDE9D9"), rows = 19:24, cols = 11:19,
               gridExpand = TRUE, stack = TRUE)
    }
    
    setColWidths(wb = wb, sheet = "Povzetek", cols = 1:19,
                 widths = c(16, rep(c(5, 6, 8, 8, 5, 21, 21, 8, 9), 2)))
    
    setRowHeights(wb = wb, sheet = "Povzetek", rows = 17:20, heights = c(20, 20, 25, 44))
  }
  
  if(warning_counter == FALSE){
    removeWorksheet(wb = wb, sheet = "Opozorila")
  }
  
  saveWorkbook(wb = wb, file = file, overwrite = TRUE)
}