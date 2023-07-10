data <- haven::read_spss("Test files/Mnenje-potrošnikov-SJM-baza-merged-filtrirano.sav")

ravnotezje_mp <- function(vec){
  mean_prop <- function(x){
    mean(x, na.rm = TRUE)*100
  }
  
  (mean_prop(vec == 5) + 1/2*mean_prop(vec == 4)) - (1/2*mean_prop(vec == 2) + mean_prop(vec == 1))
}

boot <- replicate(1000, ravnotezje_mp(sample(x = data$A1, size = sum(!is.na(data$A1)), replace = TRUE)))

hist(boot, breaks = 100)
mean(boot)
