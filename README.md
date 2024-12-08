# Aplikacija za primerjavo neuteženih in uteženih povprečji, deležev in korelacij iz dveh SPSS baz, primerjavo celotnega vzorca s podvzorcem in razbitje

Pred izvedbo analiz se avtomatsko preveri, da so imena izbranih spremenljivk enaka v obeh bazah in, da so enake labele kategorij. V primeru, da ne, se izpišejo opozorila v prvi Excel zavihek. 

## Shiny aplikacija

Preden se aplikacija požene lokalno v programu R, priporočamo, da počistite svoje okolje (Environment), kar je še posebej relevantno, če
se aplikacijo poganja v RStudio, ki privzeto shranjuje okolje iz zadnje seje. Okolje se lahko počisti s kodo `rm(list = ls())` ali z ikono metle v zavihku `Environment` v RStudio.
Pred čiščenjem okolja seveda preverite, da s tem po pomoti ne boste odstranili za vas pomembnega preteklega dela.

Lokalno se požene aplikacijo z naslednjo kodo (za delovanje je potrebno imeti nameščen [R](https://cran.r-project.org/) in pakete `haven`, `labelled`, `weights`, `openxlsx`, `shiny`, `shinyWidgets`, `shinycssloaders`, `shinyjs`, `stringr`, `cocor`, `survey`, `multcomp`, `SimComp`):

```
# namesti se manjkajoče pakete
paketi <- c("haven", "labelled", "weights", "openxlsx", "shiny", "shinyWidgets", "shinycssloaders", "shinyjs", "stringr", "cocor", "survey", "multcomp", "SimComp")

for(p in paketi){
  if(!require(p, character.only = TRUE)) install.packages(p)
}

# požene se aplikacija
shiny::runGitHub(repo = "Weighting", username = "lukastrlekar", ref = "main", subdir = "Shiny App")
```