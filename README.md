# Skripta za primerjavo neuteženih in uteženih povprečji, deležev in korelacij iz dveh SPSS baz

Pred izvedbo analiz koda preveri, da so imena izbranih spremenljivk enaka v obeh bazah in, da so enake labele kategorij. V primeru, da ne, se izpišejo opozorila v prvi Excel zavihek. 

## Shiny aplikacija

Lokalno se požene aplikacijo z naslednjo kodo (za delovanje je potrebno imeti nameščen R in naslednje pakete `haven`, `labelled`, `weights`, `openxlsx`, `shiny`, `shinyWidgets`, `shinycssloaders`, `stringr`, `cocor`, `survey`):

```
# namesti manjkajoče pakete
paketi <- c("haven", "labelled", "weights", "openxlsx", "shiny", "shinyWidgets", "shinycssloaders", "shinyjs", "stringr", "cocor", "survey")

for(p in paketi){
  if(!require(p, character.only = TRUE)) install.packages(p)
}

# poežene se aplikacija
shiny::runGitHub(repo = "Weighting", username = "lukastrlekar", ref = "main", subdir = "Shiny App")
```