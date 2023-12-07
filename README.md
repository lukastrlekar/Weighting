# Skripta za primerjavo neuteženih in uteženih povprečji, deležev in korelacij iz dveh SPSS baz

Pred izvedbo analiz koda preveri, da so imena izbranih spremenljivk enaka v obeh bazah in, da so enake labele kategorij. V primeru, da ne, se izpišejo opozorila v prvi Excel zavihek. 

## Shiny aplikacija

Lokalno se požene aplikacijo z naslednjo kodo (za delovanje so potrebni paketi `haven`, `labelled`, `weights`, `openxlsx`, `shiny`, `shinyWidgets`, `shinycssloaders`, `stringr`, `cocor`):

```
shiny::runGitHub(repo = "Weighting", username = "lukastrlekar", ref = "main")
```