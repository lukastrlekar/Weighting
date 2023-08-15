# Skripta za primerjavo neuteženih in uteženih povprečji in deležev iz dveh SPSS baz

Pred izvedbo analiz koda preveri, da so imena izbranih spremenljivk enaka v obeh bazah in, da so enake labele kategorij. V primeru, da ne, se izpišejo opozorila v prvi Excel zavihek. 

## Shiny aplikacija

Lokalno se požene aplikacijo z naslednjo kodo:

```
shiny::runGitHub(repo = "Weighting", username = "lukastrlekar", ref = "main")
```

## R funkcija

### Navodila za uporabo

Za izvedbo skripte se pokliče funkcijo `izvoz_excel_tabel()`, ki sprejme naslednje argumente:

- `baza1` (podatkovni okvir 1. baze prebran s funkcijo `read_spss()` iz paketa `haven`)
- `baza2` (podatkovni okvir 2. baze prebran s funkcijo `read_spss()` iz paketa `haven`)
- `ime_baza1` (opcijsko ime 1. baze, privzeto je `"baza 1"`)
- `ime_baza2` (opcijsko ime 2. baze, privzeto je `"baza 2"`)
- `utezi1` (vektor uteži za 1. bazo)
- `utezi2` (vektor uteži za 2. bazo)
- `stevilske_spremenljivke` (vektor imen spremenljivk, ki naj se jih obravnava kot številske - računa se povprečje, lahko se izpusti)
- `nominalne_spremenljivke` (vektor imen spremenljivk, ki naj se jih obravnava kot nominalne - prikaže se jih v frekvenčnih tabelah, lahko se izpusti)
- `file` (ime datoteke)

Funkcija vrne Excel datoteko z imenom `"Statistike.xlsx"`, ki se shrani v mapo aktivnega direktorija (aktivni direktorij si lahko nastavite v R Studiu pod Session > Set Working Directory).

### Primer uporabe

```
# najprej se pokliče izvorno funkcijo
source("https://raw.githubusercontent.com/lukastrlekar/Weighting/main/Skripta_primerjava_povprecji.R")

# naložimo obe SPSS datoteki
# argument user_na = TRUE !

podatki1 <- haven::read_spss(file = "C:/Users/strle/OneDrive/Dokumenti/Delo/Weighting/2. CDI, PANDA, 21. val, koncano, uteži.sav",
                             user_na = TRUE)

podatki2 <- haven::read_spss(file = "C:/Users/strle/OneDrive/Dokumenti/Delo/Weighting/3. Valicon, PANDA - 21. val, končano, uteži.sav",
                             user_na = TRUE)

# izberemo številske spremenljivke
stevilske_spremenljivke <- c("VACC_EFFECT_v2",
                             "VACC_NATURAL",
                             "VACC_OBLIGATION",
                             "TRUST_JOURNAL",
                             "TRUST_HOSPITALS",
                             "TRUST_NCDC",
                             "TRUST_POLITICIANS",
                             "TRUST_DOCTOR",
                             "TRUST_NATIONAL_HEALTH",
                             "TRUST_SCIENCE",
                             "TRUST_POSSK19")

# alternativno, na primer:
# stevilske_spremenljivke <- names(podatki1)[c(233:239,241,243)]

# izberemo nominalne spremenljivke
nominalne_spremenljivke <- c("FINANCE",
                             "SMOKE",
                             "AKTIV",
                             "WHO5_1",
                             "WHO5_2",
                             "WHO5_3",
                             "WHO5_4",
                             "WHO5_55")

# pokličemo funkcijo
izvoz_excel_tabel(baza1 = podatki1,
                  baza2 = podatki2,
                  ime_baza1 = "CDI PANDA",
                  ime_baza2 = "Valicon PANDA",
                  utezi1 = podatki1$weights,
                  utezi2 = podatki2$weights,
                  stevilske_spremenljivke = stevilske_spremenljivke,
                  nominalne_spremenljivke = nominalne_spremenljivke,
                  file = "Statistike.xlsx")
```

