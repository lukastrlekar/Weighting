source("https://raw.githubusercontent.com/lukastrlekar/Weighting/main/Skripta_primerjava_povprecji.R")

library(here)

podatki1 <- haven::read_spss(file = here("2. CDI, PANDA, 21. val, končano, uteži.sav"),
                             user_na = TRUE)

podatki2 <- haven::read_spss(file = here("3. Valicon, PANDA - 21. val, končano, uteži.sav"),
                             user_na = TRUE)

stevilske_spremenljivke <- c("FINANCE",
                             "SMOKE",
                             "AKTIV",
                             "WHO5_1",
                             "WHO5_2",
                             "WHO5_3",
                             "WHO5_4",
                             "WHO5_5",
                             "STRESS_1",
                             "STRESS_3",
                             "VACCINE_COVID19",
                             "INFECTED",
                             "QUAR_TAKEN_PREBO",
                             "INF_N",
                             "INTENTION_TEST",
                             "VACC_EFFECT_v2",
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

# alternativno:
# stevilske_spremenljivke <- names(podatki1)[c(233:239,241,243)]

nominalne_spremenljivke <- c("FINANCE",
                             "SMOKE",
                             "AKTIV",
                             "WHO5_1",
                             "WHO5_2",
                             "WHO5_3",
                             "WHO5_4",
                             "WHO5_55")


izvoz_excel_tabel(baza1 = podatki1,
                  baza2 = podatki2,
                  ime_baza1 = "CDI PANDA",
                  ime_baza2 = "Valicon PANDA",
                  utezi1 = podatki1$weights,
                  utezi2 = podatki2$weights,
                  stevilske_spremenljivke = stevilske_spremenljivke,
                  nominalne_spremenljivke = nominalne_spremenljivke)
