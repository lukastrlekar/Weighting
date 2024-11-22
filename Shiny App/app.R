suppressPackageStartupMessages({
  library(shiny)
  library(shinyWidgets)
  library(shinycssloaders)
  library(shinyjs)
  
  library(haven)
  library(labelled)
  library(weights)
  library(openxlsx)
  library(stringr)
  library(cocor)
  library(survey)
  library(multcomp)
  library(SimComp)
})

source("Primerjava_povprecji.R")
source("Primerjava_korelacij.R")
source("Razbitje.R")
source("Matrix_scatter_plot.R")

# Define UI 
ui <- fluidPage(
  useShinyjs(),
  tags$head(
    tags$style(HTML("
      .dropdown-menu>li>a {
        border-bottom: 1px solid rgba(0, 0, 0, .15);
      }"))
  ),
  titlePanel("Primerjava neuteženih in uteženih povprečij, deležev in korelacij iz dveh SPSS baz"),
  br(),
  fluidRow(
    column(12,
           checkboxInput(inputId = "checkbox_vzorec_podvzorec",
                         label = "Primerjava celotnega vzorca s podvzorcem",
                         width = "100%")),
    conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 1",
                     column(12,
                            p(strong("Navodila:"), "naložite eno bazo, ki vsebuje 2 spremenljivki uteži - eno za celoten vzorec in drugo
                       za podvzorec - oboje naj bo uteženo z istimi populacijskimi marginami (spremenljivka uteži za podvzorec naj vsebuje vrednosti le pri enotah, ki spadajo
                       v podvzorec, ki ga želite primerjati s celotnim vzorcem; pri ostalih enotah naj bodo manjkajoče vrednosti)."))),
    
    column(12,
           checkboxInput(inputId = "checkbox_razbitje",
                         label = "Razbitje (analiza povprečij po skupinah kontrolnih spremenljivk za številske spremenljivke in dvodimenzionalne kontingenčne tabele za nominalne spremenljivke)",
                         width = "100%")),
    conditionalPanel(condition = "input.checkbox_razbitje == 1",
                     column(12,
                            p(strong("Opozorilo:"), "v primeru izbranega večjega števila opisnih spremenljivk (> 20) in hkrati tudi spremenljivk razbitja (> 10),
                              lahko izvedba analiz in izvoz datoteke traja dlje časa (> 1 ura). Priporočamo lokalno uporabo, saj se pri aplikaciji na spletu po 15 minutah prekine povezava.")))),
  
  fluidRow(
    column(12,
           hr()),
    column(12,
           h4(strong("Nalaganje podatkov"))),
    column(6,
           fileInput("upload_baza1", label = "Naloži 1. SPSS bazo", accept = ".sav")),
    conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 0 & input.checkbox_razbitje == 0",
                     column(6,
                            fileInput("upload_baza2", label = "Naloži 2. SPSS bazo", accept = ".sav")))),
  fluidRow(
    column(12,
           conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 0 & input.checkbox_razbitje == 0",
                            p("Opomba: relativna pristranskost (relativna razlika) se bo računala kot, da je točkovna ocena v prvi bazi prava.")),
           conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 1",
                            p("Opomba: relativna pristranskost (relativna razlika) se bo računala kot, da je točkovna ocena v celotnem vzorcu prava.")),
           conditionalPanel(condition = "input.checkbox_razbitje == 1",
           ))),
  
  conditionalPanel(condition = "input.checkbox_razbitje == 0",
                   hr(),
                   h4(strong("Imena baz")),
                   fluidRow(
                     column(6,
                            textInput("ime1", label = "Ime 1. baze prikazano v Excel datoteki:", value = "baza 1")),
                     column(6,
                            textInput("ime2", label = "Ime 2. baze prikazano v Excel datoteki:", value = "baza 2")))),
  
  hr(),
  h4(strong("Izbira spremenljivk")),
  conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 0 & input.checkbox_razbitje == 0",
                   p("Opomba: prikazane so le spremenljivke iz 1. baze.")),
  fluidRow(
    column(6,
           pickerInput(
             inputId = "stevilske_spr",
             label = "Izberi številske spremenljivke (intervalna in razmernostna merska lestvica):",
             choices = NULL,
             multiple = TRUE,
             width = "100%",
             options = pickerOptions(actionsBox = TRUE,
                                     liveSearch = TRUE)))),
  fluidRow(
    column(6,
           pickerInput(
             inputId = "nominalne_spr",
             label = "Izberi opisne spremenljivke (nominalna in ordinalna merska lestvica):",
             choices = NULL,
             multiple = TRUE,
             width = "100%",
             options = pickerOptions(actionsBox = TRUE,
                                     liveSearch = TRUE)))),
  conditionalPanel(condition = "input.checkbox_razbitje == 1",
                   fluidRow(
                     column(6,
                            br(),
                            br(),
                            pickerInput(
                              inputId = "razbitje_spr",
                              label = HTML("Izberi spremenljivke za razbitje - kontrolne spremenljivke (nominalna in ordinalna merska lestvica):"),
                              choices = NULL,
                              multiple = TRUE,
                              width = "100%",
                              options = pickerOptions(actionsBox = TRUE,
                                                      liveSearch = TRUE))))),
  hr(),
  tabsetPanel(
    tabPanel(h4(strong("Primerjava povprečij in deležev")),
             br(),
             conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 0 & input.checkbox_razbitje == 0",
                              h4(strong("Izračun standardne napake")),
                              fluidRow(
                                column(12,
                                       radioButtons("se_calculation",
                                                    label = "Izberi način izračuna standardne napake (SE) za utežene statistike:",
                                                    choiceNames = list(HTML('Aproksimativen izračun SE s Taylorjevo linearizacijo za razmernostno cenilko (<a href="https://stats.stackexchange.com/questions/525741/how-to-estimate-the-approximate-variance-of-the-weighted-mean" target="_blank">?</a>)'),
                                                                       HTML('Natančen izračun SE s paketom <tt>survey</tt> (<a href="https://stackoverflow.com/questions/77098620/analyzing-survey-data-in-r-after-creating-raked-weights-with-rake-from-the-r-s" target="_blank">?</a>)')),
                                                    choiceValues = list("taylor_se",
                                                                        "survey_se"), width = "100%")))),
             conditionalPanel(condition = 'input.se_calculation == "survey_se"',
                              hr(),
                              h4(strong("Nalaganje svydesign objektov")),
                              fluidRow(
                                column(6,
                                       fileInput("upload_svydesign_1", label = "Naloži svydesign R objekt za 1. bazo", accept = ".RData")),
                                column(6,
                                       fileInput("upload_svydesign_2", label = "Naloži svydesign R objekt za 2. bazo", accept = ".RData")))),
             
             conditionalPanel(condition = 'input.se_calculation == "taylor_se"',
                              conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 0 & input.checkbox_razbitje == 0", hr()),
                              h4(strong("Izbira uteži")),
                              conditionalPanel(condition = "input.checkbox_razbitje == 1",
                                               p("Opcijsko - če ne želite izračunati uteženih statistik, pustite izbor uteži prazen.")),
                              fluidRow(
                                column(6,
                                       pickerInput(
                                         inputId = "spr_utezi1",
                                         label = "Izberi spremenljivko uteži iz 1. baze:",
                                         choices = NULL,
                                         multiple = TRUE,
                                         width = "70%",
                                         options = pickerOptions(actionsBox = TRUE,
                                                                 liveSearch = TRUE,
                                                                 maxOptions = 1))),
                                conditionalPanel(condition = "input.checkbox_razbitje == 0",
                                                 column(6,
                                                        pickerInput(
                                                          inputId = "spr_utezi2",
                                                          label = "Izberi spremenljivko uteži iz 2. baze:",
                                                          choices = NULL,
                                                          multiple = TRUE,
                                                          width = "70%",
                                                          options = pickerOptions(actionsBox = TRUE,
                                                                                  liveSearch = TRUE,
                                                                                  maxOptions = 1)))))),
             hr(),
             h4(strong("Prenos datoteke")),
             conditionalPanel(condition = "input.checkbox_razbitje == 1",
                              checkboxInput(inputId = "checkbox_color_sig",
                                            label = "Relativne razlike naj se obarvajo le za statistično značilne rezultate (p < 0.1)",
                                            value = TRUE, width = "100%")),
             br(),
             downloadButton("prenos",
                            label = HTML("&nbsp &nbsp Poženi analizo in prenesi Excel datoteko &nbsp &nbsp &nbsp"),
                            class = "btn-primary",
                            icon = icon("play")),
             br(),
             br(),
             p("Opomba: pred analizo se bo uteži reskaliralo, da bo povprečje 1 (ne vpliva na izračun standardnih napak, povprečij in deležev, le na frekvence)."),
             conditionalPanel(condition = "input.checkbox_razbitje != 1",
                              tags$a(href = "https://github.com/lukastrlekar/Weighting/blob/main/Statisti%C4%8Dni_izra%C4%8Duni.pdf", target = "_blank",
                                     "Priloga - statistični izračuni >>>")),
             conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 1",
                              tags$a(href = "https://github.com/lukastrlekar/Weighting/blob/main/Primerjava_vzorec_podvzorec.pdf", target = "_blank",
                                     "Priloga - teoretična izpeljava variance razlike povprečij vzorca in podvzorca >>>")),
             
             conditionalPanel(condition = "input.checkbox_razbitje == 1",
             p(HTML('<strong>Metodološko pojasnilo:</strong>
             
             Za številske spremenljivke se primerja povprečja po skupinah kontrolne spremenljivke. V tem okviru se najprej izvede enosmerna analiza variance (<i>One-Way ANOVA</i>) s predpostavljenimi neenakimi variancami
             po skupinah (t.i. Welcheva ANOVA; <a href="https://doi.org/10.5334/irsp.198" target="_blank">Delacre in drugi 2019</a>), ki testira ničelno domnevo, da so vsa povprečja po skupinah enaka in poda zgolj globalen vpogled,
             ali je kakšno povprečje statistično značilno različno od vsaj enega od ostalih.
             Za informativnejši vpogled se zato testira še ničelne domneve, ali je povprečje posamezne skupine statistično različno od skupnega povprečja vseh skupin (kjer so vključene le veljavne skupine, ki imajo vsaj 2 enoti).
             To se izvede v okviru linearnih modelov s posebej nastavljeno matriko kontrastov (pristop znan kot <i>analysis of means</i> oz. ANOM). Tudi v tem primeru predpostavimo neenake variance po skupinah in sicer z izračunom
             ločenih prostostnih stopenj in kritičnih vrednosti za vsakega od kontrastov. Ker hkrati izvajamo več primerjav, ki so medsebojno odvisne, je potrebno kontrolirati skupno napako 1. vrste (<i>familywise error rate</i> oz. FWER),
             tj. verjetnost ene ali več lažno pozitivnih ugotovitev med vsemi primerjavami, na vnaprej določeni ravni α, običajno 5 %. Tako p-vrednost posamezne primerjave, ki je vezana na t-porazdelitev, prilagodimo na podlagi multivariatne
             t-porazdelitve, kar pomeni, da upoštevamo korelacije med primerjavami in smo s tem pri popravkih manj konzervativni, kot če bi uporabili bolj enostavne metode (npr. Bonferronijev popravek). Teoretičen vpogled in praktični primeri analize
             povprečij so na voljo v <a href="http://dx.doi.org/10.1080/02664763.2015.1117584" target="_blank">Pallmann & Hothorn 2016</a>.
             
             V primeru uteženih podatkov je pristop podoben, le da se standardne napake ocenijo drugače. Za ocenjevanje parametrov je uporabljen paket <a href="https://cran.r-project.org/web/packages/survey/index.html" target="_blank"><tt>survey</tt></a>, ki
             poda asimptotske p-vrednosti (in so posledično tudi popravljene na podlagi multivariatne normalne porazdelitve), variance so privzeto že predpostavljene kot neenake (pristop, podoben pri izračunu robustnih (t.i. sendvič) standardnih napak HC0).
             <br/>
             Nominalne spremenljivke...
             '))),
             br(),
             br()),
    
    tabPanel(h4(strong("Primerjava Pearsonovih korelacij")),
             conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 1 | input.checkbox_razbitje == 1",
                              br(),
                              p("Primerjava korelacij je mogoča le za dva neodvisna vzorca.")
             ),
             conditionalPanel(condition = "input.checkbox_vzorec_podvzorec == 0 & input.checkbox_razbitje == 0",
                              br(),
                              h4(strong("Grafični pregled korelacijske matrike")),
                              p("Matrika razsevnih grafikonov z vrisanimi premicami predpostavke linearne povezanosti (modra brava) in gladilnikom (rdeča barva).
                                Zaradi večje preglednosti so točke zatresene."),
                              br(),
                              fluidRow(
                                column(1, actionButton("show_plot", label = "Prikaži graf", width = "125px")),
                                column(1, actionButton("hide_plot", label = "Skrij graf", width = "125px"))),
                              br(),
                              br(),
                              uiOutput("plot"),
                              
                              hr(),
                              h4(strong("Izbira uteži")),
                              p("Opcijsko - če ne želite izračunati uteženih korelacij, pustite izbor uteži prazen."),
                              fluidRow(
                                column(6,
                                       pickerInput(
                                         inputId = "spr_utezi1_cors",
                                         label = "Izberi spremenljivko uteži iz 1. baze:",
                                         choices = NULL,
                                         multiple = TRUE,
                                         width = "70%",
                                         options = pickerOptions(actionsBox = TRUE,
                                                                 liveSearch = TRUE,
                                                                 maxOptions = 1))),
                                column(6,
                                       pickerInput(
                                         inputId = "spr_utezi2_cors",
                                         label = "Izberi spremenljivko uteži iz 2. baze:",
                                         choices = NULL,
                                         multiple = TRUE,
                                         width = "70%",
                                         options = pickerOptions(actionsBox = TRUE,
                                                                 liveSearch = TRUE,
                                                                 maxOptions = 1)))),
                              hr(),
                              h4(strong("Prenos datoteke")),
                              br(),
                              downloadButton("prenos_korelacije",
                                             label = HTML("&nbsp &nbsp Poženi analizo in prenesi Excel datoteko &nbsp &nbsp &nbsp"),
                                             class = "btn-primary",
                                             icon = icon("play")),
                              br(),
                              br(),
                              p("Če ste izbrali opisne spremenljivke, bodo pri izračunu korelacij upoštevane le dihotomne spremenljivke,
                 saj je izračun Pearsonove korelacije v primeru opisnih spremenljivk mogoč le med številsko in dihotomno spremenljivko (biserialna korelacija)
                 ali med dvema dihotomnima spremenljivkama (koeficient Phi)."),
                              p(HTML('<strong>Metodološko pojasnilo:</strong> preverja se ničelna hipoteza, da je razlika med dvema korelacijskima koeficientoma enaka 0. Ker gre za 2 neodvisna vzorca,
                     je uporabljen Fisherjev <i>z</i> test (1925), implementiran v paketu <a href="http://comparingcorrelations.org/" target="_blank"><tt>cocor</tt></a>.
                    Relativna pomembnost velikosti te razlike je izražena s Cohenovo velikostjo učinka <i>q</i> (Cohenov q) in približnimi smernicami glede njegove
                    intepretacije (<a href="https://doi.org/10.3389/fpsyg.2022.860213" target="_blank">Di Plinio 2022</a>).')),
             ),
             br()))
)

# Define server logic 
server <- function(input, output, session) {
  options(shiny.maxRequestSize = 100 * 1024^2,
          shiny.sanitize.errors = FALSE)
  
  podatki1 <- reactive({
    req(input$upload_baza1)
    
    df1 <- try(labelled::user_na_to_na(haven::read_spss(file = input$upload_baza1$datapath,
                                                        user_na = TRUE)),
               silent = TRUE)
    
    if(!is.data.frame(df1)){
      # issue https://github.com/tidyverse/haven/issues/615
      df1 <- try(labelled::user_na_to_na(haven::read_sav(file = input$upload_baza1$datapath,
                                                         user_na = TRUE,
                                                         encoding = "latin1")),
                 silent = TRUE)
    }
    return(df1)
  })
  
  podatki2 <- reactive({
    req(input$upload_baza2)
    
    df2 <- try(labelled::user_na_to_na(haven::read_spss(file = input$upload_baza2$datapath,
                                                        user_na = TRUE)),
               silent = TRUE)
    
    if(!is.data.frame(df2)){
      # issue https://github.com/tidyverse/haven/issues/615
      df2 <- try(labelled::user_na_to_na(haven::read_sav(file = input$upload_baza2$datapath,
                                                         user_na = TRUE,
                                                         encoding = "latin1")),
                 silent = TRUE)
    }
    return(df2)
  })
  
  observe({
    if(input$checkbox_vzorec_podvzorec == TRUE) {
      updateTextInput(session = session, inputId = "ime1", label = "Ime vzorca prikazano v Excel datoteki:", value = "vzorec")
      updateTextInput(session = session, inputId = "ime2", label = "Ime podvzorca prikazano v Excel datoteki:", value = "podvzorec")
      updatePickerInput(session = session, inputId = "spr_utezi1", label = "Izberi spremenljivko uteži za celoten vzorec:")
      updatePickerInput(session = session, inputId = "spr_utezi2",
                        label = "Izberi spremenljivko uteži za podvzorec (ki je hkrati tudi indikatorska spremenljivka pripadnosti podvzorca):")
      updateRadioButtons(session = session, inputId = "se_calculation", selected = "taylor_se")
      shinyjs::disable(id = "checkbox_razbitje")
    } else if(input$checkbox_razbitje == TRUE) {
      updateRadioButtons(session = session, inputId = "se_calculation", selected = "taylor_se")
      shinyjs::disable(id = "checkbox_vzorec_podvzorec")
    } else {
      updateTextInput(session = session, inputId = "ime1", label = "Ime 1. baze prikazano v Excel datoteki:", value = "baza 1")
      updateTextInput(session = session, inputId = "ime2", label = "Ime 2. baze prikazano v Excel datoteki:", value = "baza 2")
      updatePickerInput(session = session, inputId = "spr_utezi1", label = "Izberi spremenljivko uteži iz 1. baze:")
      updatePickerInput(session = session, inputId = "spr_utezi2", label = "Izberi spremenljivko uteži iz 2. baze:")
      shinyjs::enable(id = "checkbox_razbitje")
      shinyjs::enable(id = "checkbox_vzorec_podvzorec")
    }
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "stevilske_spr",
                      choices = names(podatki1()))
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "nominalne_spr",
                      choices = names(podatki1()))
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "razbitje_spr",
                      choices = names(podatki1()))
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "spr_utezi1",
                      choices = names(podatki1()))
  })
  
  observe({
    if(input$checkbox_vzorec_podvzorec == TRUE) {
      updatePickerInput(session = session,
                        inputId = "spr_utezi2",
                        choices = names(podatki1()))
    } else {
      updatePickerInput(session = session,
                        inputId = "spr_utezi2",
                        choices = names(podatki2()))
    }
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "spr_utezi1_cors",
                      choices = names(podatki1()))
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "spr_utezi2_cors",
                      choices = names(podatki2()))
  })
  
  observeEvent(input$show_plot,{
    shinyjs::show("plot")
    
    output$plot <- renderUI({
      n_panels <- floor(length(input$stevilske_spr)/4)
      if(n_panels == 0 || n_panels == 1){
        n_panels <- 1
        stevilske_spremenljivke_split <- list(input$stevilske_spr)
      } else {
        stevilske_spremenljivke_split <- split(x = input$stevilske_spr,
                                               f = cut(seq_along(input$stevilske_spr), breaks = n_panels, labels = FALSE))
      }
      
      displayed_plots <- lapply(seq_len(n_panels), function(i){
        tabPanel(title = i,
                 fluidRow(
                   br(),
                   column(6,
                          renderText(input$ime1),
                          renderPlot({
                            pairs(podatki1()[, stevilske_spremenljivke_split[[i]], drop = FALSE],
                                  lower.panel = panel.smooth,
                                  upper.panel = panel.cor,
                                  diag.panel = panel.hist)
                          })
                   ), 
                   column(6,
                          renderText(input$ime2),
                          renderPlot({
                            pairs(podatki2()[, stevilske_spremenljivke_split[[i]], drop = FALSE],
                                  lower.panel = panel.smooth,
                                  upper.panel = panel.cor,
                                  diag.panel = panel.hist)
                          }))))})
      
      do.call(tabsetPanel, displayed_plots)
    })
  })
  
  observeEvent(input$hide_plot, {
    shinyjs::hide("plot")
  })
  
  output$prenos <- downloadHandler(
    filename = function() {
      paste("Statistike.xlsx")
    },
    content = function(file) {
      showModal(modalDialog(HTML("<h3><center>Prenašanje datoteke</center></h3>"),
                            shinycssloaders::withSpinner(uiOutput("loading"), type = 8),
                            footer = NULL))
      on.exit(removeModal())
      
      if(input$checkbox_vzorec_podvzorec == FALSE && input$checkbox_razbitje == FALSE) {
        
        if(input$se_calculation == "taylor_se"){
          utezi1 <- podatki1()[[input$spr_utezi1]]
          utezi2 <- podatki2()[[input$spr_utezi2]]
        } else {
          utezi1 <- NULL
          utezi2 <- NULL
        }
        
        izvoz_excel_tabel(baza1 = podatki1(),
                          baza2 = podatki2(),
                          ime_baza1 = input$ime1,
                          ime_baza2 = input$ime2,
                          utezi1 = utezi1,
                          utezi2 = utezi2,
                          stevilske_spremenljivke = input$stevilske_spr,
                          nominalne_spremenljivke = input$nominalne_spr,
                          se_calculation = input$se_calculation,
                          survey_design1 = load_to_environment(input$upload_svydesign_1$datapath),
                          survey_design2 = load_to_environment(input$upload_svydesign_2$datapath),
                          file = file,
                          compare_sample_subsample = FALSE)
        
      } else if(input$checkbox_vzorec_podvzorec == TRUE && input$checkbox_razbitje == FALSE) {
        
        izvoz_excel_tabel(baza1 = podatki1(),
                          baza2 = NULL,
                          ime_baza1 = input$ime1,
                          ime_baza2 = input$ime2,
                          utezi1 = podatki1()[[input$spr_utezi1]],
                          utezi2 = podatki1()[[input$spr_utezi2]],
                          stevilske_spremenljivke = input$stevilske_spr,
                          nominalne_spremenljivke = input$nominalne_spr,
                          se_calculation = "taylor_se",
                          file = file,
                          compare_sample_subsample = TRUE)
        
      } else if(input$checkbox_vzorec_podvzorec == FALSE && input$checkbox_razbitje == TRUE) {
        
        if((length(input$spr_utezi1) == 1)){
          utezi_spr <- input$spr_utezi1
        } else {
          utezi_spr <- NULL
        }
        
        izvoz_excel_razbitje(baza1 = podatki1(),
                             utezi_spr = utezi_spr,
                             stevilske_spremenljivke = input$stevilske_spr,
                             nominalne_spremenljivke = input$nominalne_spr,
                             razbitje_spremenljivke = input$razbitje_spr,
                             file = file,
                             color_sig = input$checkbox_color_sig)
      }
    }
  )
  
  output$prenos_korelacije <- downloadHandler(
    filename = function() {
      paste("Primerjava_korelacij.xlsx")
    },
    content = function(file) {
      
      showModal(modalDialog(HTML("<h3><center>Prenašanje datoteke</center></h3>"),
                            shinycssloaders::withSpinner(uiOutput("loading"), type = 8),
                            footer = NULL))
      on.exit(removeModal())
      
      if((length(input$spr_utezi1_cors) == 1) && (length(input$spr_utezi2_cors) == 1)){
        utezi1 <- podatki1()[[input$spr_utezi1_cors]]
        utezi2 <- podatki2()[[input$spr_utezi2_cors]]
      } else {
        utezi1 <- NULL
        utezi2 <- NULL
      }
      
      izvoz_excel_korelacije(baza1 = podatki1(),
                             baza2 = podatki2(),
                             ime_baza1 = input$ime1,
                             ime_baza2 = input$ime2,
                             utezi1 = utezi1,
                             utezi2 = utezi2,
                             stevilske_spremenljivke = input$stevilske_spr,
                             nominalne_spremenljivke = input$nominalne_spr,
                             file = file)
    }
  )
}

# Run the application 
shinyApp(ui = ui, server = server)
