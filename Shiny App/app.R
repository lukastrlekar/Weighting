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

source("Primerjava_povprecji.R")
source("Primerjava_korelacij.R")
source("Matrix_scatter_plot.R")

# Define UI 
ui <- fluidPage(
  useShinyjs(),
  titlePanel("Primerjava neuteženih in uteženih povprečji, deležev in korelacij iz dveh SPSS baz"),
  br(),
  h4(strong("Nalaganje podatkov")),
  fluidRow(
    column(4,
           fileInput("upload_baza1", label = "Naloži 1. SPSS bazo", accept = ".sav")),
    column(4,
           fileInput("upload_baza2", label = "Naloži 2. SPSS bazo", accept = ".sav"))),
  fluidRow(
    column(8,
           p("Opomba: relativna pristranskost (relativna razlika) se bo računala kot, da je točkovna ocena v prvi bazi prava."))),
  hr(),
  h4(strong("Imena baz")),
  fluidRow(
    column(4,
           textInput("ime1", label = "Ime 1. baze prikazano v Excel datoteki:", value = "baza 1")),
    column(4,
           textInput("ime2", label = "Ime 2. baze prikazano v Excel datoteki:", value = "baza 2"))),
  
  hr(),
  h4(strong("Izbira spremenljivk")),
  fluidRow(
    column(6,
           pickerInput(
             inputId = "stevilske_spr",
             label = HTML("Izberi številske spremenljivke (intervalna in razmernostna merska lestvica):</br>
                                        <small>(prikazane so le spremenljivke iz 1. baze)</small>"),
             choices = NULL,
             multiple = TRUE,
             width = "100%",
             options = pickerOptions(actionsBox = TRUE,
                                     liveSearch = TRUE)))),
  fluidRow(
    column(6,
           pickerInput(
             inputId = "nominalne_spr",
             label = HTML("Izberi opisne spremenljivke (nominalna in ordinalna merska lestvica):</br>
                                        <small>(prikazane so le spremenljivke iz 1. baze)</small>"),
             choices = NULL,
             multiple = TRUE,
             width = "100%",
             options = pickerOptions(actionsBox = TRUE,
                                     liveSearch = TRUE)))),
  hr(),
  tabsetPanel(
    tabPanel(h4(strong("Primerjava povprečji in deležev")),
             br(),
             h4(strong("Izračun standardne napake")),
             fluidRow(
               column(12,
                      radioButtons("se_calculation",
                                   label = "Izberi način izračuna standardne napake (SE):",
                                   choiceNames = list(HTML('Aproksimativen izračun SE s Taylorjevo linearizacijo za razmernostno cenilko (<a href="https://stats.stackexchange.com/questions/525741/how-to-estimate-the-approximate-variance-of-the-weighted-mean" target="_blank">?</a>)'),
                                                      HTML('Natančen izračun SE s paketom <tt>survey</tt> (<a href="https://stackoverflow.com/questions/77098620/analyzing-survey-data-in-r-after-creating-raked-weights-with-rake-from-the-r-s" target="_blank">?</a>)')),
                                   choiceValues = list("taylor_se",
                                                       "survey_se"), width = "100%"))),
             conditionalPanel(condition = 'input.se_calculation == "survey_se"',
                              hr(),
                              h4(strong("Nalaganje svydesign objektov")),
                              fluidRow(
                                column(4,
                                       fileInput("upload_svydesign_1", label = "Naloži svydesign R objekt za 1. bazo", accept = ".RData")),
                                column(4,
                                       fileInput("upload_svydesign_2", label = "Naloži svydesign R objekt za 2. bazo", accept = ".RData")))),
             
             conditionalPanel(condition = 'input.se_calculation == "taylor_se"',
                              hr(),
                              h4(strong("Izbira uteži")),
                              fluidRow(
                                column(4,
                                       pickerInput(
                                         inputId = "spr_utezi1",
                                         label = "Izberi spremenljivko uteži iz 1. baze:",
                                         choices = NULL,
                                         multiple = TRUE,
                                         width = "70%",
                                         options = pickerOptions(actionsBox = TRUE,
                                                                 liveSearch = TRUE,
                                                                 maxOptions = 1))),
                                column(4,
                                       pickerInput(
                                         inputId = "spr_utezi2",
                                         label = "Izberi spremenljivko uteži iz 2. baze:",
                                         choices = NULL,
                                         multiple = TRUE,
                                         width = "70%",
                                         options = pickerOptions(actionsBox = TRUE,
                                                                 liveSearch = TRUE,
                                                                 maxOptions = 1))))),
             hr(),
             h4(strong("Prenos datoteke")),
             br(),
             downloadButton("prenos",
                            label = HTML("&nbsp &nbsp Prenesi Excel datoteko &nbsp &nbsp &nbsp"),
                            class = "btn-primary",
                            icon = icon("file-excel")),
             br(),
             br(),
             p("Opomba: pred analizo se bo uteži reskaliralo, da bo povprečje 1 (ne vpliva na izračun standardnih napak, povprečji in deležev, le na frekvence)."),
             tags$a(href = "https://github.com/lukastrlekar/Weighting/blob/main/Statisti%C4%8Dni_izra%C4%8Duni.pdf", target = "_blank",
                    "Priloga - statistični izračuni >>>"),
             br(),
             br()),
    
    tabPanel(h4(strong("Primerjava Pearsonovih korelacij")),
             br(),
             h4(strong("Grafični pregled korelacijske matrike (v izdelavi...)")),
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
             HTML("<h4><strong>Izbira uteži </strong><small>(opcijsko - če ne želite izračunati uteženih korelacij, pustite izbor uteži prazen)</small></h4>"),
             fluidRow(
               column(4,
                      pickerInput(
                        inputId = "spr_utezi1_cors",
                        label = "Izberi spremenljivko uteži iz 1. baze:",
                        choices = NULL,
                        multiple = TRUE,
                        width = "70%",
                        options = pickerOptions(actionsBox = TRUE,
                                                liveSearch = TRUE,
                                                maxOptions = 1))),
               column(4,
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
                            label = HTML("&nbsp &nbsp Prenesi Excel datoteko &nbsp &nbsp &nbsp"),
                            class = "btn-primary",
                            icon = icon("file-excel")),
             br(),
             br(),
             p("Če ste izbrali opisne spremenljivke, bodo pri izračunu korelacij upoštevane le dihotomne spremenljivke,
                 saj je izračun Pearsonove korelacije v primeru opisnih spremenljivk mogoč le med številsko in dihotomno spremenljivko (biserialna korelacija)
                 ali med dvema dihotomnima spremenljivkama (koeficient Phi)."),
             p(HTML('Preverja se ničelna hipoteza, da je razlika med dvema korelacijskima koeficientoma enaka 0. Ker gre za 2 neodvisna vzorca,
                     je uporabljen Fisherjev <i>z</i> test (1925), implementiran v paketu <a href="http://comparingcorrelations.org/" target="_blank"><tt>cocor</tt></a>.
                    Relativna pomembnost velikosti te razlike je izražena s Cohenovo velikostjo učinka <i>q</i> (Cohenov q) in približnimi smernicami glede njegove
                    intepretacije (<a href="https://doi.org/10.3389/fpsyg.2022.860213" target="_blank">Di Plinio 2022</a>).')),
             br()))
)

# Define server logic 
server <- function(input, output, session) {
  options(shiny.maxRequestSize = 100 * 1024^2,
          shiny.sanitize.errors = FALSE)
  
  podatki1 <- reactive({
    req(input$upload_baza1)
    
    labelled::user_na_to_na(haven::read_spss(file = input$upload_baza1$datapath,
                                             user_na = TRUE))
  })
  
  podatki2 <- reactive({
    req(input$upload_baza2)
    
    labelled::user_na_to_na(haven::read_spss(file = input$upload_baza2$datapath,
                                             user_na = TRUE))
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
                      inputId = "spr_utezi1",
                      choices = names(podatki1()))
  })
  
  observe({
    updatePickerInput(session = session,
                      inputId = "spr_utezi2",
                      choices = names(podatki2()))
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
                        file = file)
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
