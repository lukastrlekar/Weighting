library(shiny)
library(shinyWidgets)
library(shinycssloaders)

source("Skripta_primerjava_povprecji.R")

# Define UI 
ui <- fluidPage(

    # Application title
    titlePanel("Primerjava neuteženih in uteženih povprečji in deležev iz dveh SPSS baz"),
    
    tags$a(href = "https://github.com/lukastrlekar/Weighting/blob/main/Statisti%C4%8Dni_izra%C4%8Duni.pdf", target = "_blank",
           "Priloga - statistični izračuni >>>"),
    
    h3("Nalaganje podatkov"),
    br(),
    fluidRow(
      column(4,
             fileInput("upload_baza1", label = "Naloži 1. SPSS bazo", accept = ".sav")),
      column(4,
             fileInput("upload_baza2", label = "Naloži 2. SPSS bazo", accept = ".sav"))),
    
    hr(),
    h3("Izbira spremenljivk"),
    br(),
    fluidRow(
      column(4,
             pickerInput(
               inputId = "stevilske_spr",
               label = HTML("Izberi številske spremenljivke (intervalna, razmernostna merska lestvica): </br>
                    <small>(prikazane so le spremenljivke iz 1. baze)</small>"),
               choices = NULL,
               multiple = TRUE,
               width = "100%",
               options = pickerOptions(actionsBox = TRUE,
                                       liveSearch = TRUE)))),
    fluidRow(
      column(4,
             pickerInput(
               inputId = "nominalne_spr",
               label = HTML("Izberi opisne spremenljivke (nominalna, ordinalna merska lestvica): </br>
                    <small>(prikazane so le spremenljivke iz 1. baze)</small>"),
               choices = NULL,
               multiple = TRUE,
               width = "100%",
               options = pickerOptions(actionsBox = TRUE,
                                       liveSearch = TRUE)))),
    
    hr(),
    h3("Izbira uteži"),
    br(),
    fluidRow(
      column(4,
             pickerInput(
               inputId = "spr_utezi1",
               label = "Izberi spremenljivko uteži iz 1. baze:",
               choices = NULL,
               multiple = TRUE,
               width = "100%",
               options = pickerOptions(actionsBox = TRUE,
                                       liveSearch = TRUE,
                                       maxOptions = 1))),
      column(4,
             pickerInput(
               inputId = "spr_utezi2",
               label = "Izberi spremenljivko uteži iz 2. baze:",
               choices = NULL,
               multiple = TRUE,
               width = "100%",
               options = pickerOptions(actionsBox = TRUE,
                                       liveSearch = TRUE,
                                       maxOptions = 1)))),
    hr(),
    h3("Imena baz"),
    br(),
    fluidRow(
      column(4,
             textInput("ime1", label = "Ime 1. baze prikazano v Excel datoteki:", value = "baza 1")),
      column(4,
             textInput("ime2", label = "Ime 2. baze prikazano v Excel datoteki:", value = "baza 2"))),
    
    hr(),
    h3("Prenos datoteke"),
    br(),
    downloadButton("prenos", label = "Prenesi Excel datoteko", class = "btn-primary", icon = icon("file-excel")),
    br(),
    br()
    
)

# Define server logic 
server <- function(input, output, session) {
  options(shiny.maxRequestSize = 100 * 1024^2)
  
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
  
  output$prenos <- downloadHandler(
    filename = function() {
      paste("Statistike.xlsx")
    },
    content = function(file) {
      # if(!isTRUE(all.equal(mean(podatki1()[[input$spr_utezi1]], na.rm = TRUE), 1))){
      #   showModal(modalDialog(HTML("Povprečje uteži v 1. bazi ni enako 1."),
      #                         easyClose = TRUE,
      #                         footer = NULL))
      # }
      # 
      # if(!isTRUE(all.equal(mean(podatki2()[[input$spr_utezi2]], na.rm = TRUE), 1))){
      #   showModal(modalDialog(HTML("Povprečje uteži v 2. bazi ni enako 1."),
      #                         easyClose = TRUE,
      #                         footer = NULL))
      # }
      showModal(modalDialog(HTML("<h3><center>Prenašanje datoteke</center></h3>"),
                            shinycssloaders::withSpinner(uiOutput("loading"), type = 8),
                            footer = NULL))
      on.exit(removeModal())
      
      izvoz_excel_tabel(baza1 = podatki1(),
                        baza2 = podatki2(),
                        ime_baza1 = input$ime1,
                        ime_baza2 = input$ime2,
                        utezi1 = podatki1()[[input$spr_utezi1]],
                        utezi2 = podatki2()[[input$spr_utezi2]],
                        stevilske_spremenljivke = input$stevilske_spr,
                        nominalne_spremenljivke = input$nominalne_spr,
                        file = file)
    }
  )
}

# Run the application 
shinyApp(ui = ui, server = server)
