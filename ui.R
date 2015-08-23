library(shiny)
library(DBI)
#library(DT)

shinyUI(fluidPage(title="MasterList Curriculum Data Retrieval Interface",
    tags$head(tags$style(".tableClass{float:left;}"),
              tags$script(type="text/javascript", src = "busy.js")),
    h3("CIS Sheet Creation Interface"),
    br(),
    p("Please save the MasterList and CopyFile.xlsm files in the current working directory of this program: "),
    p(getwd()),
    p("The MasterList must be unprotected. Make sure the file has been saved after any manual unprotection before use."),
    textInput(inputId = "year", "Please enter the academic session of the Master List (eg. 2015-2016)"),
    br(), br(), #crude method of adding space to the top
    #well panel separates the buttons from the page intro and the table (if generated)
    wellPanel(
      fluidRow(
        column(1, offset = 3, actionButton('readButton', 'Read MasterList'))
      ),
      fluidRow(
        column(2, offset = 3, textOutput("readSuccess"))
      ),
      uiOutput("selection1"),
      uiOutput("selection2"),
      uiOutput("selection3"),
      uiOutput("selection4")
    ) #end wellPanel
  )#end fluidPage
)#end shinyUI