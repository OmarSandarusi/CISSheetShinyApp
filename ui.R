
#
#Author: Omar Sandarusi
#Queen's University
#September 2015
#
#This is the ui.R file for the Shiny application "CISSheetShinyApp".
#The general purpose of this app is to pull data from master list excel files that contain the finalized
#curriculum course data, process the data by cutting out deleted courses and uneeded rows, and then generate 
#CIS sheets for whichever departments/courses the user chooses.
#

require(shiny)

shinyUI(fluidPage(title="CIS Sheet Shiny App",
    #source the busy.js script file that interacts with the 'busy' class in selectionRow4/genSuccess
    tags$head(tags$script(type="text/javascript", src = "busy.js")),
    h3("CIS Sheet Creation Interface"),
    p("This app pulls course data out of finalized MasterList files, which are created using Jim Mason's Excel Macro-Enabled files."),
    br(),
    p("Ensure that the MasterList, excelScript.R, CopyFile.xlsm, and 3.1.1_3.1.2_A6C.xlsm files have been saved in the current working directory of this program: "),
    p(getwd()),
    p("A folder 'www' must be present in the directory as well, and it should contain the busy.js javascript file. 
      Please ensure that no duplicates of the above files exist. Any previous output files in this directory will be overwritten if they are re-generated during this session."),
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
      uiOutput("selectionRow1"), #selection lists (deptSelect, courseSelection)
      uiOutput("selectionRow2"), #'select all' boxes (allBox, courseSelectionAll)
      uiOutput("selectionRow3"), #genButton
      uiOutput("selectionRow4")  #genButton text (genSuccess, working div)
    ) #end wellPanel
  )#end fluidPage
)#end shinyUI