library(shiny)
library(DBI)
library(excel.link)
source("excelScript.R", local = TRUE)


#
#
#
#Do We not create CIS sheets for deleted course entries?
#
#
#




trim <- function (x) gsub("(^ +)|( +$)", "", x)#gsub("^\\s+|\\s+$", "", x)
#adds whitespace back to names
addWhiteUnd <- function (x) gsub("_", " ", x, fixed = TRUE)

shinyServer(function(input, output, session){
  
  generated <- reactiveValues()
  generated$gen <- FALSE                        #used after gen button click to check whether retrieval was successful
  generated$read <- FALSE                       #logic value for read button
  generated$gcount <- 0                          #count of gen button clicks used to control fields updating
  generated$rcount <- 0                         #count for read button
  generated$gsuccess <- "@#%^No Change Made#$%^" #placeholder for success/failure value to be printed on the click of the gen button
  generated$rsuccess <- "#^#^#^NO Change Made^^^^^@#$#@"
  generated$index <- data.frame()               #holds indicies of selected courses
  generated$testt <- "testt"
  
  com <- xl.get.excel() #COMIDispatch object that points to excel
  dir <- getwd()
  #reactive data that stores the masterList
  viewData <- reactiveValues(
    masterList = list()
  )
  
  #Reactivity to button clicks
  observe({
    #check if read button has been pressed
    if (input$readButton > generated$rcount) {
      if (input$year != "") {
        if (file.exists(paste0("Course Master List - ", input$year, ".xlsm"))) {
          viewData$masterList <- loadMaster(dir, input$year)
          generated$read <- TRUE
          generated$rsuccess <- "Success!"
        } else {
          generated$read <- FALSE
          generated$rsuccess <- "Master List specified by the session does not exist in the current directory."
        }
      } else {
        generated$read <- FALSE
        generated$rsuccess <- "Please enter a valid session, eg. \"2015-2016\""
      }
      output$readSuccess <- renderText({
        generated$rsuccess
      })
      generated$rcount <- generated$rcount + 1
    }
    #gen button reactivity
    if (generated$read == TRUE){
       if (!is.null(input$genButton) && !is.null(input$deptSelect)) {
        if (input$genButton > generated$gcount){
          if (length(viewData$masterList) < 1) { #checking existence of pulled data
            generated$gen <- FALSE
            generated$gsuccess <- "NULL/invalid Data in viewData$masterList, possible Error"
          } else if (input$deptSelect == "Choose" && input$allBox == FALSE) {
            generated$gen <- FALSE
            generated$gsuccess <- "Please choose a Department."
          } else if (input$allBox == TRUE) { #all departments and all courses
            for (i in 1:length(viewData$masterList)) { #for every department
              genCISsheets(viewData$masterList[i], dir, com, input$year)
            }
            generated$testt <- "all generated"
            generated$gen <- TRUE
            generated$gsuccess <- paste("Success!")
          } else if (length(input$courses) < 1 && input$courseAllBox == FALSE) {#no courses selected
            generated$gen <- FALSE
            generated$gsuccess <- "Please Select at least one course."
          } else {#generate CIS sheets based on the selected dept. and courses
            generated$testt <- paste("department:", input$deptSelect, "/////course 1:", input$courses[1])
            indecies <- c()
            for (i in 1:length(input$courses)) {#for each selected course
              j <- 1
              #check what value in the list of courses matches the current selected course
              while (j < length(generated$index[,"index"]) && generated$index[j,"vals"] != input$courses[i]) {
                j <- j + 1
              }
              indecies <- c(indecies, generated$index[j,"index"])
            }#now generate CIS sheets on the subset of the masterList specified by the indicies
            genCISsheets(viewData$masterList[[trimWhiteUnd(input$deptSelect)]][indecies,], dir, com, input$year)
            generated$gen <- TRUE
            generated$gsuccess <- paste("Success!")
          }
          output$genSuccess <- renderText({
            generated$gsuccess
          })
          generated$gcount <- generated$gcount + 1
        }
      }
    }#end gen button reactiity
  })#end observe
  #selection lists
  output$selection1 <- renderUI({
    if (generated$read == TRUE) {
      #department selection list
      choice <- c("Choose", addWhiteUnd(names(viewData$masterList)))
      fluidRow(
        column(3, offset = 1, 
               selectInput("deptSelect", label = "Select a Department",
                                          choices = choice, 
                                          selected = "Choose", 
                                          multiple = FALSE)),
        column(1),
        column(5, uiOutput("courseSelection"))
      )
    } else {}
  })
  #department all box
  output$selection2 <- renderUI({
    if (generated$read == TRUE) {
      fluidRow(
        column(3, offset = 1, checkboxInput("allBox", label = "Select all Departments and Courses")),
        column(1),
        column(3, uiOutput("courseSelectionAll"))
      )
    } else {}
  })
  #gen button
  output$selection3 <- renderUI({
     if (generated$read == TRUE) {
       fluidRow(
        column(1, offset = 3, actionButton('genButton', 'Generate CIS Sheets'))
       )
     } else {}
  })
  #gen button text
  output$selection4 <- renderUI({
    if (generated$read == TRUE) {
      fluidRow(
        column(3, offset = 3, textOutput('genSuccess'))
      )
    } else {}
  })
  #course select menu
  output$courseSelection <- renderUI({
    if (generated$read == TRUE) {
      if (!is.null(input$allBox)) {
        if (input$allBox == TRUE) {} #return no UI
        else if (input$deptSelect == "Choose") {
           selectInput("courses", "Select a Department first", choices = NULL)
        }
        else if (length(input$deptSelect) > "") {
          courseCode <- viewData$masterList[[trimWhiteUnd(input$deptSelect)]][,1]
          courseTitle <- viewData$masterList[[trimWhiteUnd(input$deptSelect)]][,2]
          vals <- c()
          index <- c()
          for (i in 1:length(courseCode)) {
            if (!is.na(courseCode[i])) {
              vals <- c(vals, paste(courseCode[i], courseTitle[i]))
              index <- c(index, i)
            }
          }
          generated$index <- data.frame(vals, index)
          selectInput("courses", "Choose Specific Courses", 
                      choices = vals, multiple = TRUE, selectize = TRUE)
        }
      }
    }
  })
  #course selection all box
  output$courseSelectionAll <- renderUI({
    if (generated$read == TRUE) {
      if (!is.null(input$allBox)) {
        if (input$allBox == TRUE) {} #return no UI
        else {
          checkboxInput("courseAllBox", "All Courses")
        }
      }
    }
  })
  output$test <- renderText({generated$testt})
  #})#end observe testing
})#end shinyServer