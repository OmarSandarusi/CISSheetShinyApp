
#
#Author: Omar Sandarusi
#Queen's University
#September 2015
#
#This is the server.R file for the Shiny application "CISSheetShinyApp".
#The general purpose of this app is to pull data from master list excel files that contain the finalized
#curriculum course data, process the data by cutting out deleted courses and uneeded rows, and then generating 
#CIS sheets for whichever departments/courses the user chooses.
#

require(shiny)
require(excel.link)
source("excelScript.R", local = TRUE)

#trim <- function (x) gsub("(^ +)|( +$)", "", x)
#adds whitespace back to names
addWhiteUnd <- function (x) gsub("_", " ", x, fixed = TRUE)

#creates a department-specific file and file name, returns the name. Existing files WILL be overwritten.
genFile <- function(masterName) {
  file <- paste0(masterName, "_", "3.1.1_3.1.2_A6C.xlsm")
  file.copy("3.1.1_3.1.2_A6C.xlsm", file, overwrite = TRUE)
  return (file)
}

shinyServer(function(input, output, session){
  #reactive data that tracks the logic values associated with reacting to the read and generate buttons
  generated <- reactiveValues()
  generated$gen <- FALSE                             #used after gen button click to check whether retrieval was successful
  generated$read <- FALSE                            #similar logic value for read button
  generated$gcount <- 0                              #count of gen button clicks used to control fields updating
  generated$rcount <- 0                              #similar count for read button
  generated$gsuccess <- "@#%^No Change Made#$%^"     #placeholder for success/failure value to be printed on the click of the gen button
  generated$rsuccess <- "#^#^#^NO Change Made^@#$#@" #similar placeholder for read button
  
  com <- xl.get.excel() #COMIDispatch object that points to excel
  dir <- getwd()
  #reactive data that stores the masterList and the selected courses
  viewData <- reactiveValues(
    masterList = list(),
    index = data.frame()
  )
  
  #Reactivity to button clicks
  observe({
    #check if read button has been pressed
    if (input$readButton > generated$rcount) {
      #check if the user has specified the session of the master list they wish to read
      if (input$year != "") {
        #check existence of specified file and load its data if it exists
        if (file.exists(paste0("Course Master List - ", input$year, ".xlsm"))) {
          viewData$masterList <- loadMaster(dir, input$year)
          generated$read <- TRUE
          generated$rsuccess <- "Success!"
        } else {
          generated$read <- FALSE
          generated$rsuccess <- "Master List specified by the session does not exist in the current directory."
        }
      } else { #when the input year is empty
        generated$read <- FALSE
        generated$rsuccess <- "Please enter a valid session, eg. \"2015-2016\""
      }
      #update the text below the read button to the generated message
      output$readSuccess <- renderText({
        generated$rsuccess
      })
      #increment the count to block this loop until the read button is pressed again
      generated$rcount <- generated$rcount + 1
    }
    #gen button reactivity, CIS sheets are generated if the button is clicked while valid selections have been made
    #first check if the file has been successfully read
    if (generated$read == TRUE){
       #ensure that the program has built the gen button and the drop-down department list
       if (!is.null(input$genButton) && !is.null(input$deptSelect)) {
        #check if the gen button has been pressed
        if (input$genButton > generated$gcount){
          #check that the master list has been populated, otherwise go through possible input combinations
          if (length(viewData$masterList) < 1) { 
            generated$gen <- FALSE
            generated$gsuccess <- "NULL/invalid Data in viewData$masterList, possible Error"
          } 
          #Option 1: No department chosen
          else if (input$deptSelect == "Choose" && input$allBox == FALSE) { 
            generated$gen <- FALSE
            generated$gsuccess <- "Please choose a Department."
          } 
          #Option 2: All departments and all courses selected
          else if (input$allBox == TRUE) { 
            for (i in 1:length(viewData$masterList)) { #for every department
              genCISsheets(viewData$masterList[[i]], dir, com, input$year, genFile(names(viewData$masterList)[i]))
            }
            generated$gen <- TRUE
            generated$gsuccess <- "Success!"
          } 
          #Option 3: No courses selected
          else if (length(input$courses) < 1 && input$courseAllBox == FALSE) {
            generated$gen <- FALSE
            generated$gsuccess <- "Please Select at least one course."
          } 
          #Option 4: All Courses selected from one department
          else if (input$courseAllBox == TRUE) { 
            trim <- trimWhiteUnd(input$deptSelect)
            genCISsheets(viewData$masterList[[trim]], dir, com, input$year, genFile(trim))
            generated$gsuccess <- "Success!"
          } 
          #Option 5: Some courses selected from one department
          else {
            indecies <- c()
            for (i in 1:length(input$courses)) {#for each selected course
              j <- 1
              #check what value in the list of courses matches the current selected course
              while (j < length(viewData$index[,"index"]) && viewData$index[j,"vals"] != input$courses[i]) {
                j <- j + 1
              }
              indecies <- c(indecies, viewData$index[j,"index"])
            }#now generate CIS sheets on the subset of the masterList specified by the indicies
            trim <- trimWhiteUnd(input$deptSelect)
            genCISsheets(viewData$masterList[[trim]][indecies,], dir, com, input$year, genFile(trim))
            generated$gen <- TRUE
            generated$gsuccess <- paste("Success!")
          }
          #update the gen button text to the generated value
          output$genSuccess <- renderText({
            generated$gsuccess
          })
          #increment the count so that the loop is closed until the gen button is pressed again
          generated$gcount <- generated$gcount + 1
        }
      }
    }#end gen button reactiity
    #reset the department selection list if the select all box has been checked
    if (generated$read == TRUE && !is.null(input$allBox)) {
      if (input$allBox == TRUE) {
        updateSelectInput(session, inputId = "deptSelect", selected = "Choose")
      }
    }
  })#end observe
  
  #selection lists, generate the department list but only create a placeholder for the course list
  output$selectionRow1 <- renderUI({
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
    } else {} #return no ui element if file isnt read yet
  })
  
  #'select all' boxes, generate the department box but only create a placeholder for the course box
  output$selectionRow2 <- renderUI({
    if (generated$read == TRUE) {
      fluidRow(
        column(3, offset = 1, checkboxInput("allBox", label = "Select all Departments and Courses")),
        column(1),
        column(3, uiOutput("courseSelectionAll"))
      )
    } else {} #return no ui element if file isnt read yet
  })
  
  #gen button
  output$selectionRow3 <- renderUI({
     if (generated$read == TRUE) {
       fluidRow(
        column(1, offset = 3, actionButton('genButton', 'Generate CIS Sheets'))
       )
     } else {} #return no ui element if file isnt read yet
  })
  
  #gen button text, uses the div and 'busy' class to run javascript from ui.R to hide/show the 'Working...' text
  output$selectionRow4 <- renderUI({
    if (generated$read == TRUE) {
      fluidRow(
        column(3, offset = 3, textOutput('genSuccess'), 
               div( 
                 class = "busy", 
                 p("Working...")
               ))
      )
    } else {} #return no ui element if file isnt read yet
  })
  
  #course select menu
  output$courseSelection <- renderUI({
    if (generated$read == TRUE) {
      #check existence of the selet all departments box before trying to access its value to avoid errors
      if (!is.null(input$allBox)) {
        #return no UI if all departments are selected
        if (input$allBox == TRUE) {} 
        #if no department was selected, generate an empty course selection list that requests a dept. selection
        else if (input$deptSelect == "Choose") {
           selectInput("courses", "Select a Department first", choices = NULL)
        }
        #otherwise create the course list using the selected department, checking the input's validity first
        else if (length(input$deptSelect) > "") {
          courseCode <- viewData$masterList[[trimWhiteUnd(input$deptSelect)]][,1]
          courseTitle <- viewData$masterList[[trimWhiteUnd(input$deptSelect)]][,2]
          vals <- c()
          index <- c()
          for (i in 1:length(courseCode)) {
            #record the position of the course only if the course code is not null
            if (!is.na(courseCode[i])) {
              #do not record courses with the word "Deleted" in the title either
              if (grepl("Deleted", courseTitle[i]) == FALSE) {
                vals <- c(vals, paste(courseCode[i], courseTitle[i]))
                index <- c(index, i)
              }
            }
          }
          #store the courses for access by the rest of the program, and create the list UI
          viewData$index <- data.frame(vals, index)
          selectInput("courses", "Choose Specific Courses", 
                      choices = vals, multiple = TRUE, selectize = TRUE)
        }
      }
    }
  })
  
  #course 'select all' box
  output$courseSelectionAll <- renderUI({
    if (generated$read == TRUE) {
      if (!is.null(input$allBox) && !is.null(input$deptSelect)) {
        if (input$allBox == TRUE || input$deptSelect == "Choose") {} #return no UI
        else {
          checkboxInput("courseAllBox", "All Courses")
        }
      }
    }
  })
  
})#end shinyServer