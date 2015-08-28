
#
#Author: Omar Sandarusi
#Queen's University
#September 2015
#
#Support file for CISSheetShinyApp
#

require(excel.link)

#removes all whitespace from a string and replaces with an underscore
trimWhiteUnd <- function (x) gsub(" ", "_", x, fixed = TRUE)
#takes a string and switches forward slashes with back slashes
replaceFwB <- function (x) gsub("/", "\\", x, fixed = TRUE)
#eliminates all whitespace in a string
trimWhite <- function (x) gsub(" ", "", x, fixed = TRUE)

#controls xl.write commands that may incur an error by sending an NA value to an excel cell with range limits
#if an NA value is being written, it writes the 'err' parameter instead. Otherwise it writes the 'val' parameter
writeNA <- function(val, pos, err) {
  if (is.na(val)) {
    xl.write(err, pos)
  } else {
    xl.write(val, pos)
  }
}

#find where the departments are located and return a data.frame with the pairs of indices for the start and end of 
#each department's range, along with its name
getDeptBounds <- function(df) {
  lastDeptIndex <- 4 #holds the position of the first department initially, used to grab the department names
  buffer <- 0 #creates an index gap for when more than two headings occur in a row
  i <- 6
  j <- 1
  start <- c(5) #we know the first range begins after the first department name 
  end <- c()
  deptName <- c()
  while (i <= nrow(df)) {
    #department separator found when course code field is missing more than once
    if (is.na(df[i,1]) && is.na(df[i+1,1])) {
      #set the end of the previous department's range, as well as its name
      end[j] <- i-1
      deptName[j] <- df[lastDeptIndex,2]
      lastDeptIndex <- i
      j <- j + 1
      #reset the buffer and increment it until the headers are passed and an actual course code is found
      buffer <- 1
      while (is.na(df[i+buffer,1]) && is.na(df[i+buffer+1,1]) && buffer < 10) {
        buffer <- buffer + 1
      }
      #case where the end is reached is indicated by a buffer of 10, set it back to 1 for following arithmetic
      if (buffer >= 10) {
        buffer = 1
      }
      #set the start of the next department's range by adding the buffer, which at minimum is 1
      start[j] <- i + buffer
      buffer <- buffer - 1
    } #the buffer is decremented so that normally i+1+buffer = i+1
      #  ^this is while it is incrementing through the actual list of courses
    i <- i + 1 + buffer
  }
  #set the final end and deptName because the while loop has reached the end of the 
  #course list but never set the final values
  end[j] <- i-1
  deptName[j] <- df[lastDeptIndex,2]
  #erase an extra entry that occurs when the last row is a course type header
  if(end[j] > nrow(df) || start[j] > nrow(df)) {
    end <- end[-j]
    start <- start[-j]
    deptName <- deptName[-j]
  }
  return (data.frame(start,end,deptName))
} #end getDeptBounds

#load the master list and separate it by department
#returns a named list made of data.frames of each department's course list, with headers still intact
loadMaster <- function (dir, year) {
  setwd(dir)
  wbName <- paste(getwd(), "/Course Master List - ", year, ".xlsm", sep = '', collapse = '')
  df <- xl.read.file(wbName, header = FALSE, xl.sheet = "Output")
  names(df) <- df[3,] #setting the column headers
  bounds <- getDeptBounds(df)
  numRows <- nrow(bounds)
  dept <- list()
  for (i in 1:numRows) {
    dept[[i]] <- df[bounds$start[i]:bounds$end[i],] #each entry in dept is a department's range of courses (headers included)
    names(dept)[i] <- trimWhiteUnd(bounds$deptName[i]) #set the name of the department, switching whitespace to underscores
    dept[[i]] <- dept[[i]][-1,] #removes first header, since it is known and unnecessary
  }
  return (dept)
} #end loadMaster

#generates CIS sheets for every entry of the data frame passed to it.
#designed to be called separately for each department in the list returned by loadMaster().
#writes data to CopyFile.xlsm, which forces it to run VBA code that duplicates 3.1.1_3.1.2_A6C.xlsm and
#creates CIS sheets named after every course in the department (course names pulled from df). 
#The generated excel file is named after the specified department and then populated with course data from df.
genCISsheets <- function (df, dir, com, year, path2) {
  if (!is.null(df)) {
    setwd(dir)
    #Open CopyFile.xlsm
    path <- "CopyFile.xlsm"
    xl.workbook.open(path)
    xl.workbook.activate(path)
    #Set CopyFile's target directory
    #xl.write(paste0(replaceFwB(dir),"\\", "3.1.1_3.1.2_A6C.xlsm"), com[["Activesheet"]]$Cells(2,1))
    xl.write(paste0(replaceFwB(dir), "\\", path2), com[["Activesheet"]]$Cells(2,1))
    #Identify all valid courses in df
    index <- c()
    name <- c()
    for (i in 1:nrow(df)) {
      #if course code exists and it is not deleted, record the index and course code
      if (!is.na(df[i,1]) && !grepl("Deleted", df[i,2])){ 
         index <- c(index, i)
         name <- c(name, trimWhite(df[i,1]))
      }
    }
    #Write the course codes to CopyFile.xlsm so it can properly name the CIS sheet copies
    for (i in 1:length(name)) {
      xl.write(name[i], com[["Activesheet"]]$Cells(2,i+3))
    }
    #Fill out the iterations field.
    xl.write(length(index), com[["Activesheet"]]$Cells(2,2))
    #This should begin the VBA code within CopyFile.xlsm. Note that it is never closed by R, it closes itself in the VBA code
    xl.write(2, com[["Activesheet"]]$Cells(6,1))
    #iterate through df on each copied sheet
    xl.workbook.open(path2) 
    #browser()
    xl.workbook.activate(path2)
    #Edit each CIS sheet
    for (i in 1:length(name)) {
        xl.sheet.activate(name[i])
        xl.write(df[index[i],1], com[["Activesheet"]]$Cells(3,3))  #course number
        xl.write(df[index[i],2], com[["Activesheet"]]$Cells(4,3))  #course title
        writeNA(df[index[i],8], com[["Activesheet"]]$Cells(11,5), 0) #Math
        writeNA(df[index[i],9], com[["Activesheet"]]$Cells(11,7), 0) #NS
        writeNA(df[index[i],10], com[["Activesheet"]]$Cells(11,9), 0) #CS
        writeNA(df[index[i],11], com[["Activesheet"]]$Cells(11,11), 0) #ES
        writeNA(df[index[i],12], com[["Activesheet"]]$Cells(11,13), 0) #ED
        #checking total times
        lec <- as.numeric(df[index[i],4])
        lab <- as.numeric(df[index[i],5])
        tut <- as.numeric(df[index[i],6])
        writeNA(lec+lab+tut, com[["Activesheet"]]$Cells(26,4), "NA") #Total Instructional Units
        writeNA(lec, com[["Activesheet"]]$Cells(26,6), "NA") #Lecture Time
        writeNA(lab+tut, com[["Activesheet"]]$Cells(26,7), "NA") #Lab + Tutorial Time
        setwd(dir)
    }#end for (iterated through CIS sheet copies)
    xl.workbook.save(path2)
    xl.workbook.close(path2)
  }#end if(is.null(df))
} #end genCISsheets

#to run this script on its own, just hange dir to your active directory and year to the appropriate value of 
#the Master List file you are trying to read, and uncomment all the following lines
#com <- xl.get.excel() #COMIDispatch object that points to excel
#dir <- "C:/Users/Omar/Documents/Database_Job/CISSheetApp"
#year <- "2015-2016"
#masterList <- loadMaster(dir, year)
#genCISsheets(masterList$Mathematics_and_Engineering[1:4, ], dir, com, year)
