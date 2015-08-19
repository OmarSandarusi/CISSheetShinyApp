require(excel.link)

#removes all whitespace and replaces with _
trimWhiteUnd <- function (x) gsub(" ", "_", x, fixed = TRUE)#gsub("(^ +)|( +$)", "", x) for appended/prepended whitespaces
#switches forward slash with back slash
replaceFwB <- function (x) gsub("/", "\\", x, fixed = TRUE)
#eliminates all whitespace
trimWhite <- function (x) gsub(" ", "", x, fixed = TRUE)

#find where the departments are located and return pairs of indices
getDeptBounds <- function(df) {
  lastDeptIndex <- 1
  i <- 2
  j <- 1
  c1 <- c()
  c2 <- c()
  while (i <= nrow(df)) {
    #department separator found when course code field is missing twice in a row (separator followed by new course code banner)
    if (is.na(df[i,1]) && is.na(df[i+1,1])) {
      c1[j] <- lastDeptIndex
      c2[j] <- i-1
      lastDeptIndex <- i
      j <- j + 1
    }
    i <- i + 1
  }
  return (data.frame(c1,c2))
} #end getDeptBounds

#load the master list and separate it by department
loadMaster <- function (dir, year) {
  setwd(dir)
  wbName <- paste(getwd(), "/Course Master List - ", year, ".xlsm", sep = '', collapse = '')
  df <- xl.read.file(wbName, header = FALSE, xl.sheet = "Output")
  names(df) <- df[3,] #setting the column headers
  df <- df[-c(1:3),] #removing title rows
  bounds <- getDeptBounds(df)
  numRows <- nrow(bounds)
  dept <- list()
  for (i in 1:numRows) {
    dept[[i]] <- df[bounds$c1[i]:bounds$c2[i],]
    names(dept)[i] <- trimWhiteUnd(dept[[i]][1,2])
    dept[[i]] <- dept[[i]][-1,]
  }
  return (dept)
} #end loadMaster

#generates CIS sheets for every entry of the data frame passed to it
genCISsheets <- function (df, dir, com, year) {
  if (!is.null(df)) {
    setwd(dir)
    #Open CopyFile.xlsm
    path <- "CopyFile.xlsm"
    xl.workbook.open(path)
    xl.workbook.activate(path)
    #Set CopyFile's target directory
    xl.write(paste0(replaceFwB(dir),"\\", "3.1.1_3.1.2_A6C.xlsm"), com[["Activesheet"]]$Cells(2,1))
    #Identify all valid courses in df and pull each index within df as well as each course code
    index <- c()
    name <- c()
    for (i in 1:nrow(df)) {
      #check if course code exists
      if (!is.na(df[i,1])){
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
    #This should begin the VBA code within CopyFile.xlsm
    xl.write(2, com[["Activesheet"]]$Cells(6,1))
    #iterate through df on each copied sheet
    path2 <- "3.1.1_3.1.2_A6C.xlsm"
    xl.workbook.open(path2)
    xl.workbook.activate(path2)
    for (i in 1:length(name)) {
        #CIS sheet editing, some 
        xl.sheet.activate(name[i])
        xl.write(df[index[i],1], com[["Activesheet"]]$Cells(3,3))  #course number
        xl.write(df[index[i],2], com[["Activesheet"]]$Cells(4,3))  #course title
        xl.write(df[index[i],8], com[["Activesheet"]]$Cells(11,5)) #Math
        xl.write(df[index[i],9], com[["Activesheet"]]$Cells(11,7)) #NS
        xl.write(df[index[i],10], com[["Activesheet"]]$Cells(11,9)) #CS
        xl.write(df[index[i],11], com[["Activesheet"]]$Cells(11,11)) #ES
        xl.write(df[index[i],12], com[["Activesheet"]]$Cells(11,13)) #ED
        total <- as.numeric(df[index[i],8]) + as.numeric(df[index[i],9]) + as.numeric(df[index[i],10]) + as.numeric(df[index[i],11]) + as.numeric(df[index[i],12])
        xl.write(total, com[["Activesheet"]]$Cells(11,4)) # Total CEAB units
        total <- 0
        if (!is.na(df[index[i],4])) { total <- total + as.numeric(df[index[i],4]) }
        if (!is.na(df[index[i],5])) { total <- total + as.numeric(df[index[i],5]) }
        if (!is.na(df[index[i],6])) { total <- total + as.numeric(df[index[i],6]) }
        xl.write(total, com[["Activesheet"]]$Cells(26,4)) #Total Instructional Units
        if (is.na(df[index[i],4])) { lec <- 0 } else { lec <- as.numeric(df[index[i], 4]) }
        xl.write(df[index[i],4], com[["Activesheet"]]$Cells(26,6)) #Lecture Time
        if (!is.na(df[index[i],4])) { total <- total - as.numeric(df[index[i], 4]) }
        xl.write(total, com[["Activesheet"]]$Cells(26,7)) #Lab + Tutorial Time
        setwd(dir)
    }#end for (iterates through CIS sheet copies)
    #xl.workbook.save(path)
    #xl.workbook.save(path2)
    #xl.workbook.close(path)
    #xl.workbook.close(path2)
  }#end if(is.null(df))
}

#com <- xl.get.excel() #COMIDispatch object that points to excel
#dir <- "C:/Users/Omar/Documents/Database_Job/CISSheetApp"
#year <- "2015-2016"
#masterList <- loadMaster(dir, year)
#genCISsheets(masterList$Mathematics_and_Engineering[1:4, ], dir, com, year)
