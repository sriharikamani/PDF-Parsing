###########################################################################################################################################
# Program Name  : R1_Extract (Master Program)
# Purpose       : Program to extract contents of R1 category PDF files  
############################################################################################################################################
# TO remove the objects stored in workspace
rm(list=ls(all=T)) 
cat("\014")

library(pdftools)
library(stringr)
library(qdapRegex)
library(sqldf)
library(tcltk)
library(xlsx)
library(readxl)
library(readr)
#
#######################################################################################################################################
# Function to extract the data 
#######################################################################################################################################
# Parameters:
# ----------
# PDF              : The name of the PDF document
# ID_length        : Thelength of the ID to be extracted
# Content_end_case : The "Table Of Contenets" end string Case Type ("/M) 
#######################################################################################################################################
#
extractPdf <- function(PDF,ID_length,Content_end_case){
#
############################################################################################################################################
# Set the Path 
############################################################################################################################################
#
  str_stmt <- paste(("C:/Users/SRIHARI/Documents/My Data/Certification Courses/Rennes/PDF/45 reports/"))
  path     <- paste(str_stmt,PDF,"Converted PDF", sep='/') 
  setwd(path)
#
############################################################################################################################################
# Read the PDF 
############################################################################################################################################
#
  PDF   = paste(PDF,".pdf",sep='')
  text  <- strsplit(pdf_text(PDF), "\r\n")
#
#######################################
# Set the Page boundries to fetch data
#######################################
#
  First_Page  <-"Table of contents"
  Startpage   <- grep(First_Page, trimws(text))
  i           <- Startpage
  
  Last_Page   <-"Closing of the meeting"
  EndpageNo   <- grep(Last_Page, trimws(text))
  
  if(length(EndpageNo) > 1) {
    Endpage        <-EndpageNo[2]
  } else {Endpage  <-EndpageNo[1]}
#
############################################################################################################################################
# Extract the Source and Title  
############################################################################################################################################
#
  source_extrctd_line = data.frame()
  line=1
  
  for(i in 1:Endpage) {
    
    for(j in 1:length(text[[i]][])) {
      source_extrctd_line <- rbind(cbind(paste(line),text[[i]][j]),source_extrctd_line)
      line = line + 1
    }
    source_extrctd_line <- rbind(source_extrctd_line)
  }
  
  colnames(source_extrctd_line)   <- c("RRN","data")
  source_extrctd_line             <- source_extrctd_line[order(source_extrctd_line$RRN,decreasing = TRUE), ]
  
################################
# Extract ALL IDs 
################################

  Line_num   <- grep('^R1-1', source_extrctd_line[[2]][])
  temp              <- data.frame()
  Source_Title_RRN  <- data.frame()
  
  for(i in 1:(length(Line_num))) {
    
    j    = Line_num[i]
    ID   = substr(source_extrctd_line[[2]][j], 1, ID_length)
    temp = source_extrctd_line[[2]][j]
#  
####################################
# Remove words which ends with ")"
####################################
#  
   ends_with <- grep(')$', source_extrctd_line[[2]][j])
    
    if(!identical(ends_with, integer(0))) {
      pos = unlist(gregexpr(pattern = ')',source_extrctd_line[[2]][j]))
      len = length(pos)
      if (len > 1) {
        Source = trimws(substr(source_extrctd_line[[2]][j],1,pos[2]-(ID_length + 2)))
      } else {
        Source = trimws(substr(source_extrctd_line[[2]][j],1,pos[1]-(ID_length + 2)))
      }
    } else {
      Source = source_extrctd_line[[2]][j]
    } 
#  
####################################
# Remove words which ends with "(R1-"
####################################
#  
    pos = 0
    ends_with <- (grepl('\\(R1-$', Source) || grepl('R1-1$', Source) || grepl('R2-1$', Source) || grepl('R4-1$', Source)) 
    Source
    
    if(ends_with == TRUE) {
      if(((pos = (gregexpr("R1-1$", Source))) > 0) || ((pos = (gregexpr("R1-$", Source))) > 0) || ((pos = (gregexpr("\\(R1-$", Source))) > 0) || ((pos = (gregexpr("\\(R4-$", Source))) > 0)) {
        Source = trimws(substr(Source,1,pos[[1]][1]-4))
      }
    }
#  
###################################################
# Remove last word of sentence if ends with number 
###################################################
#  
    pos = 0
    ends_with_num <- grepl("[[:digit:]]$", Source) 
    
    if(ends_with_num == TRUE) {
      pos <- gregexpr("[0-9]$+", Source)
      Source = trimws(substr(Source,1,pos[[1]][1]-(ID_length + 2)))
    } 
    
    # Replace the character "(" with " "  
    
    Source = str_replace(Source, "\\(", "")     
#
###################################################
# Identify the number of spaces to break the words 
# Steps:
# -----
# 1. Get the count of white spaces between each word and store in a vector
# 2. Get the length of total words
# 3. Skip the first space since pattern of spaces are high and also belongs to ID
# 4. Select maximum space count for rest of the words as pattern of space count is high for the Source
# 5. Pick the last word Space count and pad a String vector with Whitespaces based on the maximum space count  
# 6. Now break the sentence by using string split on the String vector
# 7. Capture the last word as SOURCE
###################################################
# 
    countWhiteSpaces <- function(x) attr(gregexpr("(?<=[^ ])[ ]+(?=[^ ])", Source, perl = TRUE)[[1]], "match.length")
    space_vector     <- countWhiteSpaces(Source)
    
    if (space_vector != -1) {
      
      if (space_vector >1) {
        len              <- length(space_vector)
        space_vector     <- space_vector[2:len]
      }
      max_space_count  <- max(space_vector)
      spaces           <- data.frame()
      spaces           <- str_pad("", width = max_space_count)
      Source           <- str_trim((str_split(Source, pattern = spaces)[[1]]))
      WordCount        <- length(Source)
      Source           <- Source[WordCount]
      Source           <- str_replace(Source, "\\(", "")
      }
#
#################
# Title 
#################
#
      endpos = 0
      from = ID_length+1
      endpos            <- gregexpr(Source, temp)
      Title             <- trimws(substr(temp,from,endpos[[1]][1]-1))
      Source_Title_RRN  <- rbind(cbind(i,ID,cbind(gsub("\uf0b7","",temp)), Source,Title),Source_Title_RRN)
    }
  #}
  
  colnames(Source_Title_RRN)  <- c("RRN","ID","Text","Source","Title")
  Source_Title_RRN            <- Source_Title_RRN[order(Source_Title_RRN$RRN, decreasing = TRUE), ]
  write.xlsx(Source_Title_RRN, file = "Source.xlsx")
  rm(Source_Title_RRN)
  rm(source_extrctd_line)
#
#######################################################################################################################################
############################# E N D    O F    T H E    S O U R C E    T I T L E    E X T R A C T I O N ################################
############################################################################################################################################
# Logic to Extract Table of Contents Information
############################################################################################################################################
#
  Contents_colums = data.frame(stringsAsFactors=FALSE)
#
######################################################################
# Set the Page boundries to fetch get "Table of contents" information
######################################################################
#
  First_Page  <-"Table of contents"
  Startpage   <- grep(First_Page, trimws(text))
  i           <- Startpage
  
  if(Content_end_case == "M") {
    
    Last_Page   <-"Closing of the meeting"
    EndpageNo   <- grep(Last_Page, trimws(text))
    Endpage     <-EndpageNo[1]
  } else {
    Last_Page   <-toupper("CLOSING OF THE MEETING")
    EndpageNo   <- grep(Last_Page, trimws(text))
  }
  
  Endpage     <-EndpageNo[1]
  j           <-Endpage
  
  for(i in 1:EndpageNo) {
    for(j in 1:length(text[[i]][])) {
      text[[i]][j] = trimws(text[[i]][j])   
      text[[i]][j] = trimws(gsub("[0-9^]+$", "", text[[i]][j])) # remove the page number
      text[[i]][j] <- gsub("\\s+"," ",text[[i]][j])  # Remove > 1 space between the words
      
      spl = "N"
      
      if(substring(text[[i]][j],1,1) == "\uf0b7") {
        spl = "Y"
        print("blnk")
        text[[i]][j] = trimws(str_replace(text[[i]][j], " ", ""))   
      }

      ###################
      # Split the text
      ###################
      
      Number       <- trimws(stringr::word(text[[i]][j], 1))
      Number  
      have_digit <- grepl("[[:digit:]]", Number)
      
      if((have_digit == TRUE) & (spl == "Y")) {
        
        Number <- parse_number(Number)
        #Number <- parse_number(Number,locale=locale(grouping_mark=".", decimal_mark=","))
        Number <- as.character(Number)
      }
      
      if(spl == "N") {
        Text         <- trimws(substr(text[[i]][j], nchar(Number)+1,200))
      } else {
        Text         <- trimws(substr(text[[i]][j], nchar(Number)+2,200))
      }
      Text         <- trimws(gsub(".", "", Text,fixed = TRUE))
      data         <- paste(Number,Text)
      data         <- trimws(substr(data, 1, 15))
      
      ###########################
      # Identify the Header type 
      ###########################
      
      numOcc_OptA = 0
      numOcc_OptB = 0 
      numOcc_OptA = nchar(Number) - nchar(gsub(".", "", Number,fixed = TRUE)) # fixed=TRUE instead of escaping the character "."
      
      pos      <- which(strsplit(Number, "")[[1]]==".")
      
      if(!identical(pos, integer(0))) {
        pos         <- pos[[1]][1]
        after_.     <- trimws(substr(Number, pos+1,10))
        numOcc_OptB <- nchar(gsub(".", "", after_.,fixed = TRUE))
      }
      
      if(is.na(numOcc_OptA)) {
        break
      }
      
      if(numOcc_OptA < numOcc_OptB) {
        numOcc = numOcc_OptA
      } else {
        numOcc = numOcc_OptB
      }
      
      if(numOcc == 0) {Type = 'HEAD'}
      if(numOcc == 1) {Type = 'SUBHEAD1'}
      if(numOcc == 2) {Type = 'SUBHEAD2'}
      if(numOcc == 3) {Type = 'SUBHEAD3'}
      if(numOcc == 4) {Type = 'SUBHEAD4'}      
      
      Contain_Char <-grepl("[A-Za-z]", Number)
      
      if(Contain_Char == TRUE) {
        Number = ""
      }
      
      if(Text == "") {        
        Number = ""
        Text = ""
        data = ""
        Type = ""
      }
      Contents_colums = rbind(Contents_colums,cbind(gsub("\uf0b7","",Number),gsub("\uf0b7","",Text),gsub("\uf0b7","",tolower(data)),Type))
      
    }
  }
  colnames(Contents_colums)  <- c("Number","Text","data","Type")
  
  # Delete all rows whose TYPE & Number = "" 
  
  Contents_colums <- subset(Contents_colums, Contents_colums$data!="")
  Contents_colums <- subset(Contents_colums, Contents_colums$Number!="")
  
############################################################################################### 
# Populate Table Contents and remove the rows whose column values of  TYPE & Number are Blanks
###############################################################################################
#
  write.xlsx(Contents_colums, file = "Table_Of_Contents.xlsx")
  rm(Contents_colums)
  rm(Contain_Char)
  rm(data)
  rm(Number) 
  rm(Text)
  rm(numOcc)
  rm(after_.)
#
############################################################################################################################################
# End of Table of Content Information Extraction Logic
############################################################################################################################################
#
############################################################################################################################################
# Logic to Extract features from given PDF file
#############################################################################################################################################
####################
# Load the PDF File 
####################
#
  Table_Of_Contents  <- read_excel("Table_Of_Contents.xlsx")
  Table_Of_Contents  <- subset(Table_Of_Contents,select=-c(X__1))
#
#######################
# Remove white spaces 
#####################@#
#
  extrctd_line <- extrctd_line_Full <- data.frame()
  line         <- i <- j <- 1
  text_len     <- length(text)
  i            <- Endpage + 1
  Page_Count   <- (length(text) - Endpage) + 1
  
  while(i<=Page_Count) {
    
    while(j <=length(text[[i]][])) {
      
      if(is.na(text[[i]][])) {
        print("NA HERE")
      }
      
      ###################################################
      # Stop reading the text on encountering Annex = A
      ###################################################
      
      End = grep("^Annex A:", trimws(substr(text[[i]][j],1,14))) # 
      
      if(!identical(End, integer(0))) {
        
        print("Annex A:")
        i <- text_len
        break
        
      } else {
        text[[i]][j]      <- trimws(gsub("\\s+"," ",text[[i]][j]))  # Remove > 1 space between the words
        extrctd_line_Full <- rbind(cbind(paste(line),trimws(text[[i]][j])),extrctd_line_Full)
        try_error         <- tryCatch(tolower(text[[i]][j]), error = function(e) e)
        
        ########################################################################################################
        # Select all lines which are converted to lower case with tolower() function. Skip which have UTC format        
        # Converting only headers & sub-headers to lower case.
        ########################################################################################################
        
        if (!inherits(try_error, "error")) {
          
          start = grep("^[0-9]", text[[i]][j])
          
          if(!identical(start, integer(0))) {
            extrctd_line <- rbind(cbind(paste(line),trimws(substr(trimws(tolower(text[[i]][j])),1, 15))),extrctd_line)
          } else {
            extrctd_line <- rbind(cbind(paste(line),trimws(substr(trimws(text[[i]][j]),1, 15))),extrctd_line)
          }
        }
        line = line + 1
      }
      
      j = j+ 1
    }
    
    extrctd_line_Full <- rbind(extrctd_line_Full)
    extrctd_line      <- rbind(extrctd_line)
    
    i = i + 1
    j = 1
  }
  
  colnames(extrctd_line)        <- c("RRN","data")
  colnames(extrctd_line_Full)   <- c("RRN","data")
  extrctd_line                  <- extrctd_line[order(extrctd_line$RRN,decreasing = TRUE), ]
  extrctd_line_Full             <- extrctd_line_Full[order(extrctd_line_Full$RRN,decreasing = TRUE), ]
  #
  write.csv(extrctd_line,file = "extrctd_line.csv")
  extrctd_line   <- read.csv("extrctd_line.csv",header = T)
  extrctd_line   <- subset(extrctd_line,select=-c(X))  
#
############################################################################################################################################
# Delete lines other than headings,sub-headings,IDs,Decision and Discussion and sort the extracted data based on RRN
############################################################################################################################################
#
  Term     <- c('Decision:', 'Discussion:')
#
###################
# Extract Heading 
###################
#
  Head     <- merge(x=extrctd_line,y=Table_Of_Contents,by = "data")
  qry_stmt <- paste("SELECT Number,RRN, Text,Type from Head where Type = 'HEAD' ORDER BY Number ASC",sep="")
  Head     <- sqldf(qry_stmt)
  colnames(Head) <- c("Index","RRN","data","Type")
  Head$data <- paste(Head$Index,Head$data, sep=' ') 
  Head <- subset(Head,select=-c(Index))
#
######################
# Extract SubHeading 1  
######################
#
  SubHead1  <- merge(x=extrctd_line,y=Table_Of_Contents,by = "data")
  qry_stmt  <- paste("SELECT Number,RRN, Text,Type from SubHead1 where Type = 'SUBHEAD1' ORDER BY RRN ASC",sep="")
  SubHead1  <- sqldf(qry_stmt)
  colnames(SubHead1) <- c("Index","RRN","data","Type")
  SubHead1$data      <- paste(SubHead1$Index,SubHead1$data, sep=' ') 
  SubHead1           <- subset(SubHead1,select=-c(Index))
#
######################
# Extract SubHeading 2  
######################
#
  SubHead2  <- merge(x=extrctd_line,y=Table_Of_Contents,by = "data")
  qry_stmt  <- paste("SELECT Number,RRN, Text,Type from SubHead2 where Type = 'SUBHEAD2' ORDER BY RRN ASC",sep="")
  SubHead2  <- sqldf(qry_stmt)
  colnames(SubHead2) <- c("Index","RRN","data","Type")
  SubHead2$data      <- paste(SubHead2$Index,SubHead2$data, sep=' ') 
  SubHead2           <- subset(SubHead2,select=-c(Index))
#
######################
# Extract SubHeading 3  
######################
#
  SubHead3  <- merge(x=extrctd_line,y=Table_Of_Contents,by = "data")
  qry_stmt  <- paste("SELECT Number,RRN, Text,Type from SubHead3 where Type = 'SUBHEAD3' ORDER BY RRN ASC",sep="")
  SubHead3  <- sqldf(qry_stmt)
  colnames(SubHead3) <- c("Index","RRN","data","Type")
  SubHead3$data      <- paste(SubHead3$Index,SubHead3$data, sep=' ') 
  SubHead3           <- subset(SubHead3,select=-c(Index))
#
######################
# Extract SubHeading 4
######################
#
  SubHead4  <- merge(x=extrctd_line,y=Table_Of_Contents,by = "data")
  qry_stmt  <- paste("SELECT Number,RRN, Text,Type from SubHead4 where Type = 'SUBHEAD4' ORDER BY RRN ASC",sep="")
  SubHead4  <- sqldf(qry_stmt)
  colnames(SubHead4) <- c("Index","RRN","data","Type")
  SubHead4$data      <- paste(SubHead4$Index,SubHead4$data, sep=' ') 
  SubHead4           <- subset(SubHead4,select=-c(Index))
#
############# 
# Extract ID  
#############
# 
  Line_num    <- grep('^R1-1', extrctd_line_Full[[2]][])
  ID_RRN      <- data.frame()  

  for(i in 1:(length(Line_num))) {
    j      = Line_num[i]
    ID     = substr(extrctd_line_Full[[2]][j], 1, ID_length)
    ID_RRN     = rbind(cbind(Line_num[i],ID, "ID"),ID_RRN)
  }
  
  colnames(ID_RRN)     <- c("RRN","data","Type")
  ID_RRN               <- ID_RRN[order(ID_RRN$RRN, decreasing = TRUE), ]
#
################################
# Logic to extract Decision
################################
#
  qry_stmt            <- paste("SELECT * from extrctd_line_Full WHERE data LIKE '",Term[1],"%'", sep="")  
  Decision            <- sqldf(qry_stmt)
  Decision[c("Type")] <- "Decision"
  Decision[[2]]       <- str_replace(Decision[[2]], "Decision:", "")
#
################################
# Logic to extract Discussion
################################
#
  qry_stmt              <- paste("SELECT * from extrctd_line_Full WHERE data LIKE '",Term[2],"%'", sep="")  
  Discussion            <- sqldf(qry_stmt)
  Discussion[c("Type")] <- "Discussion"
  Discussion[[2]]       <- str_replace(Discussion[[2]], "Discussion:", "")
  #
  extrctd_lines         <- rbind(ID_RRN,Discussion,Decision)
  extrctd_lines         <- extrctd_lines[order(extrctd_lines$RRN,decreasing = TRUE), ]
  write.csv(extrctd_lines, file = "extrctd_lines.csv")  
#
############################################################################################################################################
#             ************** End of Extraction Logic ***************
############################################################################################################################################
#
############################
# Remove the unused objects 
############################
#
  rm(extrctd_line)
  rm(extrctd_lines)
  rm(text)
  rm(ID_RRN)
  rm(Decision)
  rm(qry_stmt)
# 
############################################################################################################################################
# Logic to Delete ID's which have no corresponding Decision  
############################################################################################################################################
#
  data <- read.csv(file="extrctd_lines.csv",header=TRUE, stringsAsFactors=F)
  data <-subset(data,select=-c(X))
  data <- data[order(data$RRN, decreasing = FALSE), ]
  Temp <- data.frame()
  Temp <- data
  
# Get the count of Number ID's V/s Decisions
  
  qry_stmt            <- paste("SELECT count(*) from data WHERE Type = 'ID'", sep="")  
  ID_Count            <- sqldf(qry_stmt)
  qry_stmt            <- paste("SELECT count(*) from data WHERE Type = 'Decision'", sep="")  
  Decision_Count      <- sqldf(qry_stmt)
  
# Check for count match. If not delete additional ID's which have no corresponding Decision based on RRN
  
  if((ID_Count !=  Decision_Count)) {
    
    i = 1
    Line_count <- nrow(data)
    
    while (i <= Line_count) {
      
      if(is.na(Temp[[3]][i]) || is.na(Temp[[3]][i+1])) { 
        
        if(Temp[[3]][i] == 'ID') {
          Temp[[3]][i]  = ""
        }
        
        break
      }
      
      Flag = 0
      j    = i+1     
      
      if((Temp[[3]][i] == 'ID') && (Temp[[3]][j] == 'ID')) {
        
        if(is.na(Temp[[3]][i])) { 
          break
        }
        
        Flag = 1
      }
      
      if(Flag == 1) {
        Temp[[3]][i] = ""
      }
      
      i = i+1
    }
  }
  
  
# Delete all rows whose TYPE = "" and split the ID and Decision
  
  Temp <- subset(Temp, Temp$Type!="")
  
#####################################################
# Now bind the Head, Subhead1 & Subhead2 & Subhead3
#####################################################
#
  extrctd_lines = rbind(Head,SubHead1,SubHead2,SubHead3,SubHead4,Temp)   
  extrctd_lines   <- extrctd_lines[order(extrctd_lines$RRN,decreasing = TRUE), ]
  write.csv(extrctd_lines, file = "extrctd_lines.csv")
#
############################################################################################################################################
# Logic to capture additional lines (IF ANY) for Discussion Information  
############################################################################################################################################
#
  rm(data)
  data <- extrctd_lines
  data <- read.csv(file="extrctd_lines.csv",header=TRUE, stringsAsFactors=F)
  data <-subset(data,select=-c(X))
  data         <- data[order(data$RRN, decreasing = FALSE), ]
  Additional_Discussion <- data.frame()
  
  i = 1
  Line_count <- nrow(data)
  
  while (i <= Line_count) {
    
    if(data[[3]][i] == 'Discussion') {
      
      Current_RRN = data[[1]][i]
      Next_RRN = Current_RRN + 1
      j= i+1
      
      # Exit loop for last record
      
      if(is.na(data[[1]][j])) { 
        print('break')
        break
      }
      
      if(data[[1]][j] != Next_RRN) {
        
        temp = trimws(data[[2]][i])
        l = i
        
        while (Next_RRN < data[[1]][i+1]) {
          
          x = extrctd_line_Full[[2]][Next_RRN]
          Additional_Discussion <- paste(Additional_Discussion,x)
          Next_RRN = Next_RRN + 1  
          
        }
        data[[2]][l] = paste(extrctd_line_Full[[2]][Current_RRN],Additional_Discussion)
        Additional_Discussion = ""
      }
      
    }
    i = i + 1 
  } 
  
  data[[2]] <- str_replace(data[[2]], "Discussion:", "")
  data[[2]] <- trimws(data[[2]])
  write.csv(data, file = "extrctd_lines.csv")
  rm(Additional_Discussion)
  #
  data  <- read.csv(file="extrctd_lines.csv",header=TRUE, stringsAsFactors=F)
  data  <- subset(data,select=-c(X))
  data  <- data[order(data$RRN, decreasing = FALSE), ]
 #
############################################################################################################################################
# Variables and Initializations
############################################################################################################################################
#
  Doc_colums = data.frame(stringsAsFactors=FALSE)
  i = 1 
  Line_count <- nrow(data)
  ID <- Source_status <- Decision_status <- Discussion_status <- ""
#
############################################################################################################################################
# Logic to get the feature text
############################################################################################################################################
#  
  while (i <= Line_count) {
    
    if((i==1) & (data[[3]][i] != 'HEAD')) {
      Head = "HEADER MISSING IN PDF"
      Sub_Head1 <- Sub_Head2 <- Sub_Head3 <- Sub_Head4  <- " "
    }
  
    if(data[[3]][i] == 'HEAD') {
      Head     = data[[2]][i]
      Sub_Head1 <- Sub_Head2 <- Sub_Head3 <- Sub_Head4  <- " "
      i = i + 1
    }  
    
    if(is.na(data[[3]][i])) { 
      break
    }
    
    if(data[[3]][i] == 'SUBHEAD1') {
      Sub_Head1  = data[[2]][i]
      Sub_Head2 <- Sub_Head3 <- Sub_Head4  <- " "    
      i = i + 1
    }  
    
    if(data[[3]][i] == 'SUBHEAD2') {
      Sub_Head2  = data[[2]][i]
      Sub_Head3 <- Sub_Head4  <- " "   
      i = i + 1
    }
    
    if(data[[3]][i] == 'SUBHEAD3') {
      Sub_Head3 = data[[2]][i]
      Sub_Head4    = " "    
      i = i + 1
    }
    
    if(data[[3]][i] == 'SUBHEAD4') {
      Sub_Head3 = data[[2]][i]
      i = i + 1
    }
    
    if(data[[3]][i] == 'ID') {
      
      ID  <- data[[2]][i]
      i   <- i + 1
    }
    
    if(is.na(data[[3]][i])) { 
      Doc_colums <- rbind(Doc_colums,cbind(i,Head,Sub_Head1,Sub_Head2,Sub_Head3,ID,Discussion_status,Decision_status))
      ID <- Source_status <- Decision_status <- Discussion_status <- ""
      break
    }
    
    if(data[[3]][i] == 'Discussion') {
      
      Discussion_status   <- trimws(data[[2]][i])
      i = i + 1
      Decision_status   <- trimws(data[[2]][i])
      i = i + 1
    }
    
    if(is.na(data[[3]][i])) { 
      Doc_colums <- rbind(Doc_colums,cbind(i,Head,Sub_Head1,Sub_Head2,Sub_Head3,Sub_Head4,ID,Discussion_status,Decision_status))
      ID <- Source_status <- Decision_status <- Discussion_status <- ""
      break
    }
    
    if(data[[3]][i] == 'Decision') {
      
      Discussion_status   <- "Not Available"
      Decision_status   <- trimws(data[[2]][i])
      i = i + 1
    }
  
    Doc_colums <- rbind(Doc_colums,cbind(i,Head,Sub_Head1,Sub_Head2,Sub_Head3,Sub_Head4,ID,Discussion_status,Decision_status))
    ID <- Source_status <- Decision_status <- Discussion_status <- ""
  } # End of While Loop 
  
  if (i > Line_count) {
    
    Doc_colums <- rbind(Doc_colums,cbind(i,Head,Sub_Head1,Sub_Head2,Sub_Head3,Sub_Head4,ID,Discussion_status,Decision_status))
    ID <- Source_status <- Decision_status <- Discussion_status <- ""
  }
#
###########################################################################################################################################
# Write to Outfile
############################################################################################################################################
#
  Doc_colums = subset(Doc_colums,select=-c(i))
  colnames(Doc_colums)   = c("Head","Subhead1","Subhead2","Subhead3","Subhead4", "ID","Discussion","Decision" )
  write.xlsx(Doc_colums, file = "Outfile.xlsx")
  rm(Doc_colums)
#
############################################################################################################## 
# Load data from populated file and merge Source & Title information for each ID
##############################################################################################################
  #
  data       <- read_excel("Outfile.xlsx")
  sourcedata <- read_excel("Source.xlsx")
  data       <- subset(data,select=-c(X__1))
  sourcedata <- subset(sourcedata,select=-c(X__1))
  #
  ###########################
  # Extract Source and Title 
  ###########################
  #
  data1    <- merge(x=data,y=sourcedata,by = "ID") 
  qry_stmt <- paste("SELECT Head,Subhead1,Subhead2,Subhead3,Subhead4,ID,Title,Source, Discussion, Decision from data1 ORDER BY ID ASC",sep="")
  data     <- sqldf(qry_stmt)
  colnames(data)   = c("Head","Subhead1","Subhead2","Sub_Head3","Sub_Head4","ID","Title", "Source", "Discussion","Decision" )
  data[is.na(data)] <- " "
  Outfile <- paste(PDF,"_Final_Extract.xlsx",sep="")
  write.xlsx(data, file = Outfile)
  
  ###########################
  # Delete objects 
  ###########################
  rm(Outfile)
  rm(Source)
  rm(extrctd_lines)
  rm(extrctd_line)
  rm(list=ls(all=T)) 
  cat("\014")
  return()
}
#######################################################################################################################################
################################ E N D   O F   T H E   E X T R A C T  F U N C T I O N #################################################
#######################################################################################################################################
#
#########################################################
#   Main Program - Call function to extarct the PDF data
#########################################################

extractPdf("R1-79",9,"M")
extractPdf("R1-80",9,"M")
extractPdf("R1-80b",9,"M")
extractPdf("R1-81",9,"M")
extractPdf("R1-82",9,"M")
extractPdf("R1-82b",9,"M")
extractPdf("R1-83",9,"M")
extractPdf("R1-84",9,"M")
extractPdf("R1-84b",9,"M") 
extractPdf("R1-85",9,"M")
extractPdf("R1-86",9,"M")
extractPdf("R1-86b",10,"M")
extractPdf("R1-87",10,"M")
extractPdf("R1-88",10,"M")  
extractPdf("R1-88b",10,"") 
extractPdf("R1-89",10,"")
extractPdf("R1-90",10,"")
extractPdf("R1-90b",10,"") 
extractPdf("R1-91",10,"")
extractPdf("R1-92",10,"")
extractPdf("R1-92b",10,"")
extractPdf("R1-93",10,"")
extractPdf("R1-94",10,"")
extractPdf("R1-94b",10,"")
rm(list=ls(all=T))
############################# E N D   O F   T H E   P R O G R A M ##########################################################################
  