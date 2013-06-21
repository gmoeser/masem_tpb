## Load and preprocess data from single studies
## we will automate as far as possible later on - this is work in progress
## Started today (June 17th, 2013)
## Author: Guido Moeser


## Import from Excel-Tables

## First, switch to relevant working directory!
## You should change object working_directory

## Workstation - MG
#working_directory <- "D:/Cloud/Google Drive/02_Science/Publikationen/Morren Bamberg Moeser 2013/masem_tpb"
## Notebook - MG
working_directory <- "D:/Google Drive/02_Science/Publikationen/Morren Bamberg Moeser 2013/GIT_Notebook/masem_tpb"

setwd(working_directory)
getwd()
## See whats in the working-directory
dir()

## We start with the original dataset from Bamberg & Moeser
## File: Daten_Final_Procedure_05_April_2006.xls

## Second, we install necessary libraries to gather data from the files

## We will use xlsx here - but this requires Java - so check if jre7 is installed
# Package information: http://cran.r-project.org/web/packages/xlsx/xlsx.pdf
# Uncomment this if packages are not installed
#install.packages("rJava", dependencies = TRUE)
# If problems with loading package rJava, modify one of the following lines:
#Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre7') # for 64-bit version
#Sys.setenv(JAVA_HOME='C:\\Program Files (x86)\\Java\\jre7') # for 32-bit version
library(rJava)

# Uncomment this if packages are not installed
#install.packages("xlsx", dependencies = TRUE)
library(xlsx)


##### F I R S T    D A T A S E T ##

## Read first sheet in table
sheet_d1 <- read.xlsx("Daten_Final_Procedure_05_April_2006.xls", sheetName = "d1", rowIndex = c(2:11), colIndex = c(2:12))
# rowIndex is not working correct...

## collect all correlations
cor_d1 <- sheet_d1[c(1:10),]
cor_d1

## Change row and column names
colnames(cor_d1) <- c("VAR_", "PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
colnames(cor_d1)
cor_d1

## Change Names in VAR_
cor_d1$VAR_ <- c("PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
cor_d1

## Inspect data.frame
#transform(cor_d1, PROB = as.numeric(PROB))
#str(cor_d1)

## Change factor to factor-levels
## TODO: LOOP
#cor_d1$PROB <- as.numeric(levels(cor_d1$PROB))[cor_d1$PROB]
#cor_d1$AC <- as.numeric(levels(cor_d1$AC))[cor_d1$AC]
#cor_d1$RESP <- as.numeric(levels(cor_d1$RESP))[cor_d1$RESP]
#cor_d1$SNORM <- as.numeric(levels(cor_d1$SNORM))[cor_d1$SNORM]
#cor_d1$GUILT <- as.numeric(levels(cor_d1$GUILT))[cor_d1$GUILT]
#cor_d1$PBC <- as.numeric(levels(cor_d1$PBC))[cor_d1$PBC]
#cor_d1$ATT <- as.numeric(levels(cor_d1$ATT))[cor_d1$ATT]
#cor_d1$MNORM <- as.numeric(levels(cor_d1$MNORM))[cor_d1$MNORM]
#cor_d1$INTENT <- as.numeric(levels(cor_d1$INTENT))[cor_d1$INTENT]
#cor_d1$BEHAV <- as.numeric(levels(cor_d1$BEHAV))[cor_d1$BEHAV]

str(cor_d1)

## Store Var-Names in Vector
cor_d1_names <- names(cor_d1)[-1]

for (i in 1:length(cor_d1_names)) {
  cat("cor_d1$",cor_d1_names[i]," <- as.numeric(levels(cor_d1$",cor_d1_names[i],"))[cor_d1$",cor_d1_names[i],"] ");
}

str(cor_d1)
cor_d1

#cat("cor_d1$",cor_d1_names[1]," <- as.numeric(levels(cor_d1$",cor_d1_names[1],"))[cor_d1$",cor_d1_names[1],"]")


#paste0("cor_d1$",cor_d1_names[1])

## Collect corresponding N
N_d1 <- sheet_d1[c(11:20),]
N_d1

## Change row and column names
colnames(N_d1) <- c("VAR_", "PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
PROBcolnames(N_d1)
N_d1

## Change Names in VAR_
N_d1$VAR_ <- c("PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
N_d1

## Clean up
rm(sheet_d1)





##### S E C O N D    D A T A S E T ##

## Read first sheet in table
sheet_d2 <- read.xlsx("Daten_Final_Procedure_05_April_2006.xls", sheetName = "d2", rowIndex = c(2:11), colIndex = c(2:12))
# rowIndex is not working correct...

## collect all correlations
cor_d2 <- sheet_d2[c(1:10),]
cor_d2

## Change row and column names
colnames(cor_d2) <- c("VAR_", "PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
colnames(cor_d2)
cor_d2

## Change Names in VAR_
cor_d2$VAR_ <- c("PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
cor_d2


## Collect corresponding N
N_d2 <- sheet_d2[c(11:20),]
N_d2

## Change row and column names
colnames(N_d2) <- c("VAR_", "PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
colnames(N_d2)
N_d2

## Change Names in VAR_
N_d2$VAR_ <- c("PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
N_d2

## Clean up
rm(sheet_d2)





##### F I R S T    M E T A - A N A L Y S I S ##
## We will conduct our first meta-analysis now  
## Classical approach - pooling correlation after correlation


## Install and load packages
#install.packages("metafor", dependencies = TRUE)
library(metafor)
# library meta is helpful in drawing Forest-Plots
#install.packages("meta", dependencies = TRUE)
library(meta)


## Extract necessary information for first meta-analysis
## PROB <> AC

## Starting point - correlation matricres and N matrices of single studies
## Build dataframes with metaanalyis relevant data

## First study
PROB_AC_rawdata <- data.frame()
PROB_AC_rawdata[1,1] <- c("d1")
PROB_AC_rawdata[1,2] <- cor_d1[2,2, drop = TRUE]
PROB_AC_rawdata[1,3] <- N_d1[2,2, drop = TRUE]
names(PROB_AC_rawdata) <- c("Study Name", "Correlation - PM", "N" )
PROB_AC_rawdata <- droplevels(PROB_AC_rawdata)
PROB_AC_rawdata

## Second study
#cor_d2 <- droplevels(cor_d2)

PROB_AC_rawdata[2,1] <- c("d2")
PROB_AC_rawdata[2,2] <- cor_d2[2,2, drop = TRUE]
PROB_AC_rawdata[2,3] <- factor(N_d2[2,2])
PROB_AC_rawdata


## NA handling








