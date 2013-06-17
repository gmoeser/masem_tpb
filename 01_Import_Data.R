## Load and preprocess data from single studies

## Import from Excel-Tables

## First, switch to relevant working directory!
## You should change object working_directory
workind_directory <- "D:/Cloud/Google Drive/02_Science/Publikationen/Morren Bamberg Moeser 2013/masem_tpb"
setwd(workind_directory)
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
library(rJava)
# Uncomment this if packages are not installed
#install.packages("xlsx", dependencies = TRUE)
library(xlsx)

## Read first sheet in table
sheet_d1 <- read.xlsx("Daten_Final_Procedure_05_April_2006.xls", sheetName = "d1", rowIndex = c(2:11), colIndex = c(2:12))

# rowIndex is not working probably...






