## Load and preprocess data from single studies
## we will automate as far as possible later on - this is work in progress
## Started today (June 17th, 2013)
## Author: Guido Moeser


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


## Collect corresponding N
N_d1 <- sheet_d1[c(11:20),]
N_d1

## Change row and column names
colnames(N_d1) <- c("VAR_", "PROB", "AC", "RESP", "SNORM", "GUILT", "PBC", "ATT", "MNORM", "INTENT", "BEHAV")
colnames(N_d1)
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
install.packages("meta", dependencies = TRUE)
library(meta)


## Extract necessary information for first meta-analysis
## PROB <> AC








