# install.packages("rtweet")
# install.packages("tidytext")
# install.packages("stringr")
# install.packages("magrittr")
# install.packages("imager")
# install.packages("openxlsx")
# install.packages("devtools")
# install_github("omegahat/RDCOMClient")
# install.packages("jjb")

## --------------------- set work directory
setwd("D:/MyProjects/Sources/12. SNS Develop/GitHub/RCallTwitter")
getwd()

# ------------------------------------------------------- #
source("./sources/TweetFunction.R")
source("./sources/Export2Excel.R")
source("./sources/CallMacroExcel.R")
source("./sources/FileAndFolder.R")

## --------------------- Run program ---------------------
addixcs_token = "ADDIXCS-Premium.rds"
func_getTokenAccount(addixcs_token)
# func_setPremiumAccount()
##Check current token
get_token()

## --------------------- Main program ---------------------
func_mainProgram <- function(excelFileName, pptFileName, username, numberRecord, sheetName){
  
  ## Copy powerpoint file from input folder to output folder
  inPowerPointFile <- paste0(getwd(),"/input/", pptFileName)
  outFolder <- paste0(getwd(),"/output")
  print(inPowerPointFile)
  print(outFolder)
  func_moveFile2Folder(inPowerPointFile, outFolder)
  
  func_write2ExcelTweet(excelFileName, func_getUserTweet_30Days_premium(username, numberRecord), sheetName)
}

  # ## Call macro from excel file
  # myFileName =  paste0(getwd(),"/output/",excelFileName)
  # macroName = "Main"
  # func_RCallExcelMacro(myFileName, macroName)
  # message("Finish output to powerpoint file")

func_move2Folder <- function(username, excelFileName, pptFileName){
  ## create folder donald trump
  myNewFolder <- paste0(getwd(),"/output/",username)
  func_createFolder(myNewFolder)
  ## Copy 2 file and paste to folder donald trump
  outExcelFile <- paste0(getwd(),"/output/",excelFileName)
  outPowerPointFile <- paste0(getwd(),"/output/", pptFileName)
  func_moveFile2Folder(outExcelFile, myNewFolder)
  func_moveFile2Folder(outPowerPointFile, myNewFolder)
}

## --------------------- Test Output Excel Case ---------------------
username = "realDonaldTrump"
sheetName = "Twitter" 
excelFileName = "Twitter-Excel.xlsm"
pptFileName = "Twitter-PPT.pptx"
numberRecord = 50

## Call macro from excel file
myFileName =  paste0(getwd(),"/output/", "Book1.xlsm")
macroName = "hello"
func_RCallExcelMacro(myFileName, macroName)
message("Finish output to powerpoint file")

# get data from Twitter -> output to excel file
func_mainProgram(excelFileName, pptFileName, username, numberRecord, sheetName)

# Call bat file to write from excel file to Power point file, then waiting them finish
Sys.sleep()

# move to new folder
func_move2Folder(username, excelFileName, pptFileName)

