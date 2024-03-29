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
setwd("C:/SNS-TWITTER")
getwd()
# ------------------------------------------------------- #
source("./sources/TweetFunction.R")
source("./sources/Export2Excel.R")
source("./sources/FileAndFolder.R")
source("./sources/CallMacroExcel.R")

## --------------------- Run program ---------------------

# addixcs_token = "ADDIXCS-Premium.rds"
# func_getTokenAccount(addixcs_token)
# func_setPremiumAccount()
##Check current token
appname <- "ADDIXCS-Twitter"
consumer_key <- "3a1N6xOBowUuF3ZWlErboyGLo"
consumer_secret <-"sZLiof7XdT8BY3HixQWpp2TQm737VKJcvdxQDovVZPH9Fmw9SB"
access_token <- "1181201713602027522-bpCt556WTeMhsYCwGLLlATHJbhdpvI"
access_secret <- "UljO4r6L9T22yxKIrDfVB35lNr7wJCXuAu49Noe1b3n5s"

twitter_token <- create_token(app = appname,
                              consumer_key = consumer_key,
                              consumer_secret = consumer_secret,
                              access_token = access_token,
                              access_secret = access_secret)

message("Finished set token file")
get_token()

## --------------------- Main Function ---------------------
func_mainProgram <- function(excelFileName, pptFileName, username, numberRecord, sheetName){
  
  ## Copy powerpoint file from input folder to output folder
  inPowerPointFile <- paste0(getwd(),"/input/", pptFileName)
  outFolder <- paste0(getwd(),"/output")
  print(inPowerPointFile)
  print(outFolder)
  func_copyFile2Folder(inPowerPointFile, outFolder)
  
  func_write2ExcelTweet(excelFileName, func_getUserTweet_30Days_premium(username, numberRecord), sheetName)
  
  ## Call macro from excel file
  myFileName =  paste0(getwd(),"/output/", excelFileName)
  macroName = "Main"
  try(func_RCallExcelMacro(myFileName, macroName))
  message("Finish output to powerpoint file")

  ## create folder donald trump
  myNewFolder <- paste0(getwd(),"/output/",username)
  func_createFolder(myNewFolder)
  ## Copy 2 file and paste to folder donald trump
  outExcelFile <- paste0(getwd(),"/output/",excelFileName)
  outPowerPointFile <- paste0(getwd(),"/output/", pptFileName)
  func_moveFile2Folder(outExcelFile, myNewFolder)
  func_moveFile2Folder(outPowerPointFile, myNewFolder)
}

##-----------------------get and set arg
args <- commandArgs(trailingOnly = TRUE)
client_name <- as.character(args[1])
Twitter_id <- as.character(args[2])
# Twitter_id = "realDonaldTrump"

## --------------------- Test Output Excel Case ---------------------
username = Twitter_id
sheetName = "Twitter" 
excelFileName = "Twitter-Excel.xlsm"
pptFileName = "Twitter-PPT.pptx"
numberRecord = 50



print("========================= Input Variables =========================")
print(client_name)
print(Twitter_id)
print("========================= End Input Variables =========================")


# get data from Twitter -> output to excel file
func_mainProgram(excelFileName, pptFileName, username, numberRecord, sheetName)



