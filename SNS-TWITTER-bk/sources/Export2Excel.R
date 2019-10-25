# install.packages("magrittr")
# install.packages("imager")
# install.packages("openxlsx")
# ------------------------------------------------------- #

##library for exist exist excel file
library(openxlsx) 
##library for imager
library(magrittr)
library(imager)
# ------------------------------------------------------- #

##
## func_downloadImage will receive 2 variables
##  @params "link" is download link (as text format)
##  @params "imageName" is name of image file (without extends .jpg .png .jpeg)
##  return image name with extension
##
func_downloadImage <- function(link, imageName){
  arrSplit <- unlist(strsplit(link, ".", fixed=TRUE))
  extendsImage <- arrSplit[length(arrSplit)]
  imgName <- paste0(imageName,".", extendsImage)
  
  ## download image
  download.file(link,imgName, mode ='wb')
  return(imgName)
}

##
## func_write2ExcelTweet will write data to excel file (in folder output)
## params excelFileName: input excel file name
## params userData: input data as dataframe
## params sheetName: is what we want to create
##
func_write2ExcelTweet <- function(excelFileName, userData, sheetName){
  
  pathToReadExcelFile <- paste0(getwd(), "/input/", excelFileName)
  pathToSaveExcelFile <- paste0(getwd(), "/output/", excelFileName)
  print(pathToReadExcelFile)
  ## get all sheet in excel files
  myWorkbook <- loadWorkbook(pathToReadExcelFile);

  ## get list sheets name
  sheet_names <- sheets(myWorkbook);

  ## check sheetName is exist. If not exist create new sheet
  if (!(sheetName %in% sheets(myWorkbook)))
    addWorksheet(myWorkbook, sheetName)

  for(col in 1:ncol(userData)){
    if(col == 2){
      ## processing with image
      for(noRowImg in 1:nrow(userData[2])){
        if(!is.na(userData[noRowImg, 2])){
          ## download image
          imgName <- func_downloadImage(as.character(userData[noRowImg, 2]), noRowImg)
          ## check image
          chekImage = load.image(imgName)
          # plot(chekImage)
          ## insert image to excel file
          insertImage(myWorkbook, sheetName, imgName, width = 2.1, height = 1.33, startRow = noRowImg + 2,
                      startCol = col + 1, units = "in", dpi = 300)
        }
      }
    }else if(col == 7){
      ## write "follower" to Excel file
      writeData(myWorkbook, sheetName, userData[col], startCol = 3 + col,
                startRow = 3, colNames = FALSE, rowNames = FALSE)
    }else if(col == 1){ #Date column
      myDateColFormat <- format(userData[col], "%Y/%m/%d %H:%M")
      ## write "date" to Excel file
      writeData(myWorkbook, sheetName, myDateColFormat, startCol = 1 + col,
                startRow = 3, colNames = FALSE, rowNames = FALSE)
    }else{
      ## write "date, content, like, reweet" to Excel file
      writeData(myWorkbook, sheetName, userData[col], startCol = 1 + col,
                startRow = 3, colNames = FALSE, rowNames = FALSE)
    }
  }


  setRowHeights(myWorkbook, sheetName, rows = 1:nrow(userData)+2, heights = 93.75)


  ## save overwrite excel file
  saveWorkbook(myWorkbook, pathToSaveExcelFile, overwrite = TRUE)

  ## remove all image
  for(noRowImg in 1:nrow(userData[2])){
    if(!is.na(userData[noRowImg, 2])){
      imgName <- func_downloadImage(as.character(userData[noRowImg, 2]), noRowImg)
      # imgName <- paste0(noRowImg,".jpg")
      file.remove(imgName)
    }
  }
  
  message("Finish output excel file")
}
