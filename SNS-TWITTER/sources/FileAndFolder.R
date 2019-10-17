# install.packages("jjb")
library(jjb)

func_createFolder <- function(fullPathName){
  if(!file.exists(fullPathName)) {
    mkdir(fullPathName)
    return(1)
  }else{
    return(0)
  }
}

func_copyFile2Folder <- function(fileName, newFolder){
  file.copy(from = fileName, to = newFolder,
            overwrite = TRUE, recursive = FALSE,
            copy.mode = TRUE)
}

func_moveFile2Folder <- function(fileName, newFolder){
  file.copy(from = fileName, to = newFolder,
            overwrite = TRUE, recursive = FALSE,
            copy.mode = TRUE)

  file.remove(fileName)
}
