## https://github.com/omegahat/RDCOMClient
# install_github("omegahat/RDCOMClient")

library("devtools")
library(RDCOMClient)

func_RCallExcelMacro <- function(pathFile, macroName){
  # Open a specific workbook in Excel:
  xlApp <- COMCreate("Excel.Application")

  # xlWbk <- xlApp$Workbooks()$Open("D:\\MyProjects\\Sources\\12. SNS Develop\\12_005\\12_005\\output\\Twitter-Excel.xlsm")
  xlWbk <- xlApp$Workbooks()$Open(pathFile)

  # this line of code might be necessary if you want to see your spreadsheet:
  xlApp[['Visible']] <- TRUE

  # Run the macro called "MyMacro":
  # xlApp$Run("MyAutomation")
  xlApp$Run(macroName)

  # Close the workbook and quit the app:
  xlWbk$Close(FALSE)
  xlApp$Quit()

  # Release resources:
  rm(xlWbk, xlApp)
  gc()
}