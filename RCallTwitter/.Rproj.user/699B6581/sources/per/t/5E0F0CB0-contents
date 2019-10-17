## https://github.com/omegahat/RDCOMClient
# install_github("omegahat/RDCOMClient")

library("devtools")
library(RDCOMClient)

func_RCallExcelMacro <- function(pathFile, macroName){
  # Create Excel Application
  xlApp <- COMCreate("Excel.Application")
  
  # Open the Macro Excel book
  xlWbk <- xlApp$Workbooks()$Open(pathFile)# Change to your directory
  
  
  # its ok to run macro without visible excel application
  # If you want to see your workbook, please set it to TRUE
  xlApp[['Visible']] <- TRUE 
  
  # Run the macro called "MyMacro": and Pass 10 and 30 as argument
  # Successful return would be NULL
  # xlApp$Run("Test_Add",10,30)
  xlApp$Run(macroName)
  #> NULL 
  
  # Close the workbook and quit the app:
  xlWbk$Close(FALSE)# not save and close excel book
  xlWbk$close(TRUE) # save and close excel book
  xlApp$Quit()
}

