# Excel Manager For PowerShell
## Must use 
![PowerShell](https://img.shields.io/badge/PowerShell_5-5391FE?style=for-the-badge&logo=powershell&logoColor=white)

## Advencement
Only the starter of this module but at the end it will excel export more easier with powershell
Feel free to suggest changes !

## Tips
  Declaration: $myExcel = [ExcelManager]@{}
  
  End of script: $myExcel.delete()

## Methods guide

### Constructor
  There is two constructor one with no param and the other where you can pass your excel.application to link it
  
### [void] addWorkSheet()
  Simply add a new worksheet to the book with default name
  
### [void] addWorkSheet([string] $name)
  Do the same but change the default name of the worksheet
   
### [void] renameWorkSheet([int] $index, [string] $name)
  Rename a worksheet by its index (lastest create will have the bigger index)
   
### [System.MarshalByRefObject] getWorkSheet([int] $index)
  Return a worksheet by its index
  
### [System.Collections.ArrayList] getWorkSheet()
  Return the array of worksheet
  
### [int] getWorkSheetLenght()
  Return the number of worksheet
  
### [bool] saveAs([string] $path, [string] $name)
  Take path and name to save your xlsx file, return false if there is an error during the saving
  
### [void] delete(){
  Always run this method when you don't need excel anymore
  It delete the excel.application and prevent excel running in background
