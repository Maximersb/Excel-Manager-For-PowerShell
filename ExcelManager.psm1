class ExcelManager{
    hidden [ValidateNotNullOrEmpty()][System.Collections.ArrayList]$workSheetList 
    hidden [ValidateNotNullOrEmpty()][System.__ComObject]$excelApp
    hidden [ValidateNotNullOrEmpty()][System.MarshalByRefObject]$workBook
    

    ExcelManager() {
        $this.excelApp = New-Object -ComObject excel.application
        $this.workBook = $this.excelApp.workbooks.add()
        $this.workSheetList= @($this.workBook.worksheets.item(1))
        if($null -ne (Get-Process -Name EXCEL -ErrorAction SilentlyContinue -ErrorVariable ProcessError))
            Stop-Process -Name EXCEL
    }

    [void] addWorkSheet(){
        $this.workSheetList += $this.workBook.worksheet.add([System.Reflection.Missing]::Value,$this.workSheetList[$this.workSheetList.Lenght-1])
    }

    [void] addWorkSheet([string] $name){
        $this.workSheetList += $this.workBook.worksheet.add([System.Reflection.Missing]::Value,$this.workSheetList[$this.workSheetList.Lenght-1])
        $this.workSheetList[$this.workSheetList.Lenght].name = $name
    }

    [void] renameWorkSheet([int] $index, [string] $name){
        $this.workSheetList[$index].name = $name
    }

    [bool] saveAs([string] $path, [string] $name){
        try{
            $this.excelApp.displayalerts = $false
            $this.workBook.Saveas($path + \ + $name)
            $this.excelApp.displayalerts = $true
            return $True
        }catch{
            return $False
        }
    }

    [System.MarshalByRefObject] getWorkSheet([int] $index){
        return $this.workSheetList[$index]
    }

    [void] endManagment(){
        $excelApp.Quit()| Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)| Out-Null
        Remove-Variable excelApp | Out-Null
    }
}