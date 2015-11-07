<# Module Name:     MTSMSExcel.psm1
## 
## Author:          David Muegge
## Purpose:         Provides PowerShell Functions for various Microsoft Excel operations
##																					                                  
##                                                                             
####################################################################################################
## Disclaimer
##  ****************************************************************
##  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
##  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
##  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
##  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
##  ****************************************************************
###################################################################################################>

# Excel Constants
Set-Variable msoFalse -Value $false -Option Constant -ErrorAction SilentlyContinue 
Set-Variable msoTrue -Value $true -Option Constant -ErrorAction SilentlyContinue
Set-Variable xlExternal -Value 2 -Option Constant -ErrorAction SilentlyContinue
Set-Variable xlSum -Value (-4157) -Option Constant -ErrorAction SilentlyContinue
Set-Variable xlAverage -Value (-4106) -Option Constant -ErrorAction SilentlyContinue
Set-Variable xl3DArea -Value (-4098) -Option Constant -ErrorAction SilentlyContinue
Set-Variable xl3DAreaStacked -Value (78) -Option Constant -ErrorAction SilentlyContinue


<#
ChartType Constants
	
	
xl3DAreaStacked100	79
xl3DBarClustered	60
xl3DBarStacked	61
xl3DBarStacked100	62
xl3DColumn	-4100
xl3DColumnClustered	54
xl3DColumnStacked	55
xl3DColumnStacked100	56
xl3DLine	-4101
xl3DPie	-4102
xl3DPieExploded	70
xlArea	1
xlAreaStacked	76
xlAreaStacked100	77
xlBarClustered	57
xlBarOfPie	71
xlBarStacked	58
xlBarStacked100	59
xlBubble	15
xlBubble3DEffect	87
xlColumnClustered	51
xlColumnStacked	52
xlColumnStacked100	53
xlConeBarClustered	102
xlConeBarStacked	103
xlConeBarStacked100	104
xlConeCol	105
xlConeColClustered	99
xlConeColStacked	100
xlConeColStacked100	101
xlCylinderBarClustered	95
xlCylinderBarStacked	96
xlCylinderBarStacked100	97
xlCylinderCol	98
xlCylinderColClustered	92
xlCylinderColStacked	93
xlCylinderColStacked100	94
xlDoughnut	-4120
xlDoughnutExploded	80
xlLine	4
xlLineMarkers	65
xlLineMarkersStacked	66
xlLineMarkersStacked100	67
xlLineStacked	63
xlLineStacked100	64
xlPie	5
xlPieExploded	69
xlPieOfPie	68
xlPyramidBarClustered	109
xlPyramidBarStacked	110
xlPyramidBarStacked100	111
xlPyramidCol	112
xlPyramidColClustered	106
xlPyramidColStacked	107
xlPyramidColStacked100	108
xlRadar	-4151
xlRadarFilled	82
xlRadarMarkers	81
xlStockHLC	88
xlStockOHLC	89
xlStockVHLC	90
xlStockVOHLC	91
xlSurface	83
xlSurfaceTopView	85
xlSurfaceTopViewWireframe	86
xlSurfaceWireframe	84
xlXYScatter	-4169
xlXYScatterLines	74
xlXYScatterLinesNoMarkers	75
xlXYScatterSmooth	72
xlXYScatterSmoothNoMarkers	73

#>



# Helper functions



# Exported Functions

function New-ExcelApplication{
<#
.Synopsis
    Create Excel application object

.DESCRIPTION
    Create Excel application object

.EXAMPLE
    $ExcelApplication = New-ExcelApplication

#>
    [CmdletBinding()]
    Param([Switch]$Visible)

    Begin
    {
        # Create Excel Application
        $ExcelApplication = New-Object -comobject Excel.Application
        if($Visible){$ExcelApplication.visible = $Visible}
    }
    Process
    {
    }
    End
    {
        $ExcelApplication   
    }
} # New-ExcelApplication

function New-ExcelWorkbook{
<#
.Synopsis
   Open an exsiting Excel workbook

.DESCRIPTION
   Open an exsiting Excel workbook

.PARAMETER Path
    Full path to Excel workbook

.PARAMETER
    Excel Application Object

.EXAMPLE
   $ExcelWorkbook = OpenExclWokbook -Path c:\Temp\MyWorkbook.xlsx -ExcelApplication $ExeclApplication


#>
    [CmdletBinding()]
    Param
    (
        # Excel Application Object
        [Parameter(Mandatory=$true)]
        $ExcelApplication
 
    )

    Begin
    {
        $ExcelWorkbook = $ExcelApplication.Workbooks.Add()
    }
    Process
    {
    }
    End
    {
        $ExcelWorkbook
    }
} # New-ExcelWorkbook

function New-ExcelWorkSheet{
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelWorkbook,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$WorksheetName)


    $ExcelWorksheet = $ExcelWorkbook.Worksheets.Add()
    $ExcelWorksheet.Name = $WorksheetName
    $ExcelWorksheet.Activate()

    $ExcelWorksheet
} # New-ExcelWorkSheet

function Open-ExcelWorkbook{
<#
.Synopsis
   Open an exsiting Excel workbook

.DESCRIPTION
   Open an exsiting Excel workbook

.PARAMETER Path
    Full path to Excel workbook

.PARAMETER
    Excel Application Object

.EXAMPLE
   $ExcelWorkbook = OpenExclWokbook -Path c:\Temp\MyWorkbook.xlsx -ExcelApplication $ExeclApplication


#>
    [CmdletBinding()]
    Param
    (
        # Full path to Excel workbook
        [Parameter(Mandatory=$true)]
        $Path,
        # Excel Application Object
        [Parameter(Mandatory=$true)]
        $ExcelApplication



        
    )

    Begin
    {
        $ExcelWorkbook = $ExcelApplication.Workbooks.Open($Path)
    }
    Process
    {
    }
    End
    {
        $ExcelWorkbook
    }
} # Open-ExcelWorkbook

function Test-WorksheetExists{
<#

.SYNOPSIS
    Tests id excel worksheet exixts

.DESCRIPTION
    Tests id excel worksheet exixts

.EXAMPLE
    
    Test-WorksheetExists -ExcelWorkbook $ExcelWorkbook -WorksheetName "Test"

.NOTES

    

#>
[CmdletBinding(DefaultParameterSetName='Parameter Set 1',SupportsShouldProcess=$true,ConfirmImpact='Medium')]
Param(
        

        # Excel Workbook Object
        [Parameter(Mandatory=$true,ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("wb")]
        $ExcelWorkbook,

        # WorksheetName - Name of Excel worksheet
        [Parameter(Mandatory=$true,ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("ws")]
        [String]$WorksheetName

    ) 

    Try{
        if(($ExcelWorkbook.WorkSheets.Item($WorksheetName).Name -ne "")){return 1}
    }
    Catch{
        Write-Error $_
    }

} # Test-WorksheetExists

function Write-PSObjectToSheet{
<#
.SYNOPSIS
    Writes PSObject data to an Excel spreadsheet page

.DESCRIPTION
    Writes PSObject data to an Excel spreadsheet page

.EXAMPLE
    $Worksheet = Write-PSObjectToSheet -InputObject $TestObject -ExcelWorkbook $ExcelWorkbook -WorksheetName "Test01"

.EXAMPLE 


.NOTES

    Returns Worksheet

    Append sitch does not work **

#>
[CmdletBinding(DefaultParameterSetName='Parameter Set 1',SupportsShouldProcess=$true,ConfirmImpact='Medium')]
Param(
        # InputObject - Object to be sent to Excel Worksheet
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("io")] 
        $InputObject,

        # WorksheetName - Name of Excel worksheet
        [Parameter(Mandatory=$true,ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("wb")]
        $ExcelWorkbook,

        # WorksheetName - Name of Excel worksheet
        [Parameter(Mandatory=$true,ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("ws")]
        [String]$WorksheetName,

        # Switch Existing Sheet
        [Parameter(Mandatory=$False,ParameterSetName='Parameter Set 1')][Switch]$Append

    )    

    If($Append){
        if(Test-WorksheetExists -ExcelWorkbook $ExcelWorkbook -Worksheetname $WorksheetName){
            #$ExcelWorkbook.Worksheets.Item($WorksheetName).Activate()
        }
        else{
            $ExcelWorksheet = $ExcelWorkbook.Worksheets.Add()
	        $ExcelWorksheet.Name = $WorksheetName
	        $ExcelWorksheet.Activate()
        }
    }
    else{
        $ExcelWorksheet = $ExcelWorkbook.Worksheets.Add()
	    $ExcelWorksheet.Name = $WorksheetName
	    $ExcelWorksheet.Activate()
    }


    ### Magic Line
    $array = ($InputObject | ConvertTo-MultiArray).Value

    $starta = [int][char]'a' - 1
    if ($array.GetLength(1) -gt 26) {
        $col = [char]([int][math]::Floor($array.GetLength(1)/26) + $starta) + [char](($array.GetLength(1)%26) + $Starta)
    } else {
        $col = [char]($array.GetLength(1) + $starta)
    }

    $range = $ExcelWorksheet.Range("a1","$col$($array.GetLength(0))")
    $range.Value2 = $array


    $ExcelWorksheet
		
} # Write-PSObjectToSheet

function Write-SheetFormatting{
<#
.SYNOPSIS
    

.DESCRIPTION
    

.EXAMPLE
    

.EXAMPLE 


.NOTES
    
    May need to add options for differnt styles
    

#>
[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Medium')]
Param(
        # InputObject - Object to be sent to Excel Worksheet
        [Parameter(Mandatory=$true, 
                   Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("wk")] 
        $Worksheet,

        # WorksheetName - Name of Excel worksheet
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("ws")]
        $ExcelWorkbook,
        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [switch]$AutoFilter,
        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [switch]$FreezeTopPane,
        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [switch]$AutoFitCol



    )

    # Get Used Range of worksheet
    $d = $Worksheet.UsedRange

    if($AutoFilter){
        # Turn on autofilter
        $d.AutoFilter() | Out-Null
    }

    if($AutoFitCol){
        #resize the columns to fit the data
        $d.EntireColumn.AutoFit() | out-null
    }

    if($FreezeTopPane){
        # lock the first row
        $Worksheet.Activate()
        $Worksheet.Range("A2").Select() | Out-Null
        $ExcelWorkbook.Application.ActiveWindow.FreezePanes = $msoTrue
    }
    

} # Write-SheetFormatting

function Write-RollupSheetTotals{
<#
.SYNOPSIS
    

.DESCRIPTION
    

.EXAMPLE
    

.EXAMPLE 


.NOTES

    

#>
[CmdletBinding(DefaultParameterSetName='Parameter Set 1',SupportsShouldProcess=$true,ConfirmImpact='Medium')]
Param(
        # InputObject - Object to be sent to Excel Worksheet
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("wk")] 
        $Worksheet,

        # WorksheetName - Name of Excel worksheet
        [Parameter(Mandatory=$true,ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("dr")]
        [Int]$DeleteRow,

        # WorksheetName - Name of Excel worksheet
        [Parameter(Mandatory=$true,ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("tr")]
        [Int]$TotalRow

    )


    $Worksheet.Rows.Item($DeleteRow) = $null
    $endtotal = $totalrow - 1 
    $Worksheet.Cells.Item($TotalRow,1) = "=SUBTOTAL(9,A2:A$endtotal)"
    $Worksheet.Cells.Item($TotalRow,2) = "=SUBTOTAL(9,B2:B$endtotal)"
    $Worksheet.Cells.Item($TotalRow,3) = "=SUBTOTAL(9,C2:C$endtotal)"
    $Worksheet.Cells.Item($TotalRow,5) = "=SUBTOTAL(9,E2:E$endtotal)"
    $Worksheet.Cells.Item($TotalRow,6) = "=SUBTOTAL(9,F2:F$endtotal)"
    $Worksheet.Cells.Item($TotalRow,7) = "=SUBTOTAL(9,G2:G$endtotal)" 
    
} # Write-RollupSheetTotals

function New-ExcelDataConnection{
<#
.SYNOPSIS
    

.DESCRIPTION
    

.PARAMETER 
    

.PARAMETER 
    

.PARAMETER 
    

.PARAMETER 


.PARAMETER XLabelTickSpacing
    "Auto" - Excel will use auto spacing
    40 - Setting for 24 hours
    

.PARAMETER YLabelMaxValue


.EXAMPLE


.NOTES

    -- TODO --
    Add Y Axis Display Unit and label functionality. The below code is the VBA example
    ActiveChart.Axes(xlValue).DisplayUnit = xlHundreds
    ActiveChart.Axes(xlValue).HasDisplayUnitLabel = True
    ActiveChart.Axes(xlValue).DisplayUnitLabel.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "ms"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "ms"
    
#>

    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelWorkbook,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ConnectionName,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$SQLServer,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$SQLCatalog,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$SQLQuery)

        $ConnectionString = "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=$SQLCatalog;Data Source=$SQLServer;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=localhost;Use Encryption for Data=False;Tag with column collation when possible=False"
        $ExcelWorkbook.Connections.Add2($ConnectionName,"",$ConnectionString,$SQLQuery,2,$True,$True) | Out-Null
        $ExcelDBConnection = ($ExcelWorkbook.Connections) | Where-Object Name -EQ $ConnectionName


        $ExcelDBConnection

} # New-ExcelDataConnection

function New-ExcelPivotCache{
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelWorkbook,
          [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelDataConnection)

    $PCache = $ExcelWorkbook.PivotCaches().Create(2,$ExcelDataConnection,5)
    $PCache


} # New-ExcelPivotCache

function Import-Xlsx {
    <#
        .SYNOPSIS
            Import data from an Excel worksheet to a custom PSObject.
        .DESCRIPTION
            Imports data from an Excel worksheet to a custom PSObject. Imports from both .xls and .xlsx formats.
        .PARAMETER Path
            The path of the Excel file.
            If the parameters SheetName and SheetIndex are omitted you will be queried for which sheet to import. If there is only one worksheet in the workbook it will import that sheet without querying.
        .PARAMETER SheetName
            The name of the worksheet to import. Will throw an error if the worksheet name can not be found, or if the sheet with that name is not a worksheet.
            Can not be used in conjuction with parameter SheetIndex.
        .PARAMETER SheetIndex
            The index of the worksheet to import. Will throw an error if the worksheet index can not be found, or if the sheet with that index is not a worksheet.
            Can not be used in conjunction with parameter SheetName
        .PARAMETER IsTransposed
            Use this switch to indicate that the worksheet contains a transposed table.
            See Export-Xlsx.
        .PARAMETER HasTransposeColumnProperty
            Use this switch to indicate that the transposed table includes a header for each column. Enabled by default.
            This switch will be ignored when the parameter IsTransposed is not used.
            See Export-Xlsx.
        .PARAMETER HasTitle
            Use this switch to indicate that the worksheet contains a title.
            See Export-Xlsx.
        .EXAMPLE
            Import-Xlsx D:\UserData.xlsx
            Imports data from a .xlsx file. Will query for worksheet if there is more than one valid worksheet in the Excel file.
        .EXAMPLE
            Import-Xlsx D:\UserData.xlsx -SheetName "Disabled Users"
            Imports data from the worksheet with the name "Disabled Users"
        .NOTES
            Author : Gilbert van Griensven
            Website : http://www.itpilgrims.com/2013/01/import-xlsx/
    #>
    [CmdletBinding(DefaultParametersetName="Default")]
    Param (
        [Parameter(Position=0,Mandatory=$True)]
        [ValidateScript({
            $ReqExt = [System.IO.Path]::GetExtension($_)
            ($ReqExt -eq ".xls") -or
            ($ReqExt -eq ".xlsx")
        })]
        $Path,
        [Parameter(ParameterSetName="ByName")]
        [String] $SheetName,
        [Parameter(ParameterSetName="ByIndex")]
        [Int] $SheetIndex,
        [Switch] $IsTransposed,
        [Switch] $HasTransposeColumnProperty=$True,
        [Switch] $HasTitle
    )
    Function ReadData ($FromFile) {
        Add-Type -AssemblyName Microsoft.Office.Interop.Excel 
        $ExcelApplication = New-Object -ComObject Excel.Application
        $ExcelApplication.DisplayAlerts = $False
        $Workbook = $ExcelApplication.Workbooks.Open($FromFile)
        $Worksheets = $Workbook.Worksheets
        If ($Worksheets.Count -ge 1) {
            Switch ($PsCmdlet.ParameterSetName) {
                "Default" {
                    If ($Worksheets.Count -gt 1) {
                        $ChoiceDesc = New-Object System.Collections.ObjectModel.Collection[System.Management.Automation.Host.ChoiceDescription]
                        $SheetChoice = 0
                        $Script:Sheets = @()
                        $Worksheets |
                        % {
                            $Sheet = $_
                            $Script:Sheets += New-Object PSObject -Property @{Choice=$SheetChoice;SheetIndex=$Sheet.Index}
                            $ChoiceDesc.Add((New-Object "System.Management.Automation.Host.ChoiceDescription" -ArgumentList "&$($SheetChoice) $($Sheet.Name)"))
                            $SheetChoice++
                        }
                        $Result = $Host.UI.PromptForChoice("Import data from $($FromFile)","Please select the sheet to import:",$ChoiceDesc,0)
                        $SelectedSheet = $Script:Sheets | Where-Object -FilterScript {$_.Choice -eq $Result} | Select-Object -ExpandProperty SheetIndex
                        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet)) {}
                    } Else {
                        $SelectedSheet = $Worksheets | % {$_.Index}
                    }
                }
                "ByName" {
                    $SelectedSheet = ($Worksheets | Where-Object -FilterScript {$_.Name -eq $SheetName}).Index
                    If (!($SelectedSheet)) { $ExceptionMessage = "A worksheet with the name '$($SheetName)' can not be found in workbook $($Path) or is not of the type 'Worksheet'." }
                }
                "ByIndex" {
                    $SelectedSheet = ($Worksheets | Where-Object -FilterScript {$_.Index -eq $SheetIndex}).Index
                    If (!($SelectedSheet)) { $ExceptionMessage = "A worksheet with index '$($SheetIndex)' can not be found in workbook $($Path) or is not of the type 'Worksheet'." }
                }
            }
        } Else {
            $ExceptionMessage = "The workbook $($Path) does not contain any valid worksheets."
        }

        If (!($ExceptionMessage)) {
            $Workbook.Sheets.Item($SelectedSheet).Activate()
            $Script:Cols = $Workbook.ActiveSheet.usedRange.Columns.Count
            $Script:Rows = $Workbook.ActiveSheet.usedRange.Rows.Count
            $Script:Data = $Workbook.ActiveSheet.usedRange.Value2
        }

        $Workbook.Close()
        $ExcelApplication.Quit()
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Worksheets)) { }
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)) { }
        While ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApplication)) { }
        [GC]::Collect()

        If ($ExceptionMessage) {
            $ExceptionMessage = New-Object System.FormatException $ExceptionMessage
            Throw $ExceptionMessage
        }

    }

    If (Test-Path $Path) {
        ReadData $Path
        $Script:Headers = @()
        $Row = 2
        $HeaderOffset = 1
        If ($HasTitle) {
            $Row = $Row + 2
            $HeaderOffset = 3
        }
        If (!($IsTransposed)) {
            1..$Script:Cols | % { $Script:Headers += $Script:Data[$HeaderOffset,$_] }
            $Row..$Script:Rows | % {
                $CurrentRow = $_
                $Props = $Null
                1..$Script:Cols | % {
                    $Props += [Ordered]@{
                        $($Script:Headers[$_ - 1]) = "$($Script:Data[$CurrentRow,$_])"
                    }
                }
                New-Object PSObject -Property $Props
            }
        } Else {
            If (!($HasTransposeColumnProperty)) { $Row-- }
            $Row..$Script:Rows | % { $Script:Headers += $Script:Data[$_,1] }
            2..$Script:Cols | % {
                $CurrentCol = $_
                $Props = $Null
                $Row..$Script:Rows | % {
                    $Props += [Ordered]@{
                        $($Script:Headers[$_ - $Row]) = "$($Script:Data[$_,$CurrentCol])"
                    }
                }
                New-Object PSObject -Property $Props
            }
        }
    } Else {
        $ExceptionMessage = New-Object System.FormatException "The workbook $($Path) could not be found."
        Throw $ExceptionMessage
    }
} # Import-Xlsx

function New-ExcelChart{
<#
.SYNOPSIS
    

.DESCRIPTION
    

.PARAMETER 
    

.PARAMETER 
    

.PARAMETER 
    

.PARAMETER 


.PARAMETER XLabelTickSpacing
    "Auto" - Excel will use auto spacing
    40 - Setting for 24 hours
    

.PARAMETER YLabelMaxValue


.EXAMPLE


.NOTES

    -- TODO --
    Add Y Axis Display Unit and label functionality. The below code is the VBA example
    ActiveChart.Axes(xlValue).DisplayUnit = xlHundreds
    ActiveChart.Axes(xlValue).HasDisplayUnitLabel = True
    ActiveChart.Axes(xlValue).DisplayUnitLabel.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "ms"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "ms"
    
#>

    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelWorkbook,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelWorksheet,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelDataConnection,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ChartType,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)]$Height=300,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)]$Width=600)

    $ConnectionName = ($ExcelDataConnection.Name)
    
    # Setup pivotcache and pivot table
    $PCache = $ExcelWorkbook.PivotCaches().Create(2,$ExcelDataConnection,5)
    $PTable = $PCache.CreatePivotTable($ExcelWorksheet.Range("A1"),$ConnectionName)
    
    # Set chart type and create chart
    switch ($ChartType)
    {
        'Line' {$ChartType = 4}
        'StackedArea' {$ChartType = 76}
        'Bar' {$ChartType = 58}
        'Column' {$ChartType = 52}
        'Pie' {$ChartType = 5}
        Default {}
    }
    $PChart = $ExcelWorksheet.Shapes.AddCHart2(201,$ChartType,0,78,$Width,$Height)
    $PChart.Select()
    $ExcelWorkbook.ActiveChart.SetSourceData($ExcelWorksheet.Range("A1:C18"))
    $PChart.Name = $ConnectionName

    $ExcelWorksheet.Shapes.Item((($PChart.Name).ToString()))


} # New-ExcelChart

function Set-ExcelChartData{
<#
.SYNOPSIS
    

.DESCRIPTION
    

.PARAMETER 
    

.PARAMETER 
    

.PARAMETER 
    

.PARAMETER 


.PARAMETER XLabelTickSpacing
    "Auto" - Excel will use auto spacing
    40 - Setting for 24 hours
    

.PARAMETER YLabelMaxValue


.EXAMPLE


.NOTES

    -- TODO --
    Add Y Axis Display Unit and label functionality. The below code is the VBA example
    ActiveChart.Axes(xlValue).DisplayUnit = xlHundreds
    ActiveChart.Axes(xlValue).HasDisplayUnitLabel = True
    ActiveChart.Axes(xlValue).DisplayUnitLabel.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "ms"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "ms"
    
#>
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelWorkbook,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ChartItem,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ConnectionName,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$XAxisField,
        [Parameter(Mandatory=$True,ValueFromPipeline=$false)]$SheetQueryName,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)][ValidateSet("sum", "average")]$DataFunction="sum",
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)][String[]]$FilterFields,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)][String[]]$DataFields,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)][String[]]$LegendFields,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)][String[]]$SelectLegendFields

        )


    # Set Chart X Axis Data field
    $XField = $ChartItem.Chart.PivotLayout.PivotTable.CubeFields.Item("[$SheetQueryName].[$XAxisField]")
    $XField.Orientation = 1
    
    # Set chart filter fields
    if($FilterFields){
        foreach($ff in $FilterFields){
            $FField = $ChartItem.Chart.PivotLayout.PivotTable.CubeFields.Item("[$SheetQueryName].[$ff]")
            $FField.Orientation = 3
            $FField.EnableMultiplePageItems = $true
        }
    }

    # Set chart legend fields
    if($LegendFields){
        foreach($lf in $LegendFields){
            $LField = $ChartItem.Chart.PivotLayout.PivotTable.CubeFields.Item("[$SheetQueryName].[$lf]")
            $LField.Orientation = 2
            $LField.EnableMultiplePageItems = $true
        }
    }


    # Set chart value fields
    $CubeFields = $ChartItem.Chart.PivotLayout.PivotTable.CubeFields
    if($DataFields){
        foreach($df in $DataFields){

            switch ($DataFunction)
            {
                'sum' {$Measure = $CubeFields.GetMeasure("[$SheetQueryName].[$df]", $xlSum, "$df")}
                'average' {$Measure = $CubeFields.GetMeasure("[$SheetQueryName].[$df]", $xlAverage, "$df")}
            }
            $ExcelWorkbook.ActiveSheet.PivotTables($ConnectionName).AddDataField($ExcelWorkbook.ActiveSheet.PivotTables($ConnectionName).CubeFields.item($Measure.Name), $df) | Out-Null
        }
    }



} # Set-ExcelChartData

function Set-ExcelChartSelectedLegendFields{
<#
.Synopsis
    

.Description
    
     

.Example
    

.Example
    

.Parameter path
    

.Parameter name
    

.Parameter create
    

.inputs
    [string]

.outputs
    [string]

.Notes
    

.Link


#Requires -Version 4.0

#>
[CmdletBinding()]
Param([Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ExcelWorkSheet,
      [Parameter(Mandatory=$True,ValueFromPipeline=$false)][String]$PivotTableName,
      [Parameter(Mandatory=$True,ValueFromPipeline=$false)][String]$SheetQueryName,
      [Parameter(Mandatory=$True,ValueFromPipeline=$false)][String]$LegendFieldName,
      [Parameter(Mandatory=$True,ValueFromPipeline=$false)][String[]]$SelectFields)


    # Set Selected Legend/Column Fields

    $VisibleItemArray = @()
    foreach($sf in $SelectFields){
        
        $VisibleItemArray += "[$SheetQueryName].[$LegendFieldName].&[$sf]"

    }

    $ExcelWorkSheet.PivotTables($PivotTableName).PivotFields("[$SheetQueryName].[$LegendFieldName].[$LegendFieldName]").VisibleItemsList = $VisibleItemArray



    <#
    ActiveSheet.PivotTables("Memory").PivotFields( _
        "[20150320_TSC Query1].[CounterName].[CounterName]").VisibleItemsList = Array( _
        "[20150320_TSC Query1].[CounterName].&[Free MBytes]", _
        "[20150320_TSC Query1].[CounterName].&[Memctl Target MBytes]")
    #>


}

function Set-ExcelChartFormat{
<#
.SYNOPSIS
    

.DESCRIPTION
    

.PARAMETER 
    

.PARAMETER 
    

.PARAMETER 
    

.PARAMETER 


.PARAMETER XLabelTickSpacing
    "Auto" - Excel will use auto spacing
    40 - Setting for 24 hours
    

.PARAMETER YLabelMaxValue


.EXAMPLE


.NOTES

    -- TODO --
    Add Y Axis Display Unit and label functionality. The below code is the VBA example
    ActiveChart.Axes(xlValue).DisplayUnit = xlHundreds
    ActiveChart.Axes(xlValue).HasDisplayUnitLabel = True
    ActiveChart.Axes(xlValue).DisplayUnitLabel.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "ms"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "ms"
    
#>
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,ValueFromPipeline=$false)]$ChartItem,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)]$ChartTitle="None",
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)]$ChartHeight=300,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)]$ChartWidth=600,
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)]$XLabelTickSpacing="Auto",
        [Parameter(Mandatory=$False,ValueFromPipeline=$false)]$YLabelMaxValue="Auto")

    # Set Chart Formatting Properties
    $ChartItem.Chart.ShowAllFieldButtons = $false
    if($ChartTitle = "None"){
        $ChartItem.Chart.HasTitle = $false
    }else{
        $ChartItem.Chart.HasTitle = $true
        $ChartItem.Chart.ChartTitle.Text = $ChartTitle
    }
    $Chartitem.Height = $ChartHeight
    $Chartitem.Width = $ChartWidth
    if(-Not ($XLabelTickSpacing -eq "Auto")){$ChartItem.Chart.Axes(1).TickLabelSpacing = $XLabelTickSpacing}
    if(-Not ($YLabelMaxValue -eq "Auto")){$ChartItem.Chart.Axes(2).MaximumScale = $YLabelMaxValue}

} # Set-ExcelChartFormat



Export-ModuleMember -Variable msoFalse
Export-ModuleMember -Variable msoTrue
Export-ModuleMember -Variable xlExternal
Export-ModuleMember -Variable xlSum
Export-ModuleMember -Variable xlAverage

Export-ModuleMember -Function New-ExcelApplication
Export-ModuleMember -Function Open-ExcelWorkbook
Export-ModuleMember -Function New-ExcelWorkbook
Export-ModuleMember -Function New-ExcelWorkSheet
Export-ModuleMember -Function New-ExcelDataConnection
Export-ModuleMember -Function New-ExcelPivotCache

Export-ModuleMember -Function New-ExcelChart
Export-ModuleMember -Function Set-ExcelChartData
Export-ModuleMember -Function Set-ExcelChartFormat
Export-ModuleMember -Function Set-ExcelChartSelectedLegendFields
Export-ModuleMember -Function Import-Xlsx


Export-ModuleMember -Function Test-WorksheetExists
Export-ModuleMember -Function Write-PSObjectToSheet
Export-ModuleMember -Function Write-SheetFormatting
Export-ModuleMember -Function Write-RollupSheetTotals

