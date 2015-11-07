<# Module Name:     MTSChart.psm1
##
## Author:          David Muegge
## Purpose:         Provides PowerShell functions for charting PowerShell objects using .Net MSChart controls 
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
####################################################################################################>

# load the appropriate assemblies 
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

Function Out-DataTable {
<#
.SYNOPSIS
    Outputs an ADO.Net datatable from an powershell object

.DESCRIPTION
    Outputs an ADO.Net datatable from an powershell object

.PARAMETER InputObject
    PowerShell Object to convert to DataTable

.EXAMPLE
    $ChartDataTable = Out-DataTable $ChartObject      

.NOTES
    


#>
	[CmdletBinding()]
	
	param ( 
		[Parameter(Mandatory=$True,ValueFromPipeline=$True)]$InputObject
	)
	
	
	Begin{
		
		$dt = new-object Data.datatable  
  		$First = $true
	}
	Process{
    
		foreach ($item in $InputObject){  
		$DR = $DT.NewRow()  

		$itemproperties = $Item.PsObject.get_properties()

		foreach($prop in $itemproperties){
			if ($first) {  
		        $Col =  new-object Data.DataColumn  
		        $Col.ColumnName = $prop.Name.ToString()  
		        $DT.Columns.Add($Col)
			}
			if ($prop.value -eq $null) {  
				$DR.Item($prop.Name) = "[empty]"  
			}   
			else {  
				$DR.Item($prop.Name) = $prop.value  
			}  
		}

		$DT.Rows.Add($DR)  
		$First = $false  
		} 

	}
	END{
  		return $dt
  	}

} # Out-DataTable

function Out-MTSChart{
<#
.SYNOPSIS
    Create chart using MSChart controls

.DESCRIPTION
    Generate Chart from a PowerShell object

.PARAMETER InputObject
    PowerShell Object to generate chart
    
.PARAMETER YValues
    Comma delimited list of properties to plot on Y axis of chart
    
.PARAMETER XValue
    PDH Time property for XAxis chart values
    
.PARAMETER XInterval
    
    
.PARAMETER 
    
.PARAMETER 
    
.EXAMPLE
    out-mtschart -InputObject $PDHdata -XValue $XValueField -Width $chartwidth -Height $chartheight -ChartTitle $ChartTitle -ChartType Line -ChartFullPath $Chartfile -YValues $counterlist | Out-Null      

.NOTES
    


#>
    
	[CmdletBinding()]
	
	param ( 
		[Parameter(Mandatory=$True,ValueFromPipeline=$True)]$InputObject, `
		[Parameter(Mandatory=$True,ValueFromPipeline=$False)][string]$YValues, `
		[Parameter(Mandatory=$True,ValueFromPipeline=$False)][string]$XValue, `
        [Parameter(Mandatory=$True,ValueFromPipeline=$False)][int]$XInterval, `
		[Parameter(Mandatory=$True,ValueFromPipeline=$False)][int]$Width, `
		[Parameter(Mandatory=$True,ValueFromPipeline=$False)][int]$Height, `
		[Parameter(Mandatory=$True,ValueFromPipeline=$False)][string]$ChartTitle, `
		[Parameter(Mandatory=$True,ValueFromPipeline=$False)][string]$ChartType, `
        [Parameter(Mandatory=$false,ValueFromPipeline=$False)][switch]$LegendOn, `
		[Parameter(Mandatory=$True,ValueFromPipeline=$False)][string]$ChartFullPath,
        [Parameter(Mandatory=$False,ValueFromPipeline=$False)][string]$ChartFileType="png"
	)
	
	BEGIN{	
	
	    # create chart object 
	    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
		
		# Set global chart settings
	    $Chart.Width = $width 
	    $Chart.Height = $height
		[void]$Chart.Titles.Add([string]$chartTitle)
		$Chart.BackColor = [System.Drawing.Color]::White
		$Chart.BackGradientStyle = [System.Windows.Forms.DataVisualization.Charting.GradientStyle]::LeftRight 
		$Chart.BackSecondaryColor = [System.Drawing.Color]::white
		$Chart.BorderColor = [System.Drawing.Color]::white
		$Chart.BorderWidth = 10
		$Chart.BorderlineColor = [System.Drawing.Color]::Black
		$Chart.BorderlineWidth = 10
        
        
        $Chart.Palette = [System.Windows.Forms.DataVisualization.Charting.ChartColorPalette]::Bright
                
        # Create legend and define properties
        $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
        $Legend.MaximumAutoSize = 30
        if($LegendOn){$chart.Legends.Add($Legend)}
                

		# create a chartarea to draw on and add to chart 
	    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea

        $chartArea.AxisX.MajorGrid.LineWidth = 0
        $chartArea.AxisY.MajorGrid.LineWidth = 1
		$chartArea.AxisX.Minimum = 0
		$chartArea.AxisX.Maximum = $InputObject.count
        $chartArea.AxisX.LineWidth = 1
		$chartArea.AxisX.Interval = $XInterval
        $chartArea.BorderWidth = 6
        $chartarea.Position.Y = 7
        $chartarea.Position.x = 1
        $chartarea.Position.Height = 90
        $chartarea.BackColor = [System.Drawing.Color]::white
        if($LegendOn){$chartarea.Position.Width = 80}Else{$chartarea.Position.Width = 90}
        $chartArea.AxisY.LineWidth = 0
	    $Chart.ChartAreas.Add($chartArea)
		
		# Set Chart Type and series data and legends
		switch($ChartType){
	
			"Line"{
			
				# Set series members names for the X and Y values
				$YProperties = $YValues.ToString().Split(",")
				foreach($prop in $YProperties){
					[void]$Chart.Series.Add($prop)
                    #$XValue
					$Chart.Series[$prop].XValueMember = $XValue
					$Chart.Series[$prop].YValueMembers = $prop         
					
					# Set series properties
					$Chart.Series[$prop].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
                    if($LegendOn){
					    $Chart.Legends.Add($prop)
                        
                    }
                    # Chart series line width
                    $chart.Series[$prop].BorderWidth = 2
					
				}

			}
			
			"Bar"{
			
			}
			
			"Pie"{
			
			}
            "StackedArea"{
			
                # Set series members names for the X and Y values
				$YProperties = $YValues.ToString().Split(",")
				foreach($prop in $YProperties){
					[void]$Chart.Series.Add($prop)
					$Chart.Series[$prop].XValueMember = $XValue
					$Chart.Series[$prop].YValueMembers = $prop
                    
					# Set series properties
					$Chart.Series[$prop].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::StackedArea
                    if($LegendOn){
					    $Chart.Legends.Add($prop)
                    }		
				}
			}
		}

		# Set array to hold pipeline objects
		$ChartObject = @()
		
	}
	PROCESS{
	
		$ChartObject += $InputObject
	
	}
	END{
		
		# Set chart data source 
		$ChartDataTable = Out-DataTable $ChartObject
		$Chart.DataSource = $ChartDataTable
		
		# Data bind to the selected data source 
		$Chart.DataBind()
	    
		# Save chart to file
		$Chart.SaveImage($ChartFullPath, $ChartFileType)
	
	}
} # Out-MTSChart

Export-ModuleMember Out-DataTable
Export-ModuleMember Out-MTSChart

