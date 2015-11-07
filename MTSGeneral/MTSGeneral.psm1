<# Module Name:     MTSGeneral.psm1
## 
## Author:          David Muegge
## Purpose:         Provides PowerShell functions for various general purpose operations
##																					                              
##                                                                             
####################################################################################################
## Disclaimer
## ****************************************************************
## * THE (MTSGeneral PowerShell Module)                           *
## * IS PROVIDED WITHOUT WARRANTY OF ANY KIND.                    *
## *                                                              *
## * This module is licensed under the terms of the MIT license.  *
## * See license.txt in the root of the github project            *
## *                                                              *
## **************************************************************** 
###################################################################################################>
			

				

Function Get-MyModule{
<#
.Synopsis
    Checks for loaded and existing modules

.Description
    If modules exists and is not loaded it wil load module and return $true
    If module is loaded it will return $true
    If module is not loaded and does not exist on system it will return $false
     
.Example
    get-mymodule -name "bitsTransfer"    

.Example
    If(Get-MyModule –name “bitsTransfer”) { call your bits code here }
    ELSE { “Bits module is not installed on this system.” ; exit }

.Parameter name
    Name on module


.inputs
    [string]

.outputs
    [boolean]

.Notes
    From Technet Scripting Guys

.Link
    http://blogs.technet.com/b/heyscriptingguy/archive/2010/07/11/hey-scripting-guy-weekend-scripter-checking-for-module-dependencies-in-windows-powershell.aspx

#Requires -Version 4.0

#> 
Param([string]$name) 
if(-not(Get-Module -name $name)){ 
    if(Get-Module -ListAvailable | Where-Object { $_.name -eq $name }){ 
        Import-Module -Name $name 
        $true 
    } #end if module available then import 
    else { $false } #module not available 
} # end if not module 
else { $true } #module already loaded 
} # Get-MyModule 

Function Get-LogNameFromDate{
<#
.Synopsis
    Creates a log file and generates the name based on parameter values and date

.Description
    Creates a log file and generates name in the followinf format
     

.Example
    Get-logNameFromDate -path "c:\fso" -name "log"
    Creates a file name like c:\fso\log20100914-122019.Txt but does not
    create the file. It returns the file name to calling code.

.Example
    Get-logNameFromDate -path "c:\fso" -name "log" -suffix "txt" -Create
    Creates a file name like c:\fso\log20100914-122019.Txt and
    create the file. It returns the file name to calling code.

.Parameter path
    path to log file

.Parameter name
    base name of log file

.Parameter create
    switch that determines whether log file or only name is created

.inputs
    [string]

.outputs
    [string]

.Notes
    NAME:  Get-LogNameFromDate
    AUTHOR: ed wilson, msft
    LASTEDIT: 09/10/2010 16:58:06
    KEYWORDS: parameter substitution, format specifier, string substitution
    HSG: WES-09-25-10

    David Muegge: 7/21/2013 - added suffix parameter

.Link
    Http://www.ScriptingGuys.com

#Requires -Version 2.0

#>
    [CmdletBinding()]
    Param(
    [string]$path = "c:\fso",
    [string]$name = "log",
    [string]$suffix = "txt",
    [switch]$Create
    )
    $logname = "{0}\{1}{2}.{3}" -f $path,$name,(Get-Date -Format yyyyMMdd-HHmmss),$suffix
    if($create) 
    { 
        New-Item -Path $logname -ItemType file -force | out-null
        $logname
    }
    else {$logname}

} # Get-logNameFromDate

function Get-DateString{
<#
.Synopsis
   Returns a long date string

.DESCRIPTION
   Returns a long date string in the format YYYYMMDDhhmmss to be used fot unique file naming

.EXAMPLE
   [String]$filedatestring = Get-FileDateString
#>

    [CmdletBinding()]
    [OutputType([String])]
    Param()

    [String]$filedate = "{0}" -f (Get-Date -Format yyyyMMddHHmmss)
    return $filedate

} # Get-DateString

function Get-TimeStampString{
<#
.Synopsis
   Returns a time stamp string

.DESCRIPTION
   Returns a TimeStamp string in the format YYYYMMDD-hh:mm:ss:mm 

   Primarily for generating strings for log file time stamps

.EXAMPLE
   [String]$TimeStampString = Get-TimeStampString
#>

    [CmdletBinding()]
    [OutputType([String])]
    Param()

    $NowValue = [Datetime]::Now
    [String]$TimeStampString = ($NowValue.Year.ToString() + $NowValue.Month.ToString() + $NowValue.Day.ToString() + "-" + $NowValue.Hour.ToString() + ":" + $NowValue.Minute.ToString() + ":" + $NowValue.Second.ToString() + ":" + $NowValue.Millisecond.ToString())
    return $TimeStampString

} # Get-TimeStampString

function ConvertTo-MultiArray {
 <#
 .Notes
 NAME: ConvertTo-MultiArray
 AUTHOR: Tome Tanasovski
 Website: http://powertoe.wordpress.com
 Twitter: http://twitter.com/toenuff
 Version: 1.0
 CREATED: 11/5/2010
 LASTEDIT:
 11/5/2010 1.0
 Initial Release
 11/5/2010 1.1
 Removed array parameter and passes a reference to the multi-dimensional array as output to the cmdlet
 11/5/2010 1.2
 Modified all rows to ensure they are entered as string values including $null values as a blank ("") string.

 .Synopsis
 Converts a collection of PowerShell objects into a multi-dimensional array

 .Description
 Converts a collection of PowerShell objects into a multi-dimensional array.  The first row of the array contains the property names.  Each additional row contains the values for each object.

 This cmdlet was created to act as an intermediary to importing PowerShell objects into a range of cells in Exchange.  By using a multi-dimensional array you can greatly speed up the process of adding data to Excel through the Excel COM objects.

 .Parameter InputObject
 Specifies the objects to export into the multi dimensional array.  Enter a variable that contains the objects or type a command or expression that gets the objects. You can also pipe objects to ConvertTo-MultiArray.

 .Inputs
 System.Management.Automation.PSObject
        You can pipe any .NET Framework object to ConvertTo-MultiArray

 .Outputs
 [ref]
        The cmdlet will return a reference to the multi-dimensional array.  To access the array itself you will need to use the Value property of the reference

 .Example
 $arrayref = get-process |Convertto-MultiArray

 .Example
 $dir = Get-ChildItem c:\
 $arrayref = Convertto-MultiArray -InputObject $dir

 .Example
 $range.value2 = (ConvertTo-MultiArray (get-process)).value

 .LINK

http://powertoe.wordpress.com

#>
 [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [PSObject[]]$InputObject
    )
    BEGIN {
        $objects = @()
        [ref]$array = [ref]$null
    }
    Process {
        $objects += $InputObject
    }
    END {
        $properties = $objects[0].psobject.properties |%{$_.name}
        $array.Value = New-Object 'object[,]' ($objects.Count+1),$properties.count
        # i = row and j = column
        $j = 0
        $properties |%{
            $array.Value[0,$j] = $_.tostring()
            $j++
        }
        $i = 1
        $objects |% {
            $item = $_
            $j = 0
            $properties | % {
                if ($item.($_) -eq $null) {
                    $array.value[$i,$j] = ""
                }
                else {
                    $array.value[$i,$j] = $item.($_).tostring()
                }
                $j++
            }
            $i++
        }
        $array
    }
} # ConvertTo-MultiArray

function Select-Random{
<#
.Synopsis
    Selects a random element from collection

.Description
    Selects a random element from the collection either passed as a parameter or input via the pipeline.
    If the collection is passed in as an argument, we simply pick a random number between 0 and count-1
    for each element you want to return, but when processing pipeline input we want to keep memory use 
    to a minimum, so we use a "reservoir sampling" algorithm[1].

    [1] http://gregable.com/2007/10/reservoir-sampling.html

    The script stores $count elements (the eventual result) at all times. It continues processing 
    elements until it reaches the end of the input. For each input element $n (the count of the inputs 
    so far) there is a $count/$n chance that it becomes part of the result.
    For each previously selected element, there is a $count/($n-1) chance of it being selected 
    For the ones selected, there's a ($count/$n * 1/$count = 1/$n) chance of it being replaced, so a 
    ($n-1)/$n chance of it remaining ... thus, it's cumulative probability of being among the selected
    elements after the nth input is processed is $count/($n-1) * ($n-1)/$n = $count/$n, as it should be.
     

.Example
    $arr = 1..5; Select-Random $arr
    1..10 | Select-Random -Count 2

.Notes
    Author: Joel "Jaykul" Bennett
    Version: 2.2.0.0

.History
    2.0.0.0: Rewrote using the reservoir sampling technique
    2.1.0.0: Fixed a bug in 2.0 which inverted the probability and resulted in the last n items being selected with VERY high probability
    2.2.0.0: Use more efficient direct random sampling if the collection is passed as an argument


#Requires -Version 4.0

#>
[CmdletBinding()]
param([int]$count=1, [switch]$collectionMethod, [array]$inputObject=$null) 

BEGIN {
   if ($args -eq '-?') {
@"
Usage: Select-Random [[-Count] <int>] [-inputObject] <array> (from pipeline) [-?]

Parameters:
 -Count            : The number of elements to select.
 -inputObject      : The collection from which to select a random element.
 -collectionMethod : Collect the pipeline input instead of using reservoir
 -?                : Display this usage information and exit

Examples:
 PS> $arr = 1..5; Select-Random $arr
 PS> 1..10 | Select-Random -Count 2

"@
exit
   } 
   else
   {
      $rand = new-object Random
      if ($inputObject) 
      {
         # Write-Output $inputObject | &($MyInvocation.InvocationName) -Count $count
      }
      elseif($collectionMethod)
      {
         Write-Verbose 'Collecting from the pipeline '
         [Collections.ArrayList]$inputObject = new-object Collections.ArrayList
      }
      else
      {
         $seen = 0
         $selected = new-object object[] $count
      }
   }
}
PROCESS {
   if($_)
   {
      if($collectionMethod)
      {
         $inputObject.Add($_) | out-null
      } else {
         $seen++
         if($seen -lt $count) {
            $selected[$seen-1] = $_
         } ## For each input element $n there is a $count/$n chance that it becomes part of the result.
         elseif($rand.NextDouble() -lt ($count/$seen))
         {
            ## For the ones previously selected, there's a 1/$n chance of it being replaced
            $selected[$rand.Next(0,$count)] = $_
         }
      }
   }
}
END {
   if (-not $inputObject)
   {  ## DO ONCE: (only on the re-invoke, not when using -inputObject)
      Write-Verbose "Selected $count of $seen elements."
      Write-Output $selected
      # foreach($el in $selected) { Write-Output $el }
   } 
   else 
   {
      Write-Verbose ('{0} elements, selecting {1}.' -f $inputObject.Count, $Count)
      foreach($i in 1..$Count) {
         Write-Output $inputObject[$rand.Next(0,$inputObject.Count)]
      }   
   }
}

} # Select-Random

function ConvertTo-CSVStringFromArray{
<#
.Synopsis
   Converts a one dimensional array to a csv sting list

.DESCRIPTION
   Converts a one dimensional array to a csv sting list

.Parameter InputObject
    Array to convert

.PARAMETER Quotes
    Specifies force quotes or no quotes
    $true = Force Quotes
    $false = Force NoQuotes

.EXAMPLE
   [String]$csvstring = ConvertTo-CSVStringFromArray
#>

    [CmdletBinding()]
    [OutputType([String])]
    Param($InputObject,$Quotes=$false)

    if($Quotes){

        Foreach($Value in $InputObject) { 

            $value = $Value.Replace("""","") + ","
            $value = """" + $value + """"
            $csvlist += $value
            
        }
    
        $lastcomma = $csvlist.LastIndexOf(",")
        $csvlist = $csvlist.Remove($lastcomma, 1)
    }
    else
    {
        Foreach($Value in $InputObject) { 
            
            $csvlist += $Value.Replace("""","") + ","
            
        }
    
        $lastcomma = $csvlist.LastIndexOf(",")
        $csvlist = $csvlist.Remove($lastcomma, 1)
    }

    $csvlist

} # ConvertTo-CSVStringFromArray

function Convert-Size {            
[cmdletbinding()]            
param(            
    [validateset("Bytes","KB","MB","GB","TB")]            
    [string]$From,            
    [validateset("Bytes","KB","MB","GB","TB")]            
    [string]$To,            
    [Parameter(Mandatory=$true)]            
    [double]$Value,            
    [int]$Precision = 4            
)            
switch($From) {            
    "Bytes" {$value = $Value }            
    "KB" {$value = $Value * 1024 }            
    "MB" {$value = $Value * 1024 * 1024}            
    "GB" {$value = $Value * 1024 * 1024 * 1024}            
    "TB" {$value = $Value * 1024 * 1024 * 1024 * 1024}            
}            
            
switch ($To) {            
    "Bytes" {return $value}            
    "KB" {$Value = $Value/1KB}            
    "MB" {$Value = $Value/1MB}            
    "GB" {$Value = $Value/1GB}            
    "TB" {$Value = $Value/1TB}            
            
}            
            
return [Math]::Round($value,$Precision,[MidPointRounding]::AwayFromZero)            
            
} # Convert-Size

Function ConvertFrom-EpochDateTime{
    [cmdletbinding()]
    param
    (
        [Object]$EpochDateTime,
        [Switch]$InSeconds,
        [Switch]$InMilliseconds
        
    )

    If($InSeconds){[timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($EpochDateTime))}
    If($InMilliseconds){[timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddMilliseconds($EpochDateTime))}
    
} # ConvertFrom-EpochDateTime

Function ConvertTo-EpochDateTime{
    [cmdletbinding()]
    param
    (
        [DateTime]$DateTime,
        [Switch]$InSeconds,
        [Switch]$InMilliseconds
        
    )

    $epoch = Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
    If($InSeconds){[math]::truncate($DateTime.ToUniversalTime().Subtract($epoch).TotalSeconds)}
    If($InMilliseconds){[math]::truncate($DateTime.ToUniversalTime().Subtract($epoch).TotalMilliSeconds)}


} # ConvertTo-EpochDateTime





Export-ModuleMember -Function Get-MyModule
Export-ModuleMember -Function Get-logNameFromDate
Export-ModuleMember -Function Get-DateString
Export-ModuleMember -Function ConvertTo-MultiArray
Export-ModuleMember -Function Get-TimeStampString
Export-ModuleMember -Function Select-Random
Export-ModuleMember -Function ConvertTo-CSVStringFromArray
Export-ModuleMember -Function Convert-Size
Export-ModuleMember -Function ConvertFrom-EpochDateTime
Export-ModuleMember -Function ConvertTo-EpochDateTime
				