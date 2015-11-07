<# Module Name:     MTSMSSQL.psm1
## 
## Author:          David Muegge
## Purpose:         Provides PowerShell functions for interacting with Microsoft SQL Server
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


Import-Module Sqlps


try {add-type -AssemblyName "Microsoft.SqlServer.ConnectionInfo, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" -EA Stop} 
catch {add-type -AssemblyName "Microsoft.SqlServer.ConnectionInfo"} 
 
try {add-type -AssemblyName "Microsoft.SqlServer.Smo, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845dcd8080cc91" -EA Stop}  
catch {add-type -AssemblyName "Microsoft.SqlServer.Smo"}  


# Helper Functions
function Get-Type{ 
    param($type) 
 
$types = @( 
'System.Boolean', 
'System.Byte[]', 
'System.Byte', 
'System.Char', 
'System.Datetime', 
'System.Decimal', 
'System.Double', 
'System.Guid', 
'System.Int16', 
'System.Int32', 
'System.Int64', 
'System.Single', 
'System.UInt16', 
'System.UInt32', 
'System.UInt64') 
 
    if ( $types -contains $type ) { 
        Write-Output "$type" 
    } 
    else { 
        Write-Output 'System.String' 
         
    } 
} #Get-Type

function Get-SqlType{  
    param([string]$TypeName)  
  
    switch ($TypeName)   
    {  
        'Boolean' {[Data.SqlDbType]::Bit}  
        'Byte[]' {[Data.SqlDbType]::VarBinary}  
        'Byte'  {[Data.SQLDbType]::VarBinary}  
        'Datetime'  {[Data.SQLDbType]::DateTime}  
        'Decimal' {[Data.SqlDbType]::Decimal}  
        'Double' {[Data.SqlDbType]::Float}  
        'Guid' {[Data.SqlDbType]::UniqueIdentifier}  
        'Int16'  {[Data.SQLDbType]::SmallInt}  
        'Int32'  {[Data.SQLDbType]::Int}  
        'Int64' {[Data.SqlDbType]::BigInt}  
        'UInt16'  {[Data.SQLDbType]::SmallInt}  
        'UInt32'  {[Data.SQLDbType]::Int}  
        'UInt64' {[Data.SqlDbType]::BigInt}  
        'Single' {[Data.SqlDbType]::Decimal} 
        default {[Data.SqlDbType]::VarChar}  
    }  
      
} #Get-SqlType 

function Test-IsDBNull{
<#
.SYNOPSIS
	Tests to see if a value is a SQL NULL or not

.DESCRIPTION
	Returns $true if the value is a SQL NULL.

.PARAMETER  value
	The value to test

	

.EXAMPLE
	PS C:\> Is-NULL $row.columnname

	
.INPUTS
    None.
    You cannot pipe objects to New-Connection

.OUTPUTS
	Boolean


.NOTES
    From adolib by Mike Sheppard

#>
[CmdletBinding()]
  param([Parameter(Position=0, Mandatory=$true)]$value)
  return  [System.DBNull]::Value.Equals($value)
} # Test-IsDBNull


# Exported Functions

function Get-SQLConnection{
<#
.SYNOPSIS
    Get SQL Database Connection Object

.DESCRIPTION
    

.PARAMETER SQLServer
    SQL Server hostname or ipaddress

.PARAMETER SQLCatalog
    SQL Catalog Name

.PARAMETER Username
    SQL Username

.PARAMETER Password
    SQL Password

.EXAMPLE
    Get-SQLConnection -SQLServer $SQLServer -SQLCatalog $SQLCatalog

.NOTES
    

#>
[CmdletBinding()]
    Param([Parameter(Mandatory=$True)]$SqlServer,
          [Parameter(Mandatory=$True)]$SQLCatalog,
          [Parameter(Mandatory=$False)]$Username=$null,
          [Parameter(Mandatory=$False)]$Password=$null)

    # Setup SQL connection
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    if($username){
        $connectionString = "Server = $SqlServer; Database = $SqlCatalog; Integrated Security = False User Id=$Username; Password=$Password;"
    }
    else{
        $connectionString = "Server = $SqlServer; Database = $SqlCatalog; Integrated Security = True"        
    }
	$SqlConnection.ConnectionString = $connectionString

    $SqlConnection

} # Get-SQLConnection

function New-Connection{
<#
.SYNOPSIS
	Create a SQLConnection object with the given parameters

.DESCRIPTION
	This function creates a SQLConnection object, using the parameters provided to construct the connection string.  You may optionally provide the initial database, and SQL credentials (to use instead of NT Authentication).

.PARAMETER  Server
	The name of the SQL Server to connect to.  To connect to a named instance, enclose the server name in quotes (e.g. "Laptop\SQLExpress")

.PARAMETER  Database
	The InitialDatabase for the connection.
	
.PARAMETER  User
	The SQLUser you wish to use for the connection (instead of using NT Authentication)
        
.PARAMETER  Password
	The password for the user specified by the User parameter.

.EXAMPLE
	PS C:\> New-Connection -server MYSERVER -database master

.EXAMPLE
	PS C:\> Get-Something -server MYSERVER -user sa -password sapassword

.INPUTS
    None.
    You cannot pipe objects to New-Connection

.OUTPUTS
	System.Data.SqlClient.SQLConnection

.NOTES
    From adolib by Mike Sheppard

#>
[CmdletBinding()]
param([Parameter(Position=0, Mandatory=$true)][string]$server, 
      [Parameter(Position=1, Mandatory=$false)][string]$database='',
      [string]$user='',
      [string]$password='')

	if($database -ne ''){
	  $dbclause="Database=$database;"
	}
	$conn=new-object System.Data.SqlClient.SQLConnection
	
	if ($user -ne ''){
		$conn.ConnectionString="Server=$server;$dbclause`User ID=$user;Password=$password;Pooling=false"
	} else {
		$conn.ConnectionString="Server=$server;$dbclause`Integrated Security=True"
	}
	$conn.Open()
    write-debug $conn.ConnectionString
	return $conn
}

function Get-Connection{
<#

.NOTES
    From adolib by Mike Sheppard

#>
[CmdletBinding()]
param([System.Data.SqlClient.SQLConnection]$conn,
      [string]$server, 
      [string]$database,
      [string]$user,
      [string]$password)
	if (-not $conn){
		if ($server){
			$conn=New-Connection -server $server -database $database -user $user -password $password 
		} else {
		    throw "No connection or connection information supplied"
		}
	}
	return $conn
}

function Get-ActiveDBConnections{
<#

.SYNOPSIS

.DESCRIPTION

.PARAMETER SQLDatabase


.NOTES
    

#>
[CmdletBinding()]
param([string]$SQLDatabase)

$Query = @"
SELECT DB_NAME(dbid) AS DBName,
COUNT(dbid) AS NumberOfConnections,
loginame
FROM    sys.sysprocesses
Where DB_ID($($SQLDatabase))
GROUP BY dbid, loginame
ORDER BY DB_NAME(dbid)


"@
Invoke-Sql -connection $SQLConnection -sql $Query




}

function Test-SQLTableExists{
<#
.SYNOPSIS
    Test if SQL table exists

.DESCRIPTION
    Test if SQL table exists

.PARAMETER SQLConnection
    SQL Server Connection Object

.PARAMETER TableName
    Table name to test

.EXAMPLE
    
    Test-SQLTableExists -SQLConnection $SQLConnection -TableName "Table01"

#>
[CmdletBinding()]
	param ( 
		    [Parameter(Mandatory=$True)]$SQLConnection,
            [Parameter(Mandatory=$True)]$TableName  
	)


    # Check if Table is already there
	[String]$Query = "Select Name From sys.tables Where Name = '" + $TableName + "'"

    [String]$message = $TableName + " " + $Query
    Write-Verbose -Message $message

	$QResults = Invoke-Query -connection $SQLConnection -sql $Query

    if($QResults){
        return $true
    }
    else{
        Return $false
    }


} # Test-SQLTableExists

function Remove-SQLTable{
<#
.SYNOPSIS
    Removes a SQL database table

.DESCRIPTION
    Removes a SQL database table

.PARAMETER SQLConnection
    SQL Server Connection Object

.PARAMETER TableName
    Table name to delete

.EXAMPLE
    
    Remove-SQLTable -SQLConnection $SQLConnection -TableName "Table01"

#>
[CmdletBinding()]
	param ( 
		    [Parameter(Mandatory=$True)]$SQLConnection,
            [Parameter(Mandatory=$True)]$SQLCatalog,
            [Parameter(Mandatory=$True)]$TableName  
	)

    $deletestatement = "DROP TABLE [" + $sqlcatalog + "].[dbo].[" + $tablename +  "]"
	Invoke-Sql -connection $SqlConnection -sql $deletestatement | Out-Null 

} # Remove-SQLTable


<#
function Invoke-Query{
<#
.SYNOPSIS
    Get SQL Query results

.DESCRIPTION
    Get SQL Query results


    

#
[CmdletBinding()]
    Param([Parameter(Mandatory=$True)]$Connection,
          [Parameter(Mandatory=$True)]$sql)
     
    $SQLCommand = New-Object System.Data.SqlClient.SqlCommand
    $SQLCommand.Connection = $Connection
    $SQLCommand.CommandText = $sql
    $SQLConnection.Open()
    $results = $SQLCommand.ExecuteReader()
    $table = new-object “System.Data.DataTable”
    $table.Load($results)

    return $table

} # Invoke-Query
#>


function Invoke-Query{
<#
	.SYNOPSIS
		Execute a sql statement, returning the results of the query.  

	.DESCRIPTION
		This function executes a sql statement, using the parameters provided (both input and output) and returns the results of the query.  You may optionally 
        provide a connection or sufficient information to create a connection, as well as input and output parameters, command timeout value, and a transaction to join.

	.PARAMETER  sql
		The SQL Statement

	.PARAMETER  connection
		An existing connection to perform the sql statement with.  

	.PARAMETER  parameters
		A hashtable of input parameters to be supplied with the query.  See example 2. 

	.PARAMETER  outparameters
		A hashtable of input parameters to be supplied with the query.  Entries in the hashtable should have names that match the parameter names, and string values that are the type of the parameters. See example 3. 
        
	.PARAMETER  timeout
		The commandtimeout value (in seconds).  The command will fail and be rolled back if it does not complete before the timeout occurs.

	.PARAMETER  Server
		The server to connect to.  If both Server and Connection are specified, Server is ignored.

	.PARAMETER  Database
		The initial database for the connection.  If both Database and Connection are specified, Database is ignored.

	.PARAMETER  User
		The sql user to use for the connection.  If both User and Connection are specified, User is ignored.

	.PARAMETER  Password
		The password for the sql user named by the User parameter.

	.PARAMETER  Transaction
		A transaction to execute the sql statement in.
    .EXAMPLE
        This is an example of a query that returns a single result.  
        PS C:\> $c=New-Connection '.\sqlexpress'
        PS C:\> $res=invoke-query 'select * from master.dbo.sysdatabases' -conn $c
        PS C:\> $res 
   .EXAMPLE
        This is an example of a query that returns 2 distinct result sets.  
        PS C:\> $c=New-Connection '.\sqlexpress'
        PS C:\> $res=invoke-query 'select * from master.dbo.sysdatabases; select * from master.dbo.sysservers' -conn $c
        PS C:\> $res.Tables[1]
    .EXAMPLE
        This is an example of a query that returns a single result and uses a parameter.  It also generates its own (ad hoc) connection.
        PS C:\> invoke-query 'select * from master.dbo.sysdatabases where name=@dbname' -param  @{dbname='master'} -server '.\sqlexpress' -database 'master'

     .INPUTS
        None.
        You cannot pipe objects to invoke-query

   .OUTPUTS
        Several possibilities (depending on the structure of the query and the presence of output variables)
        1.  A list of rows 
        2.  A dataset (for multi-result set queries)
        3.  An object that contains a dictionary of ouptut parameters and their values and either 1 or 2 (for queries that contain output parameters)

    .NOTES
        From adolib by Mike Sheppard

#>
[CmdletBinding()]
param( [Parameter(Position=0, Mandatory=$true)][string]$sql,
       [Parameter(ParameterSetName="SuppliedConnection", Position=1, Mandatory=$false)][System.Data.SqlClient.SqlConnection]$connection,
       [Parameter(Position=2, Mandatory=$false)][hashtable]$parameters=@{},
       [Parameter(Position=3, Mandatory=$false)][hashtable]$outparameters=@{},
       [Parameter(Position=4, Mandatory=$false)][int]$timeout=120,
       [Parameter(ParameterSetName="AdHocConnection",Position=5, Mandatory=$false)][string]$server,
       [Parameter(ParameterSetName="AdHocConnection",Position=6, Mandatory=$false)][string]$database,
       [Parameter(ParameterSetName="AdHocConnection",Position=7, Mandatory=$false)][string]$user,
       [Parameter(ParameterSetName="AdHocConnection",Position=8, Mandatory=$false)][string]$password,
       [Parameter(Position=9, Mandatory=$false)][System.Data.SqlClient.SqlTransaction]$transaction=$null,
       [Parameter(Position=10, Mandatory=$false)] [ValidateSet("DataSet", "DataTable", "DataRow", "Dynamic")] [string]$AsResult="Dynamic"
       )
    
	$connectionparameters=copy-hashtable $PSBoundParameters -exclude AsResult
    $cmd=new-sqlcommand @connectionparameters
    #$cmd.CommandTimeout = $timeout
    $ds=New-Object system.Data.DataSet
    $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
    $da.fill($ds) | Out-Null
    
    #if it was an ad hoc connection, close it
    if ($server){
       $cmd.connection.close()
    }
    get-outputparameters $cmd $outparameters
    switch ($AsResult)
    {
        'DataSet'   { $result = $ds }
        'DataTable' { $result = $ds.Tables }
        'DataRow'   { $result = $ds.Tables[0] }
        'Dynamic'   { $result = get-commandresults $ds $outparameters } 
    }
    return $result
}



function Add-SqlTable{  
<#  
.SYNOPSIS
Creates a SQL Server table from a DataTable.

.DESCRIPTION
Creates a SQL Server table from a DataTable using SMO. (Based on Chad Miller's original work--see Links section.)

.PARAMETER ServerInstance
Character string or SMO server object specifying the name of an instance of
the Database Engine. For default instances, only specify the computer name:
"MyComputer". For named instances, use the format "ComputerName\InstanceName".
If -ServerInstance is not specified Add-SqlTable attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies the SQL folder,
Add-SqlTable uses the server and instance specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies the SQL folder, Add-SqlTable uses the server and instance
specified in that path.

.PARAMETER Database
A character string specifying the name of a database. Add-SqlTable connects to
this database in the instance specified by -ServerInstance.
If -Database is not specified Add-SqlTable attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies both the SQL folder
and a database name, Add-SqlTable uses the database specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies both the SQL folder and a database name, Add-SqlTable uses
the database specified in that path.

.PARAMETER TableName
Name of the table to create on the SQL Server instance.

.PARAMETER DataTable
A System.Data.DataTable from which the SQL table's schema is derived:
the names and data types of the columns in the DataTable are used
to define the SQL table.

.PARAMETER Username
Specifies the login ID for making a SQL Server Authentication connection
to an instance of the Database Engine. The password must be specified
using -Password. If -Username and -Password are not specified, Add-SqlTable
attempts a Windows Authentication connection using the Windows account running
the PowerShell session.  When possible, use Windows Authentication.  

.PARAMETER Password
Specifies the password for the SQL Server Authentication login ID specified
in -Username. Passwords are case-sensitive.
When possible, use Windows Authentication.
SECURITY NOTE: If you type -Password followed by your password, the password
is visible to anyone who can see your monitor. If you use -Password in
a .ps1 script, anyone reading the script file will see your password.
Assign appropriate permissions to the file to allow only authorized users
to read the file.

.PARAMETER MaxLength
Capacity for VarChar and VarBinary columns (limited to 8000 maximum).
Any data longer than this defined value will be truncated when inserted.

.PARAMETER RowId
Serving as both a switch and a column name, specifying -RowId adds
an identity column to the table with the value supplied as the column name.

.PARAMETER AsScript
If enabled, Add-SqlTable generates a script to create a table
without actually creating it.

.PARAMETER DropExisting
If enabled, Add-SqlTable will check whether the table exists before
attempting to create it. If it does exist, it is first dropped.
Attempting to use Add-SqlTable without this switch when the table already
exists generates an error.

.INPUTS
None. You cannot pipe objects to Add-SqlTable.

.OUTPUTS
None.

.EXAMPLE  
$dt = Invoke-Sqlcmd -ServerInstance "Z003\R2" -Database pubs "select *  from authors"; Add-SqlTable -ServerInstance "Z003\R2" -Database pubscopy -TableName authors -DataTable $dt  
This example loads a variable dt of type DataTable from a query and creates an empty SQL Server table.

.EXAMPLE  
$dt = Get-Alias | Out-DataTable; Add-SqlTable -ServerInstance "Z003\R2" -Database pubscopy -TableName alias -DataTable $dt  
This example creates a DataTable from the properties of Get-Alias and creates an empty SQL Server table.  

.NOTES  
Add-SqlTable uses SQL Server Management Objects (SMO).
SMO is installed with SQL Server Management Studio and is available  
as a separate download:
http://www.microsoft.com/downloads/details.aspx?displaylang=en&FamilyID=ceb4346f-657f-4d28-83f5-aae0c5c83d52  
Version History  
v1.0   - Chad Miller - Initial Release  
v1.1   - Chad Miller - Updated documentation 
v1.2   - Chad Miller - Add loading Microsoft.SqlServer.ConnectionInfo 
v1.3   - Chad Miller - Added error handling 
v1.4   - Chad Miller - Add VarCharMax and VarBinaryMax handling 
v1.5   - Chad Miller - Added AsScript switch to output script instead of creating table 
v1.6   - Chad Miller - Updated Get-SqlType types 

##########################################################################
# Source: http://gallery.technet.microsoft.com/scriptcenter/c193ed1a-9152-4bda-b5c0-acd044e68b2c
# (obsoletes the version at http://poshcode.org/3295)
# Author: Chad Miller
# Date:   2012.07.20
# ** Includes modification by msorens
##########################################################################

.LINK
[Add-SqlTable original by Chad Miller](http://gallery.technet.microsoft.com/scriptcenter/c193ed1a-9152-4bda-b5c0-acd044e68b2c)
.LINK
Out-DataTable
.LINK
Write-DataTable
.LINK
Out-SqlTable
.LINK
Update-DBEnvironment

#> 
    [CmdletBinding()]  
    param(  
    [Parameter(Position=0, Mandatory=$false)] $SQLConnection,
    [Parameter(Position=1, Mandatory=$true)] [String]$TableName,  
    [Parameter(Position=2, Mandatory=$true)] [System.Data.DataTable]$DataTable,  
    [Parameter(Position=3, Mandatory=$false)] [Int32]$MaxLength=1000, 
    [Parameter(Position=4, Mandatory=$false)] [string]$RowId,  
    [Parameter(Position=5, Mandatory=$false)] [switch]$AsScript,
    [Parameter(Position=6, Mandatory=$false)] [switch]$DropExisting
    )  
 
 
try { 
      
  	if ($DropExisting) { 
		
		$dropQuery = "IF OBJECT_ID('{0}') IS NOT NULL DROP TABLE {0}" -f $TableName
        Invoke-Sql -connection $SQLConnection -sql $dropQuery
	}

    $server = new-object ("Microsoft.SqlServer.Management.Smo.Server") $SQLConnection
    
    #$srv = new-Object Microsoft.SqlServer.Management.Smo.Server("(local)")
    $db = New-Object Microsoft.SqlServer.Management.Smo.Database
    $db = $server.Databases.Item(($SQLConnection.Database).ToString())  
    #$db = $SQLConnection.Database  
    $table = new-object Microsoft.SqlServer.Management.Smo.Table($db, $TableName)
    #$table = new-object ("Microsoft.SqlServer.Management.Smo.Table") ($SQLConnection.Database), $TableName  
	
	if ($RowId) {
		$dataType = new-object ("Microsoft.SqlServer.Management.Smo.DataType") "Int"
	    $col = new-object ("Microsoft.SqlServer.Management.Smo.Column") $table, $RowId, $dataType  
        $col.Nullable = $false  
		$col.Identity = $true
        $table.Columns.Add($col)  

	}
	
    foreach ($column in $DataTable.Columns)  
    {  
        $sqlDbType = [Microsoft.SqlServer.Management.Smo.SqlDataType]"$(Get-SqlType $column.DataType.Name)"  
        if ($sqlDbType -eq 'VarBinary' -or $sqlDbType -eq 'VarChar')  
        {  
            if ($MaxLength -gt 0)  
            {$dataType = new-object ("Microsoft.SqlServer.Management.Smo.DataType") $sqlDbType, $MaxLength} 
            else 
            { $sqlDbType  = [Microsoft.SqlServer.Management.Smo.SqlDataType]"$(Get-SqlType $column.DataType.Name)Max" 
              $dataType = new-object ("Microsoft.SqlServer.Management.Smo.DataType") $sqlDbType 
            } 
        }  
        else  
        { $dataType = new-object ("Microsoft.SqlServer.Management.Smo.DataType") $sqlDbType }  
        $col = new-object ("Microsoft.SqlServer.Management.Smo.Column") $table, $column.ColumnName, $dataType  
        $col.Nullable = $column.AllowDBNull
        $table.Columns.Add($col)  
    }  
  
    if ($AsScript) { 
        $table.Script() 
    } 
    else { 
        $table.Create() 
    } 
} 
catch { 
    $message = $_.Exception.GetBaseException().Message 
    Write-Error $message 
} 
   
} #Add-SqlTable
 
function Out-DataTable{
<# 
.SYNOPSIS 
Creates a DataTable for an object.

.DESCRIPTION 
Creates a DataTable based on an object's properties. (Based on Chad Miller's original work--see Links section.)

.INPUTS 
Object. Any object can be piped to Out-DataTable.

.OUTPUTS 
System.Data.DataTable 

.PARAMETER InputObject
Specifies the objects to be converted to a DataTable.
If an array of objects is passed, any with a type that differs 
from the first object in the list are ignored and a warning is displayed.
To suppress the warning, use the TypeFilter parameter to pre-filter the list to a single type. 

.PARAMETER TypeFilter
Specifies a selection filter by data type name.
If not specified, all objects are processed and those with differing types
generate a warning message.
If TypeFilter is specified, only those objects matching the type name are processed.

.EXAMPLE 
$dt = Get-PsDrive | Out-DataTable 
This example creates a DataTable from the properties of Get-PsDrive and assigns output to the $dt variable.

.EXAMPLE
$dt = Get-ChildItem | Out-DataTable -TypeFilter FileInfo
This example creates a DataTable selecting only those objects that have type FileInfo (i.e. ignoring those of type DirectoryInfo, the other object output by Get-ChildItem).

.NOTES 
Adapted from script by Marc van Orsouw see link 
Version History 
v1.0  - Chad Miller - Initial Release 
v1.1  - Chad Miller - Fixed Issue with Properties 
v1.2  - Chad Miller - Added setting column datatype by property as suggested by emp0 
v1.3  - Chad Miller - Corrected issue with setting datatype on empty properties 
v1.4  - Chad Miller - Corrected issue with DBNull 
v1.5  - Chad Miller - Updated example 
v1.6  - Chad Miller - Added column datatype logic with default to string 

##########################################################################
# Source: http://gallery.technet.microsoft.com/scriptcenter/4208a159-a52e-4b99-83d4-8048468d29dd
# (obsoletes the version at http://poshcode.org/2954)
# Author: Chad Miller
# Date:   2012.07.20
# ** includes modification by msorens
##########################################################################

.LINK 
[PowerShell Guy's original](http://thepowershellguy.com/blogs/posh/archive/2007/01/21/powershell-gui-scripblock-monitor-script.aspx)
[Out-DataTable original by Chad Miller based on PowerGuy's](http://gallery.technet.microsoft.com/scriptcenter/4208a159-a52e-4b99-83d4-8048468d29dd)
Add-SqlTable
Write-DataTable
Out-SqlTable
#> 
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject,
		[string]$TypeFilter
	)
 
    Begin 
    { 
        $dt = new-object Data.DataTable   
        $First = $true  
		$count = 0
    } 
    Process 
    { 
        foreach ($object in $InputObject) 
        { 
			# 2012.10.11 msorens: filter to silence warnings, too.
			if ($TypeFilter -and $object.GetType().Name -ne $TypeFilter) { continue }
			
			# 2012.10.11 msorens: warn about different types instead of throwing exception
			$count++
			if ($First) { $firstObjectType = $object.GetType() }
			elseif ( $object.GetType() -ne $firstObjectType) {
				Write-Warning ("Skipping {0}th object (type={1}, expected={2})" `
				-f $count, $object.GetType(), $firstObjectType)
				continue
			}
            $DR = $DT.NewRow()   
            foreach($property in $object.PsObject.get_properties()) 
            {   
                if ($first) 
                {   
                    $Col =  new-object Data.DataColumn   
                    $Col.ColumnName = $property.Name.ToString()   
                    # 2012.10.11 msorens: Modified test to allow zero to pass; 
					# otherwise, zero in the first record prevents data type assignment for the column. 
					$valueExists = Get-Member -InputObject $property -Name value
					if ($valueExists)
                    { 
						# 2012.10.11 msorens: Modified test for nulls to also include $null
                        if ($property.value -isnot [System.DBNull] -and $property.value -ne $null) {
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
                         } 
                    } 
                    $DT.Columns.Add($Col) 
                }
                # 2012.10.11 msorens: Changed from .IsArray because, when present, was null;
				# other times caused error (property 'IsArray' not found...).
                if ($property.Value -is [array]) {
                    $DR.Item($property.Name) = $property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                }   
                # 2012.10.11 msorens: Added support for XML fields
                elseif ($property.Value -is [System.Xml.XmlElement]) {
					$DR.Item($property.Name) = $property.Value.OuterXml
				}
                else { 
                    $DR.Item($property.Name) = $property.value 
                } 
            }   
            $DT.Rows.Add($DR)   
            $First = $false 
        } 
    }  
      
    End 
    { 
        Write-Output @(,($dt)) 
    } 
 
} # Out-DataTable 

function Out-SqlTable{
<#  
.SYNOPSIS  
Creates a SQL Server table from Powershell objects.

.DESCRIPTION 
Out-SqlTable  is simply a composition of Out-DataTable, Add-SqlTable, and Write-DataTable.

.PARAMETER TableName
Name of the table to create on the SQL Server instance.

.PARAMETER ServerInstance
Character string or SMO server object specifying the name of an instance of
the Database Engine. For default instances, only specify the computer name:
"MyComputer". For named instances, use the format "ComputerName\InstanceName".
If -ServerInstance is not specified Out-SqlTable attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies the SQL folder,
Out-SqlTable uses the server and instance specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies the SQL folder, Out-SqlTable uses the server and instance
specified in that path.

.PARAMETER Database
A character string specifying the name of a database. Out-SqlTable connects to
this database in the instance specified by -ServerInstance.
If -Database is not specified Out-SqlTable attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies both the SQL folder
and a database name, Out-SqlTable uses the database specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies both the SQL folder and a database name, Out-SqlTable uses
the database specified in that path.

.PARAMETER Username
Specifies the login ID for making a SQL Server Authentication connection
to an instance of the Database Engine. The password must be specified
using -Password. If -Username and -Password are not specified, Out-SqlTable
attempts a Windows Authentication connection using the Windows account running
the PowerShell session.  When possible, use Windows Authentication.  

.PARAMETER Password
Specifies the password for the SQL Server Authentication login ID specified
in -Username. Passwords are case-sensitive.
When possible, use Windows Authentication.
SECURITY NOTE: If you type -Password followed by your password, the password
is visible to anyone who can see your monitor. If you use -Password in
a .ps1 script, anyone reading the script file will see your password.
Assign appropriate permissions to the file to allow only authorized users
to read the file.

.PARAMETER MaxLength
Capacity for VarChar and VarBinary columns (limited to 8000 maximum).
Any data longer than this defined value will be truncated when inserted.

.PARAMETER RowId
Serving as both a switch and a column name, specifying -RowId adds
an identity column to the table with the value supplied as the column name.

.PARAMETER DropExisting
If enabled, Out-SqlTable will check whether the table exists before
attempting to create it. If it does exist, it is first dropped.
Attempting to use Out-SqlTable without this switch when the table already
exists generates an error.

.PARAMETER BatchSize
Number of rows to send to server at one time (set to 0 to send all rows).
This parameter maps to the same named parameter in the .NET class SqlBulkCopy. 

.PARAMETER QueryTimeout
Number of seconds for batch to complete before failing. Though not advised,
you can use 0 to indicate no limit.
This parameter maps to the BulkCopyTimeout parameter in the .NET class SqlBulkCopy. 

.PARAMETER ConnectionTimeout
Number of seconds for connection to complete before failing. Though not advised,
you can use 0 to indicate no limit.

.INPUTS
Object. Any object can be piped to Out-SqlTable.

.OUTPUTS
None. Produces no output.

.EXAMPLE  
Get-Process | select ProcessName, Handle | Out-SqlTable -TableName "processes"
Puts selected columnar output from Get-Process into a table in the current database, where the "current database" is specifed by the current location on a SQLSERVER: drive.  This is equivalent to these three consecutive commands:
    $dt = Get-Process | select ProcessName, Handle | Out-DataTable
    Add-SqlTable -TableName "processes" -DataTable $dt
    Write-DataTable -TableName "processes" -Data $dt

.EXAMPLE  
ps | select ProcessName, Handle | Out-SqlTable -TableName "processes" -DropExisting -RowId "MyId"
Puts selected columnar output from Get-Process into a table in the current database, dropping the table first if it exists, and adding an identity column "MyId".

.EXAMPLE  
Get-SvnLog . | Out-SqlTable -TableName "svndata" -MaxLength 3000
Puts the output of the CleanCode Get-SvnLog cmdlet into a database table, allowing string fields to be up to 3000 characters rather than the default 1000.  This is useful primarily for the msg property of Subversion data.

.NOTES
This function is part of the CleanCode toolbox
from http://cleancode.sourceforge.net/.

Since CleanCode 1.1.05.

#
# ==============================================================
# @ID       $Id: Out-SqlTable.ps1 1395 2013-06-05 02:13:32Z ms $
# @created  2012-11-01
# @project  http://cleancode.sourceforge.net/
# ==============================================================
#
# The official license for this file is shown next.
# Unofficially, consider this e-postcardware as well:
# if you find this module useful, let us know via e-mail, along with
# where in the world you are and (if applicable) your website address.
#
#
# ***** BEGIN LICENSE BLOCK *****
# Version: MPL 1.1
#
# The contents of this file are subject to the Mozilla Public License Version
# 1.1 (the "License"); you may not use this file except in compliance with
# the License. You may obtain a copy of the License at
# http://www.mozilla.org/MPL/
#
# Software distributed under the License is distributed on an "AS IS" basis,
# WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
# for the specific language governing rights and limitations under the
# License.
#
# The Original Code is part of the CleanCode toolbox.
#
# The Initial Developer of the Original Code is Michael Sorens.
# Portions created by the Initial Developer are Copyright (C) 2012
# the Initial Developer. All Rights Reserved.
#
# Contributor(s):
#
# ***** END LICENSE BLOCK *****
#

.LINK
Out-DataTable
.LINK
Add-SqlTable
.LINK
Write-DataTable
.LINK
Update-DBEnvironment
.LINK
[SqlBulkCopy](http://msdn.microsoft.com/en-us/library/30c3y597.aspx)

#>  
[CmdletBinding()] 
param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)] [PSObject[]]$InputObject,
	[Parameter(Position=1, Mandatory=$true)] $SQLConnection,
    [Parameter(Position=2, Mandatory=$true)] [string]$TableName,
    [Parameter(Mandatory=$false)] [Int32]$MaxLength, 
    [Parameter(Mandatory=$false)] [string]$RowId,  
    [Parameter(Mandatory=$false)] [switch]$DropExisting,
    [Parameter(Mandatory=$false)] [Int32]$BatchSize,
    [Parameter(Mandatory=$false)] [Int32]$QueryTimeout,
    [Parameter(Mandatory=$false)] [Int32]$ConnectionTimeout
)
 
End {
	
	$dataTable = $InputObject | Out-DataTable

	Add-SqlTable -SQLConnection $SQLConnection -TableName $TableName -DataTable $dataTable
	Write-DataTable -SQLConnection $SQLConnection -TableName $TableName -Data $dataTable
	}
} # Out-SqlTable
 
function Write-DataTable{
<# 
.SYNOPSIS 
Writes data to a SQL Server table.

.DESCRIPTION 
Writes data only to SQL Server tables.  However, the data source is not limited to SQL Server; any data source can be used, as long as the data can be loaded to a DataTable instance or read with an IDataReader instance. (Based on Chad Miller's original work--see Links section.)

.INPUTS 
None. You cannot pipe objects to Write-DataTable.

.OUTPUTS 
None. Produces no output.

.PARAMETER ServerInstance
Character string or SMO server object specifying the name of an instance of
the Database Engine. For default instances, only specify the computer name:
"MyComputer". For named instances, use the format "ComputerName\InstanceName".
If -ServerInstance is not specified Write-DataTable attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies the SQL folder,
Write-DataTable uses the server and instance specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies the SQL folder, Write-DataTable uses the server and instance
specified in that path.

.PARAMETER Database
A character string specifying the name of a database. Write-DataTable connects to
this database in the instance specified by -ServerInstance.
If -Database is not specified Write-DataTable attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies both the SQL folder
and a database name, Write-DataTable uses the database specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies both the SQL folder and a database name, Write-DataTable uses
the database specified in that path.

.PARAMETER TableName
Name of the table to create on the SQL Server instance.

.PARAMETER Data
A System.Data.DataTable containing the data to write to the SQL table.

.PARAMETER Username
Specifies the login ID for making a SQL Server Authentication connection
to an instance of the Database Engine. The password must be specified
using -Password. If -Username and -Password are not specified, Write-DataTable
attempts a Windows Authentication connection using the Windows account running
the PowerShell session.  When possible, use Windows Authentication.  

.PARAMETER Password
Specifies the password for the SQL Server Authentication login ID specified
in -Username. Passwords are case-sensitive.
When possible, use Windows Authentication.
SECURITY NOTE: If you type -Password followed by your password, the password
is visible to anyone who can see your monitor. If you use -Password in
a .ps1 script, anyone reading the script file will see your password.
Assign appropriate permissions to the file to allow only authorized users
to read the file.

.PARAMETER BatchSize
Number of rows to send to server at one time (set to 0 to send all rows).
This parameter maps to the same named parameter in the .NET class SqlBulkCopy. 

.PARAMETER QueryTimeout
Number of seconds for batch to complete before failing. Though not advised,
you can use 0 to indicate no limit.
This parameter maps to the BulkCopyTimeout parameter in the .NET class SqlBulkCopy. 

.PARAMETER ConnectionTimeout
Number of seconds for connection to complete before failing. Though not advised,
you can use 0 to indicate no limit.

.EXAMPLE 
$dt = Invoke-Sqlcmd -ServerInstance "Z003\R2" -Database pubs "select *  from authors"; Write-DataTable -ServerInstance "Z003\R2" -Database pubscopy -TableName authors -Data $dt 
This example loads a variable dt of type DataTable from query and write the datatable to another database.  This requires the target table to already exist.  Use Add-SqlTable to create the table if necessary.

.NOTES 
Write-DataTable uses the SqlBulkCopy class see links for additional information on this class. 
Version History 
v1.0   - Chad Miller - Initial release 
v1.1   - Chad Miller - Fixed error message 

##########################################################################
# Source: http://gallery.technet.microsoft.com/scriptcenter/2fdeaf8d-b164-411c-9483-99413d6053ae
# (obsoletes the version at http://poshcode.org/2943)
# Author: Chad Miller
# Date:   2010.10.04
##########################################################################

.LINK 
[Write-DataTable original by Chad Miller](http://gallery.technet.microsoft.com/scriptcenter/2fdeaf8d-b164-411c-9483-99413d6053ae)
.LINK
Out-DataTable
.LINK
Add-SqlTable
.LINK
Out-SqlTable
.LINK
Update-DBEnvironment
.LINK
[SqlBulkCopy](http://msdn.microsoft.com/en-us/library/30c3y597.aspx)

#> 
    [CmdletBinding()] 
    param( 
    [Parameter(Position=0, Mandatory=$true)] $SQLConnection, 
    [Parameter(Position=2, Mandatory=$true)] [string]$TableName, 
    [Parameter(Position=3, Mandatory=$true)] $Data, 
    [Parameter(Position=6, Mandatory=$false)] [Int32]$BatchSize=0, # same as SqlBulkCopy default
    [Parameter(Position=7, Mandatory=$false)] [Int32]$QueryTimeout=30, # same as SqlBulkCopy default
    [Parameter(Position=8, Mandatory=$false)] [Int32]$ConnectionTimeout=15 # same as SqlConnection default 
    ) 
  
    try 
    { 
        $bulkCopy = new-object ("Data.SqlClient.SqlBulkCopy") $SQLConnection 
        $bulkCopy.DestinationTableName = "[" + $tableName + "]"
        $bulkCopy.BatchSize = $BatchSize 
        $bulkCopy.BulkCopyTimeout = $QueryTimeOut


        <#
        # 2012.11.15 msorens: Allow for identity columns to exist
        $ServerInstance = "MSSQL"
        $Database = $SQLConnection.Database.ToString()
        $tableObj = Get-Item SQLSERVER:\sql\$ServerInstance\Databases\$Database\tables\dbo.$tableName
        $identCount = @($tableObj.Columns | ? { $_.Identity}).Count
        if ($identCount -gt 0) {
            foreach ($col in $Data.Columns) {
                $colmap = new-object ("Data.SqlClient.SqlBulkCopyColumnMapping") $col.ColumnName, $col.ColumnName
                [Void] $bulkCopy.ColumnMappings.Add($colmap)
            }
        }
        #>

        $bulkCopy.WriteToServer($Data) 
    }
    catch 
    { 
        $ex = $_.Exception 
        Write-Error "$ex.Message" 
        continue 
    } 
 
} #Write-DataTable

function Update-DBEnvironment{
<#  
.SYNOPSIS  
Infer values for server and database into reference variables.

.DESCRIPTION
Updates the supplied server instance and the database with inferred values if the supplied values are undefined.  It extends the notion of default context. That is, you do not actually have to be on the SQLSERVER: drive, only to have set the current location on that drive to a server and database path.

.INPUTS
None. You cannot pipe objects to Update-DBEnvironment.

.OUTPUTS
None. Produces no output.

.EXAMPLE
Update-DBEnvironment ([ref]$ServerInstance) ([ref]$Database)
If either of the supplied server or database values are $null or empty string, Get-DBPathInfo is used to attempt to infer a values for it, updating the supplied variables by reference.

.PARAMETER ServerInstance
Reference to a character string specifying the name of an instance of
the Database Engine. For default instances, only specify the computer name:
"MyComputer". For named instances, use the format "ComputerName\InstanceName".
If -ServerInstance is not specified Update-DBEnvironment attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies the SQL folder,
Update-DBEnvironment uses the server and instance specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies the SQL folder, Update-DBEnvironment uses the server and instance
specified in that path.

.PARAMETER Database
Reference to a character string specifying the name of a database.
Update-DBEnvironment connects to this database in the instance specified by -ServerInstance.
If -Database is not specified Update-DBEnvironment attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies both the SQL folder
and a database name, Update-DBEnvironment uses the database specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies both the SQL folder and a database name, Update-DBEnvironment uses
the database specified in that path.

.NOTES
This function is part of the CleanCode toolbox
from http://cleancode.sourceforge.net/.

Since CleanCode 1.1.05.

#
# ==============================================================
# @ID       $Id: DefaultContext.ps1 1391 2013-05-31 03:00:02Z ms $
# @created  2012-11-01
# @project  http://cleancode.sourceforge.net/
# ==============================================================
#
# The official license for this file is shown next.
# Unofficially, consider this e-postcardware as well:
# if you find this module useful, let us know via e-mail, along with
# where in the world you are and (if applicable) your website address.
#
#
# ***** BEGIN LICENSE BLOCK *****
# Version: MPL 1.1
#
# The contents of this file are subject to the Mozilla Public License Version
# 1.1 (the "License"); you may not use this file except in compliance with
# the License. You may obtain a copy of the License at
# http://www.mozilla.org/MPL/
#
# Software distributed under the License is distributed on an "AS IS" basis,
# WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
# for the specific language governing rights and limitations under the
# License.
#
# The Original Code is part of the CleanCode toolbox.
#
# The Initial Developer of the Original Code is Michael Sorens.
# Portions created by the Initial Developer are Copyright (C) 2012
# the Initial Developer. All Rights Reserved.
#
# Contributor(s):
#
# ***** END LICENSE BLOCK *****
#
#>
[CmdletBinding()]  
    param([ref] $ServerInstance, [ref] $Database)
    if (!$ServerInstance.Value -or !$Database.Value) {
        ($sCandidate, $dCandidate) = Get-DBPathInfo
        if (!$ServerInstance.Value) { $ServerInstance.Value = $sCandidate }
        if (!$Database.Value) { $Database.Value = $dCandidate }
        if (!$ServerInstance.Value -or !$Database.Value) {
            throw "ServerInstance and Database must be specified"
        }
    }
} # Update-DBEnvironment

function Get-DBPathInfo{
<#  
.SYNOPSIS  
Infer values for server and database from SQLSERVER: drive.

.DESCRIPTION
Infers the server instance and database from the current location or, if not on a SQL Server drive, then from the default SQLSERVER: drive.  It extends the notion of default context. That is, you do not actually have to be on the SQLSERVER: drive, only to have set the current location on that drive to a server and database path.

.INPUTS
None. You cannot pipe objects to Get-DBPathInfo.

.OUTPUTS
Two-element array: first element contains server and instance name; second element contains database name.

.EXAMPLE
($serverCandidate, $dbCandidate) = Get-DBPathInfo
If your current location is *not* on a SQLSERVER: drive or alias, then the current location of the SQLSERVER: drive is used as the reference location.  If your current location includes the SQLSERVER: drive or an alias to some node on it, that is the reference location.  Either way, the reference location is examined to see if it contains a server and/or database.  If either server or db are both are not found $null is returned for that item.

.NOTES
This function is part of the CleanCode toolbox
from http://cleancode.sourceforge.net/.

Since CleanCode 1.1.05.

#
# ==============================================================
# @ID       $Id: DefaultContext.ps1 1391 2013-05-31 03:00:02Z ms $
# @created  2012-11-01
# @project  http://cleancode.sourceforge.net/
# ==============================================================
#
# The official license for this file is shown next.
# Unofficially, consider this e-postcardware as well:
# if you find this module useful, let us know via e-mail, along with
# where in the world you are and (if applicable) your website address.
#
#
# ***** BEGIN LICENSE BLOCK *****
# Version: MPL 1.1
#
# The contents of this file are subject to the Mozilla Public License Version
# 1.1 (the "License"); you may not use this file except in compliance with
# the License. You may obtain a copy of the License at
# http://www.mozilla.org/MPL/
#
# Software distributed under the License is distributed on an "AS IS" basis,
# WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
# for the specific language governing rights and limitations under the
# License.
#
# The Original Code is part of the CleanCode toolbox.
#
# The Initial Developer of the Original Code is Michael Sorens.
# Portions created by the Initial Developer are Copyright (C) 2012
# the Initial Developer. All Rights Reserved.
#
# Contributor(s):
#
# ***** END LICENSE BLOCK *****
#
#>
    [CmdletBinding()]  
    param()
    if ((Get-Location).Drive.Provider.Name -ne "SqlServer") {
        $sqlPath = Get-PSDrive | ? { $_.Name -eq "SQLSERVER" } | % { $_.Root + $_.CurrentLocation }
    }
    else {
        $sqlPath = (Get-Item .).PSPath
    }
    if ($sqlPath -match 'SQLSERVER:\\sql\\([^\\]*\\[^\\]*)\\Databases\\([^\\]*)')
    { return $Matches[1],$Matches[2] }
    elseif ($sqlPath -match 'SQLSERVER:\\sql\\([^\\]*\\[^\\]*)')
    { return $Matches[1],$null }
    else { return $null, $null }
} # Get-DBPathInfo

function Invoke-InferredSqlcmd{
<#  
.SYNOPSIS  
Executes Invoke-Sqlcmd while attempting to supply a default context.

.DESCRIPTION
This cmdlet is exactly equivalent to Invoke-Sqlcmd except that it extends the notion of default context. That is, you do not actually have to be on the SQLSERVER: drive, only to have set the current location on that drive to a server and database path.

.INPUTS
None. You cannot pipe objects to Invoke-InferredSqlcmd.

.OUTPUTS
Formatted table.

.EXAMPLE
Invoke-InferredSqlcmd -Query "SELECT GETDATE() AS TimeOfQuery;" -ServerInstance "MyComputer\MyInstance" -Database "MyDatabase"
Runs the basic T-SQL query on the explicitly provided server and database.

.EXAMPLE
Invoke-InferredSqlcmd -Query "SELECT GETDATE() AS TimeOfQuery;" 
Runs the basic T-SQL query using inferred values for ServerInstance and Database.

.PARAMETER ServerInstance
A character string specifying the name of an instance of
the Database Engine. For default instances, only specify the computer name:
"MyComputer". For named instances, use the format "ComputerName\InstanceName".
If -ServerInstance is not specified Invoke-InferredSqlcmd attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies the SQL folder,
Invoke-InferredSqlcmd uses the server and instance specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies the SQL folder, Invoke-InferredSqlcmd uses the server and instance
specified in that path.
This latter inference is the difference between Invoke-Sqlcmd
and Invoke-InferredSqlcmd.

.PARAMETER Database
A character string specifying the name of a database.
Invoke-InferredSqlcmd connects to this database in the instance specified by -ServerInstance.
If -Database is not specified Invoke-InferredSqlcmd attempts to infer it:
* If on a SQLSERVER: drive (or alias) and the path specifies both the SQL folder
and a database name, Invoke-InferredSqlcmd uses the database specified in the path.
* If not on a SQLSERVER: drive but the current location on your SQLSERVER:
drive specifies both the SQL folder and a database name, Invoke-InferredSqlcmd uses
the database specified in that path.
This latter inference is the difference between Invoke-Sqlcmd
and Invoke-InferredSqlcmd.

.PARAMETER EncryptConnection
See Invoke-Sqlcmd.

.PARAMETER Username
See Invoke-Sqlcmd.

.PARAMETER Password
See Invoke-Sqlcmd.

.PARAMETER Query
See Invoke-Sqlcmd.

.PARAMETER QueryTimeout
See Invoke-Sqlcmd.

.PARAMETER ConnectionTimeout
See Invoke-Sqlcmd.

.PARAMETER ErrorLevel
See Invoke-Sqlcmd.

.PARAMETER SeverityLevel
See Invoke-Sqlcmd.

.PARAMETER MaxCharLength
See Invoke-Sqlcmd.

.PARAMETER MaxBinaryLength
See Invoke-Sqlcmd.

.PARAMETER AbortOnError
See Invoke-Sqlcmd.

.PARAMETER DedicatedAdministratorConnection
See Invoke-Sqlcmd.

.PARAMETER DisableVariables
See Invoke-Sqlcmd.

.PARAMETER DisableCommands
See Invoke-Sqlcmd.

.PARAMETER HostName
See Invoke-Sqlcmd.

.PARAMETER NewPassword
See Invoke-Sqlcmd.

.PARAMETER Variable
See Invoke-Sqlcmd.

.PARAMETER InputFile
See Invoke-Sqlcmd.

.PARAMETER OutputSqlErrors
See Invoke-Sqlcmd.

.PARAMETER SuppressProviderContextWarning
See Invoke-Sqlcmd.

.PARAMETER IgnoreProviderContext
See Invoke-Sqlcmd.

.NOTES
This function is part of the CleanCode toolbox
from http://cleancode.sourceforge.net/.

Since CleanCode 1.1.05.

#
# ==============================================================
# @ID       $Id: DefaultContext.ps1 1391 2013-05-31 03:00:02Z ms $
# @created  2012-11-01
# @project  http://cleancode.sourceforge.net/
# ==============================================================
#
# The official license for this file is shown next.
# Unofficially, consider this e-postcardware as well:
# if you find this module useful, let us know via e-mail, along with
# where in the world you are and (if applicable) your website address.
#
#
# ***** BEGIN LICENSE BLOCK *****
# Version: MPL 1.1
#
# The contents of this file are subject to the Mozilla Public License Version
# 1.1 (the "License"); you may not use this file except in compliance with
# the License. You may obtain a copy of the License at
# http://www.mozilla.org/MPL/
#
# Software distributed under the License is distributed on an "AS IS" basis,
# WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
# for the specific language governing rights and limitations under the
# License.
#
# The Original Code is part of the CleanCode toolbox.
#
# The Initial Developer of the Original Code is Michael Sorens.
# Portions created by the Initial Developer are Copyright (C) 2012
# the Initial Developer. All Rights Reserved.
#
# Contributor(s):
#
# ***** END LICENSE BLOCK *****
#
#>
    [CmdletBinding()]  
    param(  
    [string]$ServerInstance,
    [string]$Database,
    [switch]$EncryptConnection,
    [string]$Username,
    [string]$Password,
    [Parameter(Position=0)] [string]$Query,
    [int]$QueryTimeout,
    [int]$ConnectionTimeout,
    [int]$ErrorLevel,
    [int]$SeverityLevel,
    [int]$MaxCharLength,
    [int]$MaxBinaryLength,
    [switch]$AbortOnError,
    [switch]$DedicatedAdministratorConnection,
    [switch]$DisableVariables,
    [switch]$DisableCommands,
    [string]$HostName,
    [string]$NewPassword,
    [string[]]$Variable,
    [string]$InputFile,
    [switch]$OutputSqlErrors,
    [switch]$SuppressProviderContextWarning,
    [switch]$IgnoreProviderContext
    )
    
    # Infer the DB server details and push values back into the param list
    Update-DBEnvironment ([ref]$ServerInstance) ([ref]$Database)
    $PSBoundParameters["ServerInstance"] = $ServerInstance
    $PSBoundParameters["Database"] = $Database

    # This is our raison d'etre so unless the caller has requested
    # to see the warning, suppress it.
    if (!$PSBoundParameters.ContainsKey("SuppressProviderContextWarning")) {
        $PSBoundParameters["SuppressProviderContextWarning"] = $true
    }
    
    Invoke-Sqlcmd @PSBoundParameters
} # Invoke-InferredSqlcmd




function Set-DBOffline{
<#
.Synopsis
    Set SQL database in  offline state

.Description
    Set SQL database in  offline state
     
.Parameter srv
    SQL Server Name
    
.Parameter dbname
    SQL Database Name

.Example
    Set-DBOffline -srv "SQL01" -dbname "Test01"

.Notes
    
#Requires -Version 4.0

#>
[CmdletBinding()]
Param([Parameter(Mandatory=$False)][String]$SQLServer="(local)",[Parameter(Mandatory=$True)][String]$dbname)

    $srv = new-Object Microsoft.SqlServer.Management.Smo.Server($SQLServer)
    $db = New-Object Microsoft.SqlServer.Management.Smo.Database
    $db = $srv.Databases.Item($dbname)
    $db.SetOffline()

} # Set-DBOffline


function Set-DBOnline{
<#
.Synopsis
    Set SQL database in  online state

.Description
    Set SQL database in  online state
     
.Parameter srv
    SQL Server Name
    
.Parameter dbname
    SQL Database Name

.Example
    Set-DBOnline -srv "SQL01" -dbname "Test01"

.Notes
    
#Requires -Version 4.0

#>
[CmdletBinding()]
Param([Parameter(Mandatory=$False)][String]$SQLServer="(local)",[Parameter(Mandatory=$True)][String]$dbname)

    $srv = new-Object Microsoft.SqlServer.Management.Smo.Server($SQLServer)
    $db = New-Object Microsoft.SqlServer.Management.Smo.Database
    $db = $srv.Databases.Item($dbname)
    $db.SetOnline()

} # Set-DBOnline


function Set-DBAttach{
<#
.Synopsis
    Attach SQL database

.Description
    Attach SQL database
     
.Parameter srv
    SQL Server Name
    
.Parameter dbname
    SQL Database Name

.Example
    Set-DBAttach -srv "SQL01" -dbname "Test01" -datapath "C:\MSSQL\Data\data01.mdf" -logpath "C:\MSSQL\Logs\log01.ldf"

.Notes

    Needs to accept arrays for paths to accomodate multiple files 
    
#Requires -Version 4.0

#>
[CmdletBinding()]
Param([Parameter(Mandatory=$False)][String]$SQLServer="(local)",[Parameter(Mandatory=$True)][String]$dbname,[Parameter(Mandatory=$True)][String]$datapath,[Parameter(Mandatory=$True)][String]$logpath)


    Try{

        $srv = new-Object Microsoft.SqlServer.Management.Smo.Server($SQLServer)
        $db = New-Object Microsoft.SqlServer.Management.Smo.Database
        $db = $srv.Databases.Item($dbname)
        $sc = new-object System.collections.specialized.stringcollection 
        $sc.Add($datapath)
        $sc.Add($logpath)
        $srv.AttachDatabase($dbname,$sc,[Microsoft.SqlServer.Management.Smo.AttachOptions]::None)

        $True

    }Catch{
        $False
    }
    

} # Set-DBAttach


function Set-DBDetach{
<#
.Synopsis
    Detach SQL database

.Description
    Detach SQL database
     
.Parameter srv
    SQL Server Name
    
.Parameter dbname
    SQL Database Name

.Example
    Set-DBDetach -srv "SQL01" -dbname "Test01"

.Notes
    
#Requires -Version 4.0

#>
[CmdletBinding()]
Param([Parameter(Mandatory=$False)][String]$SQLServer="(local)",[Parameter(Mandatory=$True)][String]$dbname)

    Try{
        $srv = new-Object Microsoft.SqlServer.Management.Smo.Server($SQLServer)
        $srv.DetachDatabase($dbname,$False,$False)

        $True
    }Catch{

        $False
    }
    
    

} # Set-DBDetach

				





function New-OutputParameters{
<#

.NOTES
    From adolib by Mike Sheppard

#>
[CmdletBinding()]
param([Parameter(Position=0, Mandatory=$true)][System.Data.SqlClient.SQLCommand]$cmd, 
      [Parameter(Position=1, Mandatory=$false)][hashtable]$outparams)
    if ($outparams){
    	foreach($outp in $outparams.Keys){
            $paramtype=get-paramtype $outparams[$outp]
            $p=$cmd.Parameters.Add("@$outp",$paramtype)
    		$p.Direction=[System.Data.ParameterDirection]::Output
            if ($paramtype -like '*char*'){
               $p.Size=[int]$outparams[$outp].Replace($paramtype.ToString().ToLower(),'').Replace('(','').Replace(')','')
            }
    	}
    }
} # New-OutputParameters

function Get-Outputparameters{
<#

.NOTES
    From adolib by Mike Sheppard

#>
[CmdletBinding()]
param([Parameter(Position=0, Mandatory=$true)][System.Data.SqlClient.SQLCommand]$cmd,
      [Parameter(Position=1, Mandatory=$true)][hashtable]$outparams)
	foreach($p in $cmd.Parameters){
		if ($p.Direction -eq [System.Data.ParameterDirection]::Output){
		  $outparams[$p.ParameterName.Replace("@","")]=$p.Value
		}
	}
}

function Get-ParamType{
<#

.NOTES
    From adolib by Mike Sheppard

#>
[CmdletBinding()]
param([string]$typename)
	$type=switch -wildcard ($typename.ToLower()) {
		'uniqueidentifier' {[System.Data.SqlDbType]::UniqueIdentifier}
		'int'  {[System.Data.SQLDbType]::Int}
		'datetime'  {[System.Data.SQLDbType]::Datetime}
		'tinyint'  {[System.Data.SQLDbType]::tinyInt}
		'smallint'  {[System.Data.SQLDbType]::smallInt}
		'bigint'  {[System.Data.SQLDbType]::BigInt}
		'bit'  {[System.Data.SQLDbType]::Bit}
		'char*'  {[System.Data.SQLDbType]::char}
		'nchar*'  {[System.Data.SQLDbType]::nchar}
		'date'  {[System.Data.SQLDbType]::date}
		'datetime'  {[System.Data.SQLDbType]::datetime}
        'varchar*' {[System.Data.SqlDbType]::Varchar}
        'nvarchar*' {[System.Data.SqlDbType]::nVarchar}
		default {[System.Data.SqlDbType]::Int}
	}
	return $type
	
}

function Copy-HashTable{
<#

.NOTES
    From adolib by Mike Sheppard

#>
[CmdletBinding()]
param([hashtable]$hash,
[String[]]$include,
[String[]]$exclude)

	if($include){
	   $newhash=@{}
	   foreach ($key in $include){
	    if ($hash.ContainsKey($key)){
	   		$newhash.Add($key,$hash[$key]) | Out-Null 
		}
	   }
	} else {
	   $newhash=$hash.Clone()
	   if ($exclude){
		   foreach ($key in $exclude){
		        if ($newhash.ContainsKey($key)) {
		   			$newhash.Remove($key) | Out-Null 
				}
		   }
	   }
	}
	return $newhash
}

function Get-CommandResults{
<#

.NOTES
    From adolib by Mike Sheppard


    Helper function figure out what kind of returned object to build from the results of a sql call (ds). 
    Options are:
	    1.  Dataset   (multiple lists of rows)
	    2.  Datatable (list of datarows)
	    3.  Nothing (no rows and no output variables
	    4.  Dataset with output parameter dictionary
	    5.  Datatable with output parameter dictionary
	    6.  A dictionary of output parameters

#>
[CmdletBinding()]
param([Parameter(Position=0, Mandatory=$true)][System.Data.Dataset]$ds, 
      [Parameter(Position=1, Mandatory=$true)][HashTable]$outparams)   

	if ($ds.tables.count -eq 1){
		$retval= $ds.Tables[0]
	}
	elseif ($ds.tables.count -eq 0){
		$retval=$null
	} else {
		[system.Data.DataSet]$retval= $ds 
	}
	if ($outparams.Count -gt 0){
		if ($retval){
			return @{Results=$retval; OutputParameters=$outparams}
		} else {
			return $outparams
		}
	} else {
		return $retval
	}
}

function New-SQLCommand{
<#
	.SYNOPSIS
		Create a sql command object

	.DESCRIPTION
		This function uses the information contained in the parameters to create a sql command object.  In general, you will want to use the invoke- functions directly, 
        but if you need to manipulate a command object in ways that those functions don't allow, you will need this.  Also, the invoke-bulkcopy function allows you to pass 
        a command object instead of a set of records in order to "stream" the records into the destination in cases where there are a lot of records and you don't want to
        allocate memory to hold the entire result set.

	.PARAMETER  sql
		The sql to be executed by the command object (although it is not executed by this function).

	.PARAMETER  connection
		An existing connection to perform the sql statement with.  

	.PARAMETER  parameters
		A hashtable of input parameters to be supplied with the query.  See example 2. 
        
	.PARAMETER  timeout
		The commandtimeout value (in seconds).  The command will fail and be rolled back if it does not complete before the timeout occurs.

	.PARAMETER  Server
		The server to connect to.  If both Server and Connection are specified, Server is ignored.

	.PARAMETER  Database
		The initial database for the connection.  If both Database and Connection are specified, Database is ignored.

	.PARAMETER  User
		The sql user to use for the connection.  If both User and Connection are specified, User is ignored.

	.PARAMETER  Password
		The password for the sql user named by the User parameter.

	.PARAMETER  Transaction
		A transaction to execute the sql statement in.

	.EXAMPLE
		PS C:\> $cmd=new-sqlcommand "ALTER DATABASE AdventureWorks Modify Name = Northwind" -server MyServer
        PS C:\> $cmd.ExecuteNonQuery()


	.EXAMPLE
		PS C:\> $cmd=new-sqlcommand -server MyServer -sql "Select * from MyTable"
        PS C:\> invoke-sqlbulkcopy -records $cmd -server MyOtherServer -table CopyOfMyTable

    .INPUTS
        None.
        You cannot pipe objects to new-sqlcommand

	.OUTPUTS
		System.Data.SqlClient.SqlCommand

    .NOTES
        From adolib by Mike Sheppard


#>
[CmdletBinding()]
param([Parameter(Position=0, Mandatory=$true)][Alias('storedProcName')][string]$sql,
      [Parameter(ParameterSetName="SuppliedConnection",Position=1, Mandatory=$false)][System.Data.SqlClient.SQLConnection]$connection,
      [Parameter(Position=2, Mandatory=$false)][hashtable]$parameters=@{},
      [Parameter(Position=3, Mandatory=$false)][int]$timeout=30,
      [Parameter(ParameterSetName="AdHocConnection",Position=4, Mandatory=$false)][string]$server,
      [Parameter(ParameterSetName="AdHocConnection",Position=5, Mandatory=$false)][string]$database,
      [Parameter(ParameterSetName="AdHocConnection",Position=6, Mandatory=$false)][string]$user,
      [Parameter(Position=7, Mandatory=$false)][string]$password,
      [Parameter(Position=8, Mandatory=$false)][System.Data.SqlClient.SqlTransaction]$transaction=$null,
	  [Parameter(Position=9, Mandatory=$false)][hashtable]$outparameters=@{})
   
    $dbconn=Get-Connection -conn $connection -server $server -database $database -user $user -password $password
    $close=($dbconn.State -eq [System.Data.ConnectionState]'Closed')
    if ($close) {
        $dbconn.Open()
    }	
    $cmd=new-object system.Data.SqlClient.SqlCommand($sql,$dbconn)
    $cmd.CommandTimeout=$timeout
    foreach($p in $parameters.Keys){
	    $parm=$cmd.Parameters.AddWithValue("@$p",$parameters[$p])
        if (Test-IsDBNull $parameters[$p]){
           $parm.Value=[DBNull]::Value
        }
    }
    New-OutputParameters $cmd $outparameters

    if ($transaction -is [System.Data.SqlClient.SqlTransaction]){
	$cmd.Transaction = $transaction
    }
    return $cmd


}

function Invoke-Sql{
<#
	.SYNOPSIS
		Execute a sql statement, ignoring the result set.  Returns the number of rows modified by the statement (or -1 if it was not a DML staement)

	.DESCRIPTION
		This function executes a sql statement, using the parameters provided and returns the number of rows modified by the statement.  You may optionally 
        provide a connection or sufficient information to create a connection, as well as input parameters, command timeout value, and a transaction to join.

	.PARAMETER  sql
		The SQL Statement

	.PARAMETER  connection
		An existing connection to perform the sql statement with.  

	.PARAMETER  parameters
		A hashtable of input parameters to be supplied with the query.  See example 2. 
        
	.PARAMETER  timeout
		The commandtimeout value (in seconds).  The command will fail and be rolled back if it does not complete before the timeout occurs.

	.PARAMETER  Server
		The server to connect to.  If both Server and Connection are specified, Server is ignored.

	.PARAMETER  Database
		The initial database for the connection.  If both Database and Connection are specified, Database is ignored.

	.PARAMETER  User
		The sql user to use for the connection.  If both User and Connection are specified, User is ignored.

	.PARAMETER  Password
		The password for the sql user named by the User parameter.

	.PARAMETER  Transaction
		A transaction to execute the sql statement in.

	.EXAMPLE
		PS C:\> invoke-sql "ALTER DATABASE AdventureWorks Modify Name = Northwind" -server MyServer


	.EXAMPLE
		PS C:\> $con=New-Connection MyServer
        PS C:\> invoke-sql "Update Table1 set Col1=null where TableID=@ID" -parameters @{ID=5}

    .INPUTS
        None.
        You cannot pipe objects to invoke-sql

	.OUTPUTS
		Integer

    .NOTES
        From adolib by Mike Sheppard

#>
[CmdletBinding()]
param([Parameter(Position=0, Mandatory=$true)][string]$sql,
      [Parameter(ParameterSetName="SuppliedConnection",Position=1, Mandatory=$false)][System.Data.SqlClient.SQLConnection]$connection,
      [Parameter(Position=2, Mandatory=$false)][hashtable]$parameters=@{},
      [Parameter(Position=3, Mandatory=$false)][hashtable]$outparameters=@{},
      [Parameter(Position=4, Mandatory=$false)][int]$timeout=30,
      [Parameter(ParameterSetName="AdHocConnection",Position=5, Mandatory=$false)][string]$server,
      [Parameter(ParameterSetName="AdHocConnection",Position=6, Mandatory=$false)][string]$database,
      [Parameter(ParameterSetName="AdHocConnection",Position=7, Mandatory=$false)][string]$user,
      [Parameter(ParameterSetName="AdHocConnection",Position=8, Mandatory=$false)][string]$password,
      [Parameter(Position=9, Mandatory=$false)][System.Data.SqlClient.SqlTransaction]$transaction=$null)
	

       $cmd=new-sqlcommand @PSBoundParameters

       #if it was an ad hoc connection, close it
       if ($server){
          $cmd.connection.close()
       }	

       return $cmd.ExecuteNonQuery()
	
}

function Invoke-StoredProcedure{
<#
	.SYNOPSIS
		Execute a stored procedure, returning the results of the query.  

	.DESCRIPTION
		This function executes a stored procedure, using the parameters provided (both input and output) and returns the results of the query.  You may optionally 
        provide a connection or sufficient information to create a connection, as well as input and output parameters, command timeout value, and a transaction to join.

	.PARAMETER  sql
		The SQL Statement

	.PARAMETER  connection
		An existing connection to perform the sql statement with.  

	.PARAMETER  parameters
		A hashtable of input parameters to be supplied with the query.  See example 2. 

	.PARAMETER  outparameters
		A hashtable of input parameters to be supplied with the query.  Entries in the hashtable should have names that match the parameter names, and string values that are the type of the parameters. 
        Note:  not all types are accounted for by the code. int, uniqueidentifier, varchar(n), and char(n) should all work, though.
        
	.PARAMETER  timeout
		The commandtimeout value (in seconds).  The command will fail and be rolled back if it does not complete before the timeout occurs.

	.PARAMETER  Server
		The server to connect to.  If both Server and Connection are specified, Server is ignored.

	.PARAMETER  Database
		The initial database for the connection.  If both Database and Connection are specified, Database is ignored.

	.PARAMETER  User
		The sql user to use for the connection.  If both User and Connection are specified, User is ignored.

	.PARAMETER  Password
		The password for the sql user named by the User parameter.

	.PARAMETER  Transaction
		A transaction to execute the sql statement in.
    .EXAMPLE
        #Calling a simple stored procedure with no parameters
        PS C:\> $c=New-Connection -server '.\sqlexpress' 
        PS C:\> invoke-storedprocedure 'sp_who2' -conn $c
    .EXAMPLE 
        #Calling a stored procedure that has an output parameter and multiple result sets
        PS C:\> $c=New-Connection '.\sqlexpress'
        PS C:\> $res=invoke-storedprocedure -storedProcName 'AdventureWorks2008.dbo.stp_test' -outparameters @{LogID='int'} -conne $c
        PS C:\> $res.Results.Tables[1]
        PS C:\> $res.OutputParameters
        
        For reference, here's the stored procedure:
        CREATE procedure [dbo].[stp_test]
            @LogID int output
        as
            set @LogID=5
            select * from master.dbo.sysdatabases
            select * from master.dbo.sysservers
    .EXAMPLE 
        #Calling a stored procedure that has an input parameter
        PS C:\> invoke-storedprocedure 'sp_who2' -conn $c -parameters @{loginame='sa'}
    .INPUTS
        None.
        You cannot pipe objects to invoke-storedprocedure

    .OUTPUTS
        Several possibilities (depending on the structure of the query and the presence of output variables)
        1.  A list of rows 
        2.  A dataset (for multi-result set queries)
        3.  An object that contains a hashtables of ouptut parameters and their values and either 1 or 2 (for queries that contain output parameters)

    .NOTES
        From adolib by Mike Sheppard
    
#>
[CmdletBinding()]
param([Parameter(Position=0, Mandatory=$true)][string]$storedProcName,
      [Parameter(ParameterSetName="SuppliedConnection",Position=1, Mandatory=$false)][System.Data.SqlClient.SqlConnection]$connection,
      [Parameter(Position=2, Mandatory=$false)][hashtable] $parameters=@{},
      [Parameter(Position=3, Mandatory=$false)][hashtable]$outparameters=@{},
      [Parameter(Position=4, Mandatory=$false)][int]$timeout=30,
      [Parameter(ParameterSetName="AdHocConnection",Position=5, Mandatory=$false)][string]$server,
      [Parameter(ParameterSetName="AdHocConnection",Position=6, Mandatory=$false)][string]$database,
      [Parameter(ParameterSetName="AdHocConnection",Position=7, Mandatory=$false)][string]$user,
      [Parameter(ParameterSetName="AdHocConnection",Position=8, Mandatory=$false)][string]$password,
      [Parameter(Position=9, Mandatory=$false)][System.Data.SqlClient.SqlTransaction]$transaction=$null) 

	$cmd=new-sqlcommand @PSBoundParameters
	$cmd.CommandType=[System.Data.CommandType]::StoredProcedure  
    $ds=New-Object system.Data.DataSet
    $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
    $da.fill($ds) | out-null

    get-outputparameters $cmd $outparameters

    #if it was an ad hoc connection, close it
    if ($server){
       $cmd.connection.close()
    }	
	
    return (get-commandresults $ds $outparameters)
}

function Invoke-Bulkcopy{
<#
	.SYNOPSIS
		Uses the .NET SQLBulkCopy class to quickly copy rows into a destination table.

	.DESCRIPTION
        
		Also, the invoke-bulkcopy function allows you to pass a command object instead of a set of records in order to "stream" the records
        into the destination in cases where there are a lot of records and you don't want to allocate memory to hold the entire result set.

	.PARAMETER  records
		Either a datatable (like one returned from invoke-query or invoke-storedprocedure) or
        A command object (e.g. new-sqlcommand), or a datareader object.  Note that the command object or datareader object 
        can come from any class that inherits from System.Data.Common.DbCommand or System.Data.Common.DataReader, so this will work
        with most ADO.NET client libraries (not just SQL Server).

	.PARAMETER  Server
		The destination server to connect to.  

	.PARAMETER  Database
		The initial database for the connection.  

	.PARAMETER  User
		The sql user to use for the connection.  If user is not passed, NT Authentication is used.

	.PARAMETER  Password
		The password for the sql user named by the User parameter.

	.PARAMETER  Table
		The destination table for the bulk copy operation.

	.PARAMETER  Mapping
		A dictionary of column mappings of the form DestColumn=SourceColumn

	.PARAMETER  BatchSize
		The batch size for the bulk copy operation.

	.PARAMETER  Transaction
		A transaction to execute the bulk copy operation in.

	.PARAMETER  NotifyAfter
		The number of rows to fire the notification event after transferring.  0 means don't notify.
        Ex: 1000 means to fire the notify event after each 1000 rows are transferred.
        
    .PARAMETER  NotifyFunction
        A scriptblock to be executed after each $notifyAfter records has been copied.  The second parameter ($param[1]) 
        is a SqlRowsCopiedEventArgs object, which has a RowsCopied property.  The default value for this parameter echoes the
        number of rows copied to the console
        
    .PARAMETER  Options
        An object containing special options to modify the bulk copy operation.
        See http://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqlbulkcopyoptions.aspx for values.


	.EXAMPLE
		PS C:\> $cmd=new-sqlcommand -server MyServer -sql "Select * from MyTable"
        PS C:\> invoke-sqlbulkcopy -records $cmd -server MyOtherServer -table CopyOfMyTable

	.EXAMPLE
		PS C:\> $rows=invoke-query -server MyServer -sql "Select * from MyTable"
        PS C:\> invoke-sqlbulkcopy -records $rows -server MyOtherServer -table CopyOfMyTable

    .INPUTS
        None.
        You cannot pipe objects to invoke-bulkcopy

	.OUTPUTS
		System.Data.SqlClient.SqlCommand

    .NOTES
        From adolib by Mike Sheppard

#>
[CmdletBinding()]
  param([Parameter(Position=0, Mandatory=$true)]$records,
        [Parameter(Position=1, Mandatory=$true)]$server,
        [Parameter(Position=2, Mandatory=$false)]$database,
        [Parameter(Position=3, Mandatory=$false)][string]$user,
        [Parameter(Position=4, Mandatory=$false)][string]$password,
        [Parameter(Position=5, Mandatory=$true)][string]$table,
        [Parameter(Position=6, Mandatory=$false)]$mapping=@{},
        [Parameter(Position=7, Mandatory=$false)]$batchsize=0,
        [Parameter(Position=8, Mandatory=$false)][System.Data.SqlClient.SqlTransaction]$transaction=$null,
        [Parameter(Position=9, Mandatory=$false)]$notifyAfter=0,
        [Parameter(Position=10, Mandatory=$false)][scriptblock]$notifyFunction={Write-Host "$($args[1].RowsCopied) rows copied."},
        [Parameter(Position=11, Mandatory=$false)][System.Data.SqlClient.SqlBulkCopyOptions]$options=[System.Data.SqlClient.SqlBulkCopyOptions]::Default)

	#use existing "New-Connection" function to create a connection string.        
    $connection=New-Connection -server $server -database $Database -User $user -password $password
	$connectionString = $connection.ConnectionString
	$connection.close()

	#Use a transaction if one was specified
	if ($transaction -is [System.Data.SqlClient.SqlTransaction]){
		$bulkCopy=new-object "Data.SqlClient.SqlBulkCopy" $connectionString $options  $transaction
	} else {
		$bulkCopy = new-object "Data.SqlClient.SqlBulkCopy" $connectionString
	}
	$bulkCopy.BatchSize=$batchSize
	$bulkCopy.DestinationTableName = $table
	$bulkCopy.BulkCopyTimeout=10000000
	if ($notifyAfter -gt 0){
		$bulkCopy.NotifyAfter=$notifyafter
		$bulkCopy.Add_SQlRowscopied($notifyFunction)
	}

	#Add column mappings if they were supplied
	foreach ($key in $mapping.Keys){
	    $bulkCopy.ColumnMappings.Add($mapping[$key],$key) | out-null
	}
	
	write-debug "Bulk copy starting at $(get-date)"
	if ($records -is [System.Data.Common.DBCommand]){
		#if passed a command object (rather than a datatable), ask it for a datareader to stream the records
		$bulkCopy.WriteToServer($records.ExecuteReader())
    } elsif ($records -is [System.Data.Common.DbDataReader]){
		#if passed a Datareader object use it to stream the records
		$bulkCopy.WriteToServer($records)
	} else {
		$bulkCopy.WriteToServer($records)
	}
	write-debug "Bulk copy finished at $(get-date)"
}




Export-ModuleMember Get-SQLConnection
Export-ModuleMember Test-SQLTableExists
Export-ModuleMember Remove-SQLTable
Export-ModuleMember Out-DataTable
Export-ModuleMember Out-SqlTable
Export-ModuleMember Add-SqlTable
Export-ModuleMember Write-DataTable
Export-ModuleMember Get-DBPathInfo
Export-ModuleMember Update-DBEnvironment
Export-ModuleMember Invoke-InferredSqlcmd

Export-ModuleMember Set-DBAttach
Export-ModuleMember Set-DBDetach
Export-ModuleMember Set-DBOffline
Export-ModuleMember Set-DBOnline

export-modulemember New-Connection
export-modulemember new-sqlcommand
export-modulemember invoke-sql
export-modulemember invoke-query
export-modulemember invoke-storedprocedure
export-modulemember invoke-bulkcopy