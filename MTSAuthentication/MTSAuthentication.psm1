<# Module Name:     MTSAuthentication.psm1
##
## Author:          David Muegge
## Purpose:         Provides PowerShell functions for various authentication operations
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




function Disable-CertificateValidation{
<#
.SYNOPSIS
    Disable certificate validation

.DESCRIPTION
    Ignore SSL errors - This would not be used if self signed certificates were not used and the proper CA cert was installed

.EXAMPLE
    Disable-CertificateValidation

.NOTES
    

#>

[CmdletBinding()]

param()

add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;

    public class IDontCarePolicy : ICertificatePolicy {
    public IDontCarePolicy() {}
    public bool CheckValidationResult(
        ServicePoint sPoint, X509Certificate cert,
        WebRequest wRequest, int certProb) {
        return true;
    }
}
"@
[System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy 

} # Disable-CertificateValidation

function Set-IngnoreCertificateWarnings{
 <#
.Synopsis
    Ignores certificate warnings

.Description
    Ignores certificate warnings     
    
.Example
    Set-IngnoreCertificateWarnings
    
#Requires -Version 2.0

#>
    #region Choose to ignore any SSL Warning issues caused by Self Signed Certificates    
    ## Code From http://poshcode.org/624  
    ## Create a compilation environment  
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider  
    $Compiler=$Provider.CreateCompiler()  
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters  
    $Params.GenerateExecutable=$False  
    $Params.GenerateInMemory=$True  
    $Params.IncludeDebugInformation=$False  
    $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null  
  
$TASource=@' 
    namespace Local.ToolkitExtensions.Net.CertificatePolicy{ 
    public class TrustAll : System.Net.ICertificatePolicy { 
        public TrustAll() {  
        } 
        public bool CheckValidationResult(System.Net.ServicePoint sp, 
        System.Security.Cryptography.X509Certificates.X509Certificate cert,  
        System.Net.WebRequest req, int problem) { 
        return true; 
        } 
    } 
    } 
'@   
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)  
    $TAAssembly=$TAResults.CompiledAssembly  
  
    ## We now create an instance of the TrustAll and attach it to the ServicePointManager  
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")  
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll  
  
    #endregion end code from http://poshcode.org/624
 } # Set-IngnoreCertificateWarnings


function New-PasswordFile{
<#
.SYNOPSIS
    Creates enrypted password file

.DESCRIPTION
    Will prompt for password and write encrypted password to file
    Encryption key is generated based on current windows user security token

.PARAMETER Path
    Path to location of encrypted password file

.PARAMETER Filename
    Filename of enrypted password file

.INPUTS
    Filename and path
    Will prompt for password to be encrypted

.OUTPUTS
    Encrypted password file

.EXAMPLE
    New-PasswordFile -Path "C:\Temp\Passwords" -Filename "lab-dmuegge-IS7TEST01-PWD.txt"

.NOTES
    This cmdlet is used allow the use of basic authentication and persist authentication info without the use of cookies

#>

	[CmdletBinding()]
	param ( 
		[Parameter(Mandatory=$True)][string]$Path,
		[Parameter(Mandatory=$True)][string]$Filename,
        [Parameter(Mandatory=$False)][string]$Password=$null
	)

    
            
    If(Test-Path $Path){
        if($Password){
            $FullFilePath = $Path + "\" + $Filename
            New-Item -Path $FullFilePath -ItemType File
	        $passwd = ConvertTo-SecureString -AsPlainText $Password -Force
            ConvertFrom-SecureString -securestring $passwd | Out-File -FilePath $FullFilePath.ToString()
        }else{

	        $passwd = Read-Host -prompt "Password" -assecurestring
            $FullFilePath = $Path + "\" + $Filename
            New-Item -Path $FullFilePath -ItemType File
	        ConvertFrom-SecureString -securestring $passwd | Out-File -FilePath $FullFilePath.ToString()
        }
    }
    Else
    {
        Write-Error "[New-PasswordFile] :: Path file not found: " $Path
    }

    


} # New-PasswordFile

function Get-PasswordFromFile{
<#
.SYNOPSIS
   Get password from encrypted password file 

.DESCRIPTION
    Will prompt for password and write encrypted password to file
    Encryption key is generated based on current windows user security token

.PARAMETER Path
    Path to location of encrypted password file
        Required?                    true 
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false

.PARAMETER Filename
    Filename of enrypted password file
        Required?                    true 
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false

.INPUTS
    Filename and path
    
.OUTPUTS
    Encrypted password file

.EXAMPLE
    New-PasswordFile -Path "C:\Temp\Passwords" -Filename "lab-dmuegge-IS7TEST01-PWD.txt"

.NOTES
    This cmdlet is used allow the use of basic authentication and persist authentication info without the use of cookies

#>

	[CmdletBinding()]
	
	param ( 
		[Parameter(Mandatory=$True)][string]$FullPath,
        [Parameter(Mandatory=$False)][Switch]$AsSecureString,
        [Parameter(Mandatory=$False)][Switch]$AsPlainText
	)

    Try{

        # Test file existence and retrieve file object
        If(Test-Path -Path $FullPath){

            $File = Get-item -Path $FullPath
            $filedata = Get-Content -Path $File.FullName
            $password = ConvertTo-SecureString $filedata

            If($AsSecureString){return $password}

            If($AsPlainText){

                $BSTR = [System.Runtime.InteropServices.marshal]::SecureStringToBSTR($password)
                $password = [System.Runtime.InteropServices.marshal]::PtrToStringAuto($BSTR)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
                return $password

            }

        }
        else
        {
                
            Write-Error "[Get-PasswordFromFile] :: Password file not found: " $FullPath

        }
                        

    }
    Catch{

        Write-Verbose "[Get-PasswordFromFile] :: threw an exception: $_"
        Write-Error $_

    }

    
			
} # Get-PasswordFromFile


function Set-WindowsAuthenticationCredential{
<#
.Synopsis
   Return windows credential object

.DESCRIPTION
   Return windows credential object

.PARAMETER userid
    userid[String]
        Required?                    true 
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false

.PARAMETER password
    password[SecureString]
        Required?                    true 
        Position?                    named
        Default value                
        Accept pipeline input?       false
        Accept wildcard characters?  false

.EXAMPLE
   $Cred = Set-WindowsAuthenticationCredential -userid "presidioad\dmuegge" -password (Get-PasswordFromFile -FullPath $passwordfile -AsSecureString )


.NOTES
   

#>
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true)][String]$userid,
        [Parameter(Mandatory=$true)][SecureString]$password
                
    )

 
        
    new-object -typename System.Management.Automation.PSCredential -argumentlist $userid,$password

 
} # Set-WindowsAuthenticationCredential



Export-ModuleMember Disable-CertificateValidation
Export-ModuleMember Set-IngnoreCertificateWarnings
Export-ModuleMember New-PasswordFile
Export-ModuleMember Get-PasswordFromFile
Export-ModuleMember Set-WindowsAuthenticationCredential