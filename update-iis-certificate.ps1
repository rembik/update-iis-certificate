<#

  Syntax:

    Passing parameters:
      update-iis-certificate.ps1 [".\pkcs12.pfx"] -CertDomain "example.com" [-PFXPassword "P@ssw0rd"] [-SiteName "Default Web Site"] [-Port 443] [-ExcludeLocalServerCert] 

    All parameters in square brackets are optional.
    The ExcludeLocalServerCert is forced to $True if left off. 
    You really never want this set to false, especially
    if using a wildcard certificate. It's there mainly for flexibility.

    If the password contains a $ sign, you must escape it with the `
    character.

  Script Name: update-iis-certificate.ps1
  Release:     1.0
  Written by   Jeremy@jhouseconsulting.com 21st December 2014
               (http://www.jhouseconsulting.com/2015/01/04/script-to-import-and-bind-a-certificate-to-the-default-web-site-1548)

  Modified by  brian@rimek.info 18th May 2017
               (https://github.com/rembik/update-iis-certificate)

  Note:        This script has been tested thoroughly on Windows 2012R2
               (IIS 8.5). Due to the cmdlets used I cannot guarantee full
               backward compatibility.

  A log file will either be written to %windir%\Temp or to the
  %LogPath% Task Sequence variable if running from an SCCM\MDT
  Task.

#>

#-------------------------------------------------------------
param (
  [Parameter(Position = 0)][String]$PFXPath,
  [String]$CertDomain=$(throw "Parameter CertDomain is required, please provide a value! e.g. -CertDomain 'example.com'"),
  [String]$PFXPassword,
  [String]$SiteName,
  [int]$Port,
  [switch]$ExcludeLocalServerCert
)

# Set Powershell Compatibility Mode
Set-StrictMode -Version 2.0

$ScriptPath = {Split-Path $MyInvocation.ScriptName}

if ([String]::IsNullOrEmpty($PFXPath)) {
  $PFXPath = $(&$ScriptPath) + "\pkcs12.pfx"
}

if ([String]::IsNullOrEmpty($PFXPassword)) {
  $PFXPassword = ""
}

if ([String]::IsNullOrEmpty($SiteName)) {
  $SiteName = "Default Web Site"
}

if ([int]::IsNullOrEmpty($Port)) {
  $Port = 443
}

if (!($ExcludeLocalServerCert.IsPresent)) { 
  $ExcludeLocalServerCert = $True
}

#-------------------------------------------------------------

Function IsTaskSequence() {
  # This code was taken from a discussion on the CodePlex PowerShell
  # App Deployment Toolkit site. It was posted by mmashwani.
  Try {
      [__ComObject]$SMSTSEnvironment = New-Object -ComObject Microsoft.SMS.TSEnvironment -ErrorAction 'SilentlyContinue' -ErrorVariable SMSTSEnvironmentErr
  }
  Catch {
  }
  If ($SMSTSEnvironmentErr) {
    Write-Verbose "Unable to load ComObject [Microsoft.SMS.TSEnvironment]. Therefore, script is not currently running from an MDT or SCCM Task Sequence."
    Return $false
  }
  ElseIf ($null -ne $SMSTSEnvironment) {
    Write-Verbose "Successfully loaded ComObject [Microsoft.SMS.TSEnvironment]. Therefore, script is currently running from an MDT or SCCM Task Sequence."
    Return $true
  }
}

#-------------------------------------------------------------

$invalidChars = [io.path]::GetInvalidFileNamechars() 
$datestampforfilename = ((Get-Date -format s).ToString() -replace "[$invalidChars]","-")

# Get the script path
$ScriptName = [System.IO.Path]::GetFilenameWithoutExtension($MyInvocation.MyCommand.Path.ToString())
$Logfile = "$ScriptName-$($datestampforfilename).txt"
$logPath = "$($env:windir)\Temp"

If (IsTaskSequence) {
  $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment 
  $logPath = $tsenv.Value("LogPath")

  $UserDomain = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tsenv.Value("UserDomain")))
  $UserID = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tsenv.Value("UserID")))
  $UserPassword = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($tsenv.Value("UserPassword")))
}

$logfile = "$logPath\$Logfile"

# Start the logging 
Start-Transcript $logFile
Write-Output "Logging to $logFile"

#-------------------------------------------------------------

Write-Output "Start Certificate Installation/Update"

Write-Output "Loading the Web Administration Module"
try{
    Import-Module webadministration
}
catch{
    Write-Output "Failed to load the Web Administration Module"
}

Write-Output "Locating the old cert in the Store"
If ($ExcludeLocalServerCert) {
    $oldCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertDomain*" -AND $_.subject -notmatch "CN=$env:COMPUTERNAME"}
} Else {
    $oldCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertDomain*"}
}
If ($oldCert) {
    $oldThumbprint = $oldCert.Thumbprint.ToString()
    Write-Output $oldCert
} Else {
    $oldThumbprint = ""
    Write-Output "Unable to locate current cert in certificate store"
}

Write-Output "Running certutil to import new certificate into Store"
try{
    $ImportSucceed = $True
    $ImportError = certutil.exe -f -importpfx -p $PFXPassword My $PFXPath
}
catch{
    $ImportSucceed = $False
    Write-Output "certutil failed to import new certificate: $ImportError"
}

Write-Output "Locating the new cert in the Store"
if ($ImportSucceed) {
    try{
        If ($ExcludeLocalServerCert) {
            $newCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertDomain*" -AND $_.thumbprint -notmatch $oldThumbprint -AND $_.subject -notmatch "CN=$env:COMPUTERNAME"}
        } Else {
            $newCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertDomain*" -AND $_.thumbprint -notmatch $oldThumbprint}
        }
        $newThumbprint = $newCert.Thumbprint.ToString()
        Write-Output $newCert
    }
    catch{
        Write-Output "Unable to locate new cert in certificate store"
        }

    Write-Output "Deleting old certificate from Store"
    try{
        If (Test-Path "cert:\LocalMachine\My\$oldThumbprint") {
          Remove-Item -Path cert:\LocalMachine\My\$oldThumbprint -DeleteKey
        }
    }
    catch{
        Write-Output "Unable to delete old certificate from store"
    }

    Write-Output "Removing any existing binding from the site and SSLBindings store"
    try{
      # Remove existing binding form site  
      if ($null -ne (Get-WebBinding -Name $SiteName | where-object {$_.protocol -eq "https"})) {
        $RemoveWebBinding = Remove-WebBinding -Name $SiteName -Port $Port -Protocol "https"
        Write-Output $RemoveWebBinding
      }
      # Remove existing binding in SSLBindings store
      If (Test-Path "IIS:\SslBindings\0.0.0.0!$Port") {
        $RemoveSSLBinding = Remove-Item -path "IIS:\SSLBindings\0.0.0.0!$Port"
        Write-Output $RemoveSSLBinding
      }
    }
    catch{
        Write-Output "Unable to remove existing binding"
    }

    Write-Output "Bind your certificate to IIS HTTPS listener"
    try{
      $NewWebBinding = New-WebBinding -Name $SiteName -Port $Port -Protocol "https"
      Write-Output $NewWebBinding
      $AddSSLCertToWebBinding = (Get-WebBinding $SiteName -Port $Port -Protocol "https").AddSslCertificate($newThumbprint, "My")
      Write-Output $AddSSLCertToWebBinding
    }
    catch{
        Write-Output "Unable to bind cert"
    }
}

Write-Output "Completed Certificate Installation/Update"
 
# Stop logging 
Stop-Transcript