<#

  Syntax examples:
    
    Install/update certificate in certificate store and webbindings for "Default Web Site" in IIS:
      update-iis-certificate.ps1 ".\example.com.pfx" -CertSubject "example.com" -PFXPassword "P@ssw0rd"

    Remove certificate from certificate store and webbindings for "Default Web Site" from IIS:
      update-iis-certificate.ps1 -CertSubject "example.com" -Remove
                                 
    Passing parameters:
      update-iis-certificate.ps1 [[-PFXPath] <String>] -CertSubject <String> [-PFXPassword <String> ]
                                 [-SiteName <String> ][-IP <String> ][-Port <int> ][-HostHeader <String> ]
                                 [-SNI][-Remove][-ExcludeLocalServerCert]
      
    All parameters in square brackets are optional.
    Most of them are for customized webbindings, see:
    https://technet.microsoft.com/de-de/library/hh867854(v=wps.630).aspx
    The ExcludeLocalServerCert is forced to $True if left off. 
    You really never want this set to false, especially
    if using a wildcard certificate. It's there mainly for flexibility.

    If the password contains a $ sign, you must escape it with the `
    character.

  Script Name: update-iis-certificate.ps1
  Release:     1.0
  Written by   Jeremy@jhouseconsulting.com 21st December 2014
               http://www.jhouseconsulting.com/2015/01/04/script-to-import-and-bind-a-certificate-to-the-default-web-site-1548

  Modified by  brian@rimek.info 18th May 2017
               https://github.com/rembik/update-iis-certificate

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
  [String]$CertSubject=$(throw "Parameter CertSubject is required, please provide a value! e.g. -CertSubject 'example.com'"),
  [String]$PFXPassword,
  [String]$SiteName,
  [String]$HostHeader,
  [String]$IP,
  [String]$Port,
  [switch]$SNI,
  [switch]$Remove,
  [switch]$ExcludeLocalServerCert
)

# Set Powershell Compatibility Mode
Set-StrictMode -Version 2.0

$ScriptPath = {Split-Path $MyInvocation.ScriptName}

if ([String]::IsNullOrEmpty($PFXPath)) {
  $PFXPath = $(&$ScriptPath) + "\pkcs12.pfx"
}

if ([String]::IsNullOrEmpty($PFXPassword)) {
  $secPFXPassword = ""
} else {
  $secPFXPassword = ConvertTo-SecureString -String $PFXPassword -Force -AsPlainText
}

if ([String]::IsNullOrEmpty($SiteName)) {
  $SiteName = "Default Web Site"
}

if ([String]::IsNullOrEmpty($HostHeader)) {
  $HostHeader = $False
}

if ([String]::IsNullOrEmpty($IP)) {
  $IP = "*"
}

if ([String]::IsNullOrEmpty($Port)) {
  [int]$Port = 443
} else {
  [int]$Port = [convert]::ToInt32($Port, 10)
}

if (!($SNI.IsPresent)) { 
  [int]$SNI = 0
} else {
  [int]$SNI = 1
}

if ($Remove.IsPresent) { 
  $Remove = $True
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
Write-Output "Logging to $($logFile):"

#-------------------------------------------------------------

Write-Output " +++ Start Certificate Update +++ "

Write-Output " + Loading the Web Administration Module..."
try{
    Import-Module webadministration
}
catch{
    Write-Output " + Failed to load the Web Administration Module!"
}

Write-Output " + Locating the current(old) certificate in store..."
If ($ExcludeLocalServerCert) {
    $oldCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.subject -notmatch "CN=$env:COMPUTERNAME"}
} Else {
    $oldCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*"}
}
If ($oldCert) {
    $oldThumbprint = $oldCert.Thumbprint.ToString()
    Write-Output " + Current(old) certificate:"
    Write-Output $oldCert
} Else {
    $oldThumbprint = ""
    Write-Output " + Unable to locate current(old) certificate in store!"
}


$ImportSucceed = $False
If (!$Remove){
    Write-Output " + Importing certificate into store..."
    try{
        $ImportOutput = Import-PfxCertificate –FilePath $PFXPath -CertStoreLocation "cert:\LocalMachine\My" -Exportable -Password $secPFXPassword -ErrorAction Stop -ErrorVariable ImportError
        $ImportSucceed = $True
        Write-Output " + Imported certificate:"
        write-Output $ImportOutput
    }
    catch{
        Write-Output " + Failed to import certificate: $ImportError"
    }
}

If ($ImportSucceed -OR $Remove) {
    Write-Output " + Locating the new certificate in store..."
    try{
        If ($ExcludeLocalServerCert) {
            $newCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.thumbprint -ne $oldThumbprint -AND $_.subject -notmatch "CN=$env:COMPUTERNAME"}
        } Else {
            $newCert = Get-ChildItem cert:\LocalMachine\My | Where-Object {$_.subject -like "CN=$CertSubject*" -AND $_.thumbprint -ne $oldThumbprint}
        }
        $newThumbprint = $newCert.Thumbprint.ToString()
        Write-Output " + New certificate:"
        Write-Output $newCert
    }
    catch{
        Write-Output " + Unable to locate new certificate in store!"
        }

    If ($newCert -OR $Remove) {
        Write-Output " + Removing any existing binding from the site and SSLBindings store..."
        try{
          If ($HostHeader -ne $False){
            # Remove existing binding form site  
            If ($null -ne (Get-WebBinding $SiteName -HostHeader $HostHeader -IPAddress $IP -Port $Port -Protocol "https")) {
                $RemoveWebBinding = Remove-WebBinding -Name $SiteName -HostHeader $HostHeader -IPAddress $IP -Port $Port -Protocol "https"
                Write-Output $RemoveWebBinding
            }
            # Remove existing binding in SSLBindings store
            If (Test-Path "IIS:\SslBindings\$IP!$Port!$HostHeader") {
                $RemoveSSLBinding = Remove-Item -path "IIS:\SslBindings\$IP!$Port!$HostHeader"
                Write-Output $RemoveSSLBinding
            }
          } Else { 
            if ($null -ne (Get-WebBinding $SiteName -IPAddress $IP -Port $Port -Protocol "https")) {
                $RemoveWebBinding = Remove-WebBinding -Name $SiteName -IPAddress $IP -Port $Port -Protocol "https"
                Write-Output $RemoveWebBinding
            }
            If (Test-Path "IIS:\SslBindings\$IP!$Port") {
                $RemoveSSLBinding = Remove-Item -path "IIS:\SslBindings\$IP!$Port"
                Write-Output $RemoveSSLBinding
            }
          }
        }
        catch{
            Write-Output " + Unable to remove existing binding!"
        }
    }

    If ($newCert) {
        Write-Output " + Bind your certificate to IIS HTTPS listener..."
        try{
          If ($HostHeader -ne $False){
            # Create new binding for site
            $NewWebBinding = New-WebBinding -Name $SiteName -HostHeader $HostHeader -IPAddress $IP -Port $Port -Protocol "https" -SslFlags $SNI
            Write-Output $NewWebBinding
            # Create new binding in SSLBindings store
            $NewSslBinding = Get-Item -Path "Cert:\LocalMachine\My\$($newThumbprint)" | New-Item -Path "IIS:\SslBindings\$($IP)!$($Port)!$($HostHeader)"
            Write-Output $NewSslBinding
          } Else {
            $NewWebBinding = New-WebBinding -Name $SiteName -IPAddress $IP -Port $Port -Protocol "https" -SslFlags $SNI
            Write-Output $NewWebBinding
            $NewSslBinding = Get-Item -Path "Cert:\LocalMachine\My\$($newThumbprint)" | New-Item -Path "IIS:\SslBindings\$($IP)!$($Port)"
            Write-Output $NewSslBinding
          }
        }
        catch{
            Write-Output " + Unable to bind new cert!"
        }
    }

    If ($oldCert -And ($newCert -OR $Remove)) {
        Write-Output " + Deleting old certificate from Store..."
        try{
            If (Test-Path "cert:\LocalMachine\My\$oldThumbprint") {
              Remove-Item -Path cert:\LocalMachine\My\$oldThumbprint -DeleteKey
            }
        }
        catch{
            Write-Output " + Unable to delete old certificate from store!"
        }
    }
}

Write-Output " +++ Completed Certificate Update +++ "
 
# Stop logging 
Stop-Transcript
