# Update IIS certificate
PowerShell script for automation certificate deployment (credits to [Jeremy](http://www.jhouseconsulting.com/2015/01/04/script-to-import-and-bind-a-certificate-to-the-default-web-site-1548)). Tested on Windows Server 2012 R2.

## Requirements
* PowerShell Version >= 3
* IIS Version >= 8.5 with Webadministration module
* Certificate file [.pfx], private key included

## Usage
```
PS C:\> update-iis-certificate.ps1 [[-PFXPath] <String>] -CertSubject <String> [-PFXPassword <String> ]
                                   [-SiteName <String> ][-IP <String> ][-Port <int> ][-HostHeader <String> ]
                                   [-SNI][-Remove][-ExcludeLocalServerCert]
```
All parameters in square brackets are optional. Most of them are for custimzed webbindings, see [Microsoft - Technet Docs](https://technet.microsoft.com/de-de/library/hh867854(v=wps.630).aspx). 
The ExcludeLocalServerCert is forced to $True if left off. You really never want this set to false, especially if using a wildcard certificate. It's there mainly for flexibility.

If the password contains a $ sign, you must escape it with the ` ` character.

### Examples
Install/update certificate in certificate store and webbindings for "Default Web Site" in IIS:
```
PS C:\> update-iis-certificate.ps1 ".\example.com.pfx" -CertSubject "example.com" -PFXPassword "P@ssw0rd"
```      
Remove certificate from certificate store and webbindings for "Default Web Site" from IIS:
```
PS C:\> update-iis-certificate.ps1 -CertSubject "example.com" -Remove
```

### Logs
A log file will either be written to %windir%\Temp or to the %LogPath% Task Sequence variable if running from an SCCM\MDT Task.
