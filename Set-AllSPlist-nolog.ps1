[CmdletBinding()]

param (

 

    [Parameter(Mandatory = $true)]

    [string]$Url,

    [Parameter(Mandatory = $true)]

    [string]$SPList,

    [Parameter(Mandatory = $true)]

    [hashtable]$SPListItemValues,

    [Parameter(Mandatory = $true)]

    [string]$SPListItemID

)

 

$global:erroractionpreference = 1

 

$global:ErrorActionPreference = 'Stop'

Import-Module PnP.PowerShell

 

$appInfo = Get-AutomationPSCredential -Name 'SharePoint-CertPW'

$cert = Get-AutomationCertificate -Name 'SharePointAutomation'

$encodedPfx = [System.Convert]::ToBase64String($cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $appInfo.Password))

 

$Params = @{

    Url = $Url

    ClientId = $appInfo.Username

    CertificateBase64Encoded = $encodedPfx

    Tenant = 'tenantid.onmicrosoft.com'

    CertificatePassword = $appInfo.Password

}

 

Connect-PnPOnline @Params

 

$Null = Set-PnPListItem -List $SPList -Identity $SPListItemID -Values $SPListItemValues

$String = $SPListItemValues | Out-String

 

Write-Output "Set ID: $SPListItemID in list: $SPList to: $String"

 

Disconnect-PnPOnline
