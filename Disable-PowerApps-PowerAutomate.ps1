<#
.SYNOPSIS
=========
Disable PowerApps and PowerAutomate at Tenant Level

.DESCRIPTION
============
This Script Disable PowerApps and PowerAutomate at Tenant Level 


.EXAMPLE
========
.\Disable-PowerApps-PowerAutomate.ps1 -Tenant [TenantName] -SiteUrl [SiteUrl] -ApplicationID [AzureAppID] -AppRegisterName [CertificateAppRegisterName]

.\Disable-PowerApps-PowerAutomate.ps1 -Tenant vinovijayabalan.onmicrosoft.com -SiteUrl https://vinovijayabalan.sharepoint.com -ApplicationID 10f59a78-bf91-4fc0-b83d-28260a6ce75c -AppRegisterName appregister

Connect using AppRegister certificate if ThumbprintID is not working

.\Disable-PowerApps-PowerAutomate.ps1 -Tenant [TenantName] -SiteUrl [SiteUrl] -ApplicationID [AzureAppID] -CertificatePath ['c:\mycertificate.pfx'] -CertificatePassword (ConvertTo-SecureString -AsPlainText 'myprivatekeypassword' -Force) 


.NOTES
=======
Created by   : Vinothkumar Vijayabalan
Date Coded   : 21/Jan/2021 
#>



#######################################################
# MAIN section                                        #
#######################################################


[CmdletBinding()]
param(

  [Parameter(Mandatory = $true)]
  [string]$Tenant,

  [Parameter(Mandatory = $true)]
  [string]$SiteUrl,

  [Parameter(Mandatory = $true)]
  [string]$ApplicationID,

  [Parameter(Mandatory = $true)]
  [string]$AppRegisterName,

  [Parameter(Mandatory = $false)]
  [string]$CertificatePath,

  [Parameter(Mandatory = $false)]
  [string]$CertificatePassword

)

#Get the parent folder path
$PSScriptFolder = $PSScriptRoot
$Date = Get-Date
$fileDate = $Date.ToString('MM-dd-yyyy_hh-mm-ss')
$logfile = ($PSScriptRoot + "\PSLogs" + "\Log_Disable_PowerApps and PowerAutomate_" + $filedate + ".log")

#import Modules
. "$PSScriptFolder\Common-Function.ps1"

#Check if SharePoint Online PnP PowerShell module has been installed
$GetPnP = CheckPnPPowerShellModule


########## Start Function - Disabling PowerApps and PowerAutomate #############
try
{

  #Login Service Principle
  $CertInfo = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $PSItem.Subject -eq "CN=$AppRegisterName" } | Select-Object -Last 1
  $thumb = $CertInfo.Thumbprint
  #$tenantId = (Invoke-WebRequest -UseBasicParsing https://login.windows.net/$Tenant/.well-known/openid-configuration|ConvertFrom-Json).token_endpoint.Split('/')[3]

  Write-Host "Conneccting PNP Online using ThumbPrint ID" 
  Connect-PnPOnline -Url $SiteUrl -ClientId $ApplicationID -Thumbprint $thumb -Tenant $Tenant
  if ($CertificatePath)
  { Connect-PnPOnline -Url $SiteUrl -ClientId $ApplicationID -Tenant $tenant -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword }
  Write-Log "Connected using ThumbPrindID" $logfile

  #Get All Site collections - Exclude BOT,communication site, Video Portals and MySites
  $Sites = Get-PnPTenantSite -Template GROUP#0
  $Sites | ForEach-Object {

    #Connect to each site collection and Disabling Powerapps & Power Automate
    $ctxURL = $_.URL
    $ctx = Connect-PnPOnline -Url $ctxURL -ClientId $ApplicationID -Thumbprint $thumb -Tenant $Tenant
    if ($CertificatePath)
    { $ctx = Connect-PnPOnline -Url $SiteUrl -ClientId $ApplicationID -Tenant $tenant -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword }

    Write-Host "Processing Site Collection:" $_.URL -f green -NoNewline
    $ctx = Get-PnPContext
    if (!$ctx.Site.DisableAppViews) { $ctx.Site.DisableAppViews = $true }
    if (!$ctx.Site.DisableFlows) { $ctx.Site.DisableFlows = $true }
    $ctx.ExecuteQuery();
    Write-Log "Powerapps and PowerAutomate has been disabled successfully -> Sitecollection URL is : $ctxURL" $logfile

  }


}

catch [Exception]
{
  #Logging Error message
  $ErrorMessage = $_.Exception.Message
  Write-Log $ErrorMessage $logfile
  Write-Log "Powerapps & Powerauotmate has been falied !! Please check the configuration!!" $logfile
  Write-Host " Powerapps & Powerauotmate has been falied !! Please check the configuration!!" -ForegroundColor red

}

Disconnect-PnPOnline
Write-Log "PnP Connection has been disconnected successfully!!" $logfile


#---- End Function--->

################## End Main Function ########################################################################################################3
