
function CheckPnPPowerShellModule {

  try {
    Write-Host "Checking if SharePoint Online PnP PowerShell Module is Installed..." -f Yellow -NoNewline
    $SharePointPnPPowerShellOnline = Get-Module -ListAvailable "SharePointPnPPowerShellOnline"

    if (!$SharePointPnPPowerShellOnline)
    {
      Write-Host "No!" -f Green

      #Check if script is executed under elevated permissions - Run as Administrator
      if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator"))
      {
        Write-Host "Please Run this script in elevated mode (Run as Administrator)! " -NoNewline
        Read-Host "Press any key to continue"
        exit
      }

      Write-Host "Installing SharePoint Online PnP PowerShell Module..." -f Yellow -NoNewline
      Install-Module SharePointPnPPowerShellOnline -Force -Confirm:$False
      Write-Host "Done!" -f Green
    }
    else
    {
      Write-Host "Yes!" -f Green
      Write-Host "Importing SharePoint Online PnP PowerShell Module..." -f Yellow -NoNewline
      Import-Module SharePointPnPPowerShellOnline -DisableNameChecking
      Write-Host "Done!" -f Green
    }

  }
  catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor red
  }


}

function Write-Log ($Message,$LogFile)
{

  $DateTimeStamp = Get-Date -f "yyyy-MM-dd HH:mm:ss"
  #$DateTimeStamp + "," + $Message | Out-File -Encoding Default -Append -FilePath $LogFile

  "   " | Out-File -Encoding Default -Append -FilePath $LogFile
  "************************************************************************************************************" | Out-File -Encoding Default -Append -FilePath $LogFile
  "Log happend at time: $DateTimeStamp" | Out-File -Encoding Default -Append -FilePath $LogFile
  "Log message: $Message" | Out-File -Encoding Default -Append -FilePath $LogFile
  "------------------------------------------------------------------------------------------------------------" | Out-File -Encoding Default -Append -FilePath $LogFile


}


