
$progressPreference = 'silentlyContinue'
$serial = (Get-WmiObject -Class win32_bios).serialnumber

# Downloading and installing Azure AD and WindowsAutoPilotIntune Module
Write-Host "Downloading and installing AzureAD module" -ForegroundColor Cyan
Install-Module AzureAD,WindowsAutoPilotIntune,Microsoft.Graph.Intune -Force

# Importing required modules
Import-Module -Name AzureAD,WindowsAutoPilotIntune,Microsoft.Graph.Intune 


# Downloading and installing get-windowsautopilotinfo script
Write-Host "Downloading and installing get-windowsautopilotinfo script" -ForegroundColor Cyan
Install-Script -Name Get-WindowsAutoPilotInfo -Force

# Intune Login
Write-Host "Connecting to Microsoft Graph" -ForegroundColor Cyan

Connect-MSGraph

# Creating temporary folder to store autopilot csv file 

Write-Host "Checking if Temp folder exist in C:\" -ForegroundColor Cyan

IF (!(Test-Path C:\Temp) -eq $true) {

    Write-Host "Test folder was not found in C:\. Creating Test Folder..." -ForegroundColor Cyan
    New-Item -Path C:\Temp -ItemType Directory | Out-Null
}

Else { Write-Host "Test folder already exist" -ForegroundColor Green }

# Creating Autopilot csv file
Write-Host "Creating Autopilot CSV File" -ForegroundColor Cyan
Try {
    Get-WindowsAutoPilotInfo.ps1 -OutputFile "C:\Temp\$serial.csv"
    Write-Host "Successfully created autopilot csv file" -ForegroundColor Green}

Catch {
    write-host "Error: Something went wrong. Unable to create csv file." -foregroundcolor red 
Break }
 

#Importing CSV File into Intune
Write-Host "Importing Autopilot CSV File into Intune" -ForegroundColor Cyan
Try {
    Import-AutoPilotCSV -csvFile "C:\Temp\$serial.csv"
    Write-Host "Successfully imported autopilot csv file" -ForegroundColor Green}

Catch {
    Write-Host "Error: Something went wrong. Please check your csv file and try again"
    Break}
