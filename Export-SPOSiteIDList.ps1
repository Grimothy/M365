
<#
.SYNOPSIS
This script ensures PowerShell 7 and PnP.PowerShell module are installed, connects to SharePoint Online, and exports a list of document libraries.

.DESCRIPTION
This script performs a series of checks and actions to prepare a SharePoint Online environment for management. It begins by verifying if PowerShell 7 is installed and proceeds to install it if absent. Next, it checks for the PnP.PowerShell module and installs it if it's not available. After setting up the necessary tools, the script prompts the user to connect to a SharePoint Online site interactively. Once connected, it retrieves a list of all document libraries, excluding specific system libraries, and exports the details to a CSV file.

.PARAMETER PSVersion
The PSVersion parameter contains details about the installed PowerShell version.

.PARAMETER module
The module parameter checks for the availability of the PnP.PowerShell module.

.PARAMETER SPOSite
The SPOSite parameter takes the SharePoint Online site URL input from the user.

.PARAMETER DocLibList
The DocLibList parameter holds the list of document libraries after filtering out system libraries.

.EXAMPLE
PS C:\> .\YourScriptName.ps1
Executes the script to set up the environment and export the document libraries list.

.NOTES
Author: [C.J. Coulter]
Last Updated: [6/4/2024]
Version: 1.0


.LINK
https://github.com/Grimothy/M365
#>

# Check if PowerShell 7 is installed
if ($env:Path.Contains("C:\Program Files\PowerShell\7\") -eq $true)
    {
    Write-Host "Powershell 7 already installed on this system"
} else { 
    Write-Host "Powershell 7 NOT currently installed on system. Installing the latest version..."
    # Define the download path and installer file name
    $downloadPath = "$env:TEMP\PowerShell7-Installer.msi"
    $installerUrl = "https://aka.ms/powershell-release?tag=stable"

    # Download the installer
    Invoke-WebRequest -Uri $installerUrl -OutFile $downloadPath

    # Install PowerShell 7 silently
    Start-Process -FilePath msiexec.exe -ArgumentList "/i `"$downloadPath`" /quiet" -Wait

    # Remove the installer file
    Remove-Item -Path $downloadPath -Force

    Write-Host "PowerShell 7 has been installed."
}

# Check if PnP.PowerShell module is installed
$module = Get-Module -ListAvailable -Name PnP.PowerShell

if ($module) {
    Write-Host "PnP.PowerShell module is already installed."
} else {
    Write-Host "PnP.PowerShell module is not installed. Installing now..."

    # Install the PnP.PowerShell module for the current user
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force

    Write-Host "PnP.PowerShell module has been installed."
}

#check if script is running in powershell 7 

if ($PSVersionTable.PSVersion.Major -ne "7")
{
    Write-Host -ForegroundColor Yellow "Script must run using PowerShell 7. The script will be reopened with Powershell 7"
    pwsh .\Export-SPOSiteIDList.ps1
    exit
}

#Import PNP Module
import-module PnP.PowerShell
#connect to SPO with PNP powershell with interactive connection
Write-Host -ForegroundColor Green "Connecting to SharePoint Online"
$SPOSite = Read-Host "Please enter the full sharepoint URL"
Connect-PnPOnline -Url $SPOSite -Interactive

#Get list of Document Libraries
$DocLibList = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary"`
    -and $_.Title -ne "Web Part Gallery"`
    -and $_.Title -ne "Theme Gallery"`
    -and $_.Title -ne "Style Library"`
    -and $_.Title -ne "Site Pages"`
    -and $_.Title -ne "Site Assets"`
    -and $_.Title -ne "Master Page Gallery"`
    -and $_.Title -ne "List Template Gallery"`
    -and $_.Title -ne "Form Templates"`
    -and $_.Title -ne "Documents"`

}


# Load the Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a new SaveFileDialog
$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

# Set initial directory, filter, and other properties if needed
$saveFileDialog.initialDirectory = [Environment]::GetFolderPath("Desktop")
$saveFileDialog.filter = "CSV files (*.csv)|*.csv"
$saveFileDialog.Title = "Select a location to save the document library list"

# Show the Save File dialog
$result = $saveFileDialog.ShowDialog()

# Check if the user clicked the 'Save' button
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $fileSavePath = $saveFileDialog.FileName
    # Export the document library list to the chosen path
    $DocLibList | Select-Object -Property Title,Id | Export-Csv -Path $fileSavePath -NoTypeInformation
    Write-Host "Document library list has been saved to: $fileSavePath"
} else {
    Write-Host "Export cancelled. No file location was selected."
}

# Clean up resources
$saveFileDialog.Dispose()


