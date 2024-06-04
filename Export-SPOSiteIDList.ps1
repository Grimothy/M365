


# Check if PowerShell 7 is installed
$PSVersion = Get-Variable PSVersionTable -ErrorAction SilentlyContinue
if ($PSVersion -and $PSVersion.Value.PSVersion.Major -eq 7) {
    Write-Host "PowerShell 7 is already installed."
} else {
    Write-Host "PowerShell 7 is not installed. Installing the latest version..."

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

#connect to SPO with PNP powershell with interactive connection
Write-Host -ForegroundColor Green "Connecting to SharePoint Online"
$SPOSite = Read-Host "Please enter the full sharepoint URL"
Connect-PnPOnline -Url $SPOSite -Interactive

#Get list of Document Libraries
$DocLibList = Get-PnPList | Where-Object {$_.BaseType -eq "DocumentLibrary" -and {Where-Object $_.Title -NE "Web Part Gallery"} -and {Where-Object $_.Title -NE "Web Part Gallery" }
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

$DocLibList | Select-Object -Property Title,Id | Export-Csv .\DocLibList.csv