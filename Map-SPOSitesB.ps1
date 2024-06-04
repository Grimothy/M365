# User's Office 365 email address
$userEmail = (whoami /upn | Out-String).Trim()

# SharePoint Online site details
$siteID = "72c2c06b-d443-4da8-a5f6-f2e072dd83b8"
$tenantName = "First Bank of the Lake"
$webID = "c0ac8667-2520-4edb-a35a-baf3a43e0aa8"
$webTitle = "FileIndex"
$webUrl = "https://firstbanklake365.sharepoint.com/sites/FileIndex"

# Import the list of SharePoint Online lists from a CSV file
$SPOLIST = Import-Csv .\SPOLIST.csv

# Iterate over each list and sync if not already done
foreach ($list in $SPOLIST) {
    $listId = $list.ID
    $listTitle = $list.Title
    $registryPath = "HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\Tenants\$tenantName"
    $propertyPath = "$env:USERPROFILE\$tenantName\$webTitle - $listTitle"

    # Check if the list is already synced
    if (-not (Test-Path -Path $registryPath) -and (Get-ItemProperty -Path $registryPath -Name $propertyPath)) {
        # Construct the OneDrive sync URL
        $URL = "odopen://sync/?userEmail=" + [uri]::EscapeDataString($userEmail) +
               "&siteId=" + [uri]::EscapeDataString($siteID) +
               "&webId=" + [uri]::EscapeDataString($webID) +
               "&webTitle=" + [uri]::EscapeDataString($webTitle) +
               "&webUrl=" + [uri]::EscapeDataString($webUrl) +
               "&listId=" + [uri]::EscapeDataString($listId) +
               "&listTitle=" + [uri]::EscapeDataString($listTitle)

        # Start the OneDrive sync process
        Start-Process "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" -ArgumentList "/url:$URL"
    }
}