# Import the list of SharePoint Online lists from a CSV file
$SPOLIST = Import-Csv .\SPOLIST.csv 
# User's Office 365 email address
$userEmail = (whoami /upn | Out-String).Trim()

# SharePoint Online site details
$siteID = "{4ccb4164-b6e2-45de-ad90-1de543cd9e29}"
$tenantName = "Melillo Consulting"
$webID = "{2caa6cbd-4b9a-4b91-a93c-5c9f0127e44f}"
$webTitle = "CJCTEST"
$webUrl = "https://melilloconsulting.sharepoint.com/sites/CJCTEST"
$registryPath = "HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\Tenants\$tenantName"

# Iterate over each list and sync if not already done
foreach ($list in $SPOLIST) {
    $propertyPath = "$env:USERPROFILE\$tenantName\$webTitle - " + $list.Title

    # Check if the list is already synced
    $propertyExists = $false
    try {
        $property = Get-ItemProperty -Path $registryPath -Name $propertyPath -ErrorAction Stop
        $propertyExists = $true
    } catch {
        $propertyExists = $false
    }
    if (-not (Test-Path -Path $registryPath) -or -not $propertyExists) {
        # Construct the OneDrive sync URL
        Write-Host "not synced"
        $URL = "odopen://sync/?userEmail=" + [uri]::EscapeDataString($userEmail) +
               "&siteId=" + [uri]::EscapeDataString($siteID) +
               "&webId=" + [uri]::EscapeDataString($webID) +
               "&webTitle=" + [uri]::EscapeDataString($webTitle) +
               "&webUrl=" + [uri]::EscapeDataString($webUrl) +
               "&listId=" + [uri]::EscapeDataString($list.id) +
               "&listTitle=" + [uri]::EscapeDataString($list.Title)

        # Start the OneDrive sync process
        # Uncomment the line below that corresponds to your OneDrive installation path
        start-Process "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe" -ArgumentList "/url:$URL" -ErrorAction SilentlyContinue -wait
        Start-Process "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe" -ArgumentList "/url:$URL" -ErrorAction SilentlyContinue -Wait
        #Write-Host "processing list " $list.Id  " with list title " $list.Title
    }
}
