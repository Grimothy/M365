$userEmail = $(whoami /upn | Out-String).Trim()
$siteID = "{72c2c06b-d443-4da8-a5f6-f2e072dd83b8}"
$tenantName = "First Bank of the Lake"
$webID = "{c0ac8667-2520-4edb-a35a-baf3a43e0aa8}"
$webTitle = "FileIndex"
$webUrl = "https://firstbanklake365.sharepoint.com/sites/FileIndex"

$SPOLIST = Import-Csv .\SPOLIST.csv
$SPOLIST| ForEach-Object 
{
    $listId ="{'$_.ID'}"
    $listTitle = $_.Title
    if (-not ((Test-Path -Path "HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\Tenants\$tenantName") -and (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\Tenants\$tenantName" -Name "$env:USERPROFILE\$tenantName\$webTitle - $listTitle"))) {
        $URL = "odopen://sync/?userEmail=" + [uri]::EscapeDataString($userEmail) + "&siteId=" + [uri]::EscapeDataString($siteID) + "&webId=" + [uri]::EscapeDataString($webID) + "&webTitle="+ [uri]::EscapeDataString($webTitle) + "&webUrl="+ [uri]::EscapeDataString($webUrl) + "&listId="+ [uri]::EscapeDataString($listId) + "&listTitle="+ [uri]::EscapeDataString($listTitle)
        Start-Process "$env:LOCALAPPDATA\Microsoft\OneDrive\Onedrive.exe" -ArgumentList "/url:$URL"
    }

}


$SPOLIST| ForEach-Object {
    Write-Host "SP DOC LIB IS " $_.Title  " and ID is " $_.Id
}
