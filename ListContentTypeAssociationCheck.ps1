# Import PnP PowerShell module
Import-Module PnP.PowerShell

# SharePoint online site URL
$siteUrl = Read-Host -Prompt "Enter your SharePoint site URL (e.g https://<tenant>.sharepoint.com/sites/contoso)"
# https://4l8x4l.sharepoint.com
$listName = Read-Host 'Enter your SharePoint list name'
# RiskEventRequestsList
$contentTypeName = Read-Host 'Enter your list content type name'
# RiskEventRequestsCT

# Connect to SharePoint Online site
Write-Host "Connecting to " $siteUrl -ForegroundColor Yellow 
Connect-PnPOnline -Url $siteUrl -Interactive

# Retrieve the list
$list = Get-PnPList -Identity $listName
if ($null -eq $list) {
    Write-Host "List '$listName' not found." -ForegroundColor Red
    Disconnect-PnPOnline
    exit
}

# Retrieve the content type
$contentType = Get-PnPContentType -List $listName -Identity $contentTypeName
if ($null -eq $contentType) {
    Write-Host "Content Type '$contentTypeName' not found in list '$listName'." -ForegroundColor Red
    Disconnect-PnPOnline
    exit
}

# Display content type details
Write-Host "Content Type Details:" -ForegroundColor Green
$contentType | Format-List

# Verify the form customizer properties
Write-Host "Checking form customizer properties..." -ForegroundColor Yellow

if ($contentType.NewFormClientSideComponentId) {
    Write-Host "NewFormClientSideComponentId: $($contentType.NewFormClientSideComponentId)" -ForegroundColor Green
} else {
    Write-Host "NewFormClientSideComponentId is not set." -ForegroundColor Red
}

if ($contentType.EditFormClientSideComponentId) {
    Write-Host "EditFormClientSideComponentId: $($contentType.EditFormClientSideComponentId)" -ForegroundColor Green
} else {
    Write-Host "EditFormClientSideComponentId is not set." -ForegroundColor Red
}

if ($contentType.DisplayFormClientSideComponentId) {
    Write-Host "DisplayFormClientSideComponentId: $($contentType.DisplayFormClientSideComponentId)" -ForegroundColor Green
} else {
    Write-Host "DisplayFormClientSideComponentId is not set." -ForegroundColor Red
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
