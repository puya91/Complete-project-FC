# Importing PnP module for PowerShell
Import-Module PnP.PowerShell

# SharePoint online site URL
$siteUrl = Read-Host -Prompt "Enter your SharePoint site URL (e.g https://<tenant>.sharepoint.com/sites/contoso)"

# Connect to SharePoint Online site
Write-Host "Connecting to " $siteUrl -ForegroundColor Yellow 
Connect-PnPOnline -Url $siteUrl -Interactive

# Enter SharePoint display list name and content type name
$listName = Read-Host 'Enter your SharePoint list name'
$contentTypeName = Read-Host 'Enter your list content type name'

# Enter new form component Id
# $newFormComponentId = Read-Host 'Enter New form component Id'

# Enter edit form component Id
$editFormComponentId = Read-Host 'Enter Edit form component Id'

# Enter display form component Id
$displayFormComponentId = Read-Host 'Enter Display form component Id'

# Associate form customizer extension with SharePoint list forms
Set-PnPContentType -Identity $contentTypeName -List $listName -EditFormClientSideComponentId $editFormComponentId -DisplayFormClientSideComponentId $displayFormComponentId

# Disconnect SharePoint online connection
Disconnect-PnPOnline