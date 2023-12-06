$CurrentTime = Get-Date -Format "yyyy-MM-dd_hh-mm"

# Create Log File

$ErrorActionPreference = "SilentlyContinue"

Stop-Transcript | Out-Null

$ErrorActionPreference = "Continue"

# Start Transcript with random characters

Start-Transcript -Path "C:\Temp\Logs\SharePoint_$CurrentTime.log" -Append

# Connect To Service

$SiteURL = "<your sharepoint site>"

# Connect to SharePoint Online from PowerShell using PnP PowerShell

Connect-PnPOnline -Url $SiteURL

# Get All Site Collections

$SiteCollections = Get-PnPTenantSite -Detailed

# Export sites to CSV

$SiteCollections | Select-Object Url, Owner, StorageUsageCurrent, StorageQuota, Template, CompatibilityLevel, SharingCapability, Status, LockState,

LockIssue, WebsCount, RootWeb | Export-Csv -Path "C:\Temp\Sites\SitesList.csv" -NoTypeInformation

# End Transcript
Stop-Transcript
