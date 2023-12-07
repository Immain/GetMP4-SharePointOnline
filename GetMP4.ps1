# Import necessary modules
Import-Module Microsoft.Online.SharePoint.PowerShell
Import-Module PnP.PowerShell

# Set variables for the first script
$CSVSitesPath = "C:\Temp\Sites\SitesList.csv"
$CSVOutputPath = "C:\Temp\Sites\MP4FilesList.csv"
$DataCollection = @()

# Read sites from CSV
$Sites = Import-Csv -Path $CSVSitesPath

# Loop through each site
foreach ($Site in $Sites) {
    # Connect to SharePoint Online site
    Connect-PnPOnline $Site.URL -UseWebLogin

    # Get all Documents Libraries from the site
    $ExcludedLists = @("Style Library", "Wiki", "Form Templates", "Images", "Pages", "Site Pages", "Preservation Hold Library", "Site Assets")
    $DocumentLibraries = Get-PnPList | Where { $_.Hidden -eq $False -and $_.ItemCount -gt 0 -and $ExcludedLists -notcontains $_.Title -and $_.BaseType -eq "DocumentLibrary" }

    # Loop through all document libraries
    foreach ($List in $DocumentLibraries) {
        # Get all .mp4 Files that are not modified in the past 1 year or more!
        $global:counter = 0
        $ListItems = Get-PnPListItem -List $List -PageSize 2000 -Fields Created, Modified, FileLeafRef, FileRef, Editor -ScriptBlock `
        { 
            Param($items) 
            $global:counter += $items.Count; 
            Write-Progress -PercentComplete ($global:Counter / ($List.ItemCount) * 100) -Activity `
                "Getting Documents from Library '$($List.Title)'" -Status "Getting Files $global:Counter of $($List.ItemCount)"; 
        } | Where { $_.FileSystemObjectType -eq "File" -and $_.FieldValues.Modified -lt (Get-Date).AddDays(-10) -and $_.FieldValues.FileLeafRef -match '\.mp4$' }

        # Iterate through each item and retrieve file size information
        $FileData = @()
        foreach ($Item in $ListItems) {
            # Get file size information in megabytes
            $FileSizeinMB = [Math]::Round(($Item.FieldValues.File_x0020_Size / 1MB), 2)
            $File = Get-PnPProperty -ClientObject $Item -Property File
            $Versions = Get-PnPProperty -ClientObject $File -Property Versions
            $VersionSize = $Versions | Measure-Object -Property Size -Sum | Select-Object -expand Sum
            $VersionSizeinMB = [Math]::Round(($VersionSize / 1MB), 2)
            $TotalFileSizeMB = [Math]::Round(($FileSizeinMB + $VersionSizeinMB), 2)

            # Add file size information to the data collection
            $DataCollection += New-Object PSObject -Property ([Ordered] @{
                SiteURL         = $Site.URL
                Name            = $Item.FieldValues.FileLeafRef
                RelativeURL     = $Item.FieldValues.FileRef
                CreatedOn       = $Item.FieldValues.Created
                ModifiedBy      = $Item.FieldValues.Editor.Email
                ModifiedOn      = $Item.FieldValues.Modified
                FileSizeMB      = $TotalFileSizeMB
            })
        }

        # Display progress for listing files
        $ItemCounter = 0
        foreach ($Item in $ListItems) {
            $ItemCounter++
            Write-Progress -PercentComplete ($ItemCounter / ($ListItems.Count) * 100) -Activity "Listing .mp4 Files from Library '$($List.Title)' $ItemCounter of $($ListItems.Count)" -Status "Listing file '$($Item['FileLeafRef'])"
        }
    }

    # Disconnect from SharePoint Online site
    Disconnect-PnPOnline -Force
}

# Export data to CSV File
$DataCollection | Export-Csv -Path $CSVOutputPath -NoTypeInformation
