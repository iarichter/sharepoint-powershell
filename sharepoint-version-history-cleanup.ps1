# Set the script location
$scriptLocation = "C:\Path\To\Your\Script\Location"
cd $scriptLocation

# Configuration Parameters
$csvFile = "$scriptLocation\siteList.csv"
# CSV of SharePoint sites you want to clean up.
$siteTable = Import-Csv -path $csvFile -Delimiter ","
$Admin = "admin@yourdomain.com"

# Set values for specific number of versions to keep based on criteria
$DefaultVersionsToKeep = 100
$ArchiveVersionsToKeep = 5
$OverFiveYearsOldVersionsToKeep = 5
$OverTwoYearsOldVersionsToKeep = 25
$OverOneYearOldVersionsToKeep = 50

# Initialize counters and arrays
$versionsDeletedCount = 0
$versionSize = 0
$libraySizeDeleted = 0
$deletedVersionsSizeSum = 0
$deletedFileSize = 0
$deletedFileSizeMB = 0
$backupArray = @()
$today = Get-Date

# Uncomment and modify the URL to connect to your SharePoint admin site
# Connect-SPOService -url "https://yourdomain-admin.sharepoint.com/" -modernauth $true

foreach ($row in $siteTable) {
    # Connect to PnP Online for each site
    $SiteUrl = $row.Site
    Write-Host "Checking Libraries for site:" $SiteUrl
    Connect-PnPOnline -Url $SiteUrl -Interactive

    # Get the PnP context
    $Ctx = Get-PnPContext

    # Exclude certain libraries from processing
    $ExcludedLists = @("Form Templates", "Preservation Hold Library", "Site Assets", "Pages", "Site Pages", "Images",
        "Site Collection Documents", "Site Collection Images", "Style Library", "Teams Wiki Data")

    # Get all document libraries excluding the ones in the excluded list
    $DocumentLibraries = Get-PnPList | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Title -notin $ExcludedLists -and $_.Hidden -eq $false }

    # Iterate through each document library
    ForEach ($Library in $DocumentLibraries) {
        Write-Host "Processing Document Library:" $Library.Title -ForegroundColor Magenta
        $libraryStart = Get-Date

        # Set the number of major versions to keep
        Set-PnPList -Identity $Library -MajorVersions $DefaultVersionsToKeep

        # Get all items from the list, excluding folders
        $ListItems = Get-PnPListItem -List $Library -PageSize 2000 | Where { $_.FileSystemObjectType -eq "File" }
        $libraySizeDeleted = 0

        # Prepare output object
        $outputObject = "" | Select "site", "library", "sizeGB", "monthlySavings", "annualSavings", "dateCleaned", "minutesToScan"
        $outputObject.site = $SiteUrl
        $outputObject.library = $Library.Title

        # Loop through each file
        ForEach ($Item in $ListItems) {
            $deletedFileSize = 0

            # Get file versions
            $File = $Item.File
            $Path = $Item["FileRef"]
            $Versions = $File.Versions
            $Ctx.Load($File)
            $Ctx.Load($Versions)
            $Ctx.ExecuteQuery()

            Write-Host -ForegroundColor Yellow "`tScanning File:" $File.Name
            Start-Sleep -Milliseconds 500

            # Determine the number of versions to keep based on file age and location
            $VersionsCount = $Versions.Count
            $versionMessage = ""
            $fileTimeLastModified = $File.TimeLastModified
            $timeDifference = $today - $fileTimeLastModified

            if ($Path -like "*archive*") {
                $ActualVersionsToKeep = $ArchiveVersionsToKeep
                $versionMessage = "File in an Archive folder keeping " + $ArchiveVersionsToKeep + " versions."
            } elseif ($timeDifference.Days -gt 1825) {
                $ActualVersionsToKeep = $OverFiveYearsOldVersionsToKeep
                $versionMessage = "Greater than 5 years old, keeping " + $OverFiveYearsOldVersionsToKeep + " versions."
            } elseif ($timeDifference.Days -gt 730) {
                $ActualVersionsToKeep = $OverTwoYearsOldVersionsToKeep
                $versionMessage = "Between 2 and 5 years old, keeping " + $OverTwoYearsOldVersionsToKeep + " versions."
            } elseif ($timeDifference.Days -gt 365) {
                $ActualVersionsToKeep = $OverOneYearOldVersionsToKeep
                $versionMessage = "Between 1 and 2 years old, keeping " + $OverOneYearOldVersionsToKeep + " versions."
            } else {
                $ActualVersionsToKeep = $DefaultVersionsToKeep
                $versionMessage = "Keeping Default Number of Versions: " + $DefaultVersionsToKeep
            }

            $VersionsToDelete = $VersionsCount - $ActualVersionsToKeep
            If ($VersionsToDelete -gt 0) {
                Write-Host -ForegroundColor Cyan "`t Total Number of Versions of the File:" $VersionsCount
                Write-Host $versionMessage
                $deletedVersionsSizeSum = 0

                # Delete excess versions
                $VersionCounter = 0
                For ($i = 0; $i -lt $VersionsToDelete; $i++) {
                    $versionSize = 0
                    if ($Versions[$VersionCounter].IsCurrentVersion) {
                        $VersionCounter++
                        Write-Host "Skipping current version file"
                        Continue
                    }
                    $Versions[$VersionCounter].DeleteObject()
                    $versionSize = $Versions[$VersionCounter].Size
                    Write-Host -NoNewline -ForegroundColor Cyan "`t Deleted Version:" $Versions[$VersionCounter].VersionLabel
                    Write-Host -ForegroundColor DarkCyan "`t Version Size (Bytes):" $versionSize
                    $deletedVersionsSizeSum += $versionSize
                    $versionsDeletedCount++
                    Write-Host -NoNewline -ForegroundColor Magenta "`t Sum size:" $deletedVersionsSizeSum
                }
                $Ctx.ExecuteQuery()
                $deletedFileSize = $deletedVersionsSizeSum
                Write-Host -NoNewline -ForegroundColor Green "`t Version History is cleaned for the File: " $File.Name
                $deletedFileSizeMB = $deletedFileSize / 1MB
                Write-Host -ForegroundColor DarkGreen " File size (MB):" $deletedFileSizeMB
            }
            $libraySizeDeleted += $deletedFileSize
        }

        # Calculate and display savings
        $librarySizeDeltedGB = $libraySizeDeleted / 1GB
        Write-Host "Library Total File Size Removed (GB):" $librarySizeDeltedGB
        $outputObject.sizeGB = $librarySizeDeltedGB
        $monthlySavings = [math]::Round($librarySizeDeltedGB * 0.16, 2)
        $annualSavings = $monthlySavings * 12
        Write-Host "Monthly Savings($):" $monthlySavings " | Annual Savings($):" $annualSavings

        # Store output details
        $outputObject.monthlySavings = $monthlySavings
        $outputObject.annualSavings = $annualSavings
        $libraryFinish = Get-Date
        $libraryRunTime = [math]::Round(($libraryFinish - $libraryStart).TotalMinutes, 3)
        $outputObject.minutesToScan = $libraryRunTime
        $outputObject.dateCleaned = Get-Date -Format "yyyy-MM-dd"
        $outputArray += $outputObject
        $outputObject = $null
    }

    # Export results to CSV
    $outputArray | Export-Csv -Path "$scriptLocation\fileCleanUpResults.csv" -Append
    $backupArray += $outputArray
}
