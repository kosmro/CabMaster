###
# EnRoute Backup v1.1
# By Robert Kosmac
#
# This script will create a backup of EnRoute or EzyNest core files.
# These files are typically used in migrations and are required should
# anything go wrong.
#
# If Execution Policy is an issue, run from the CMD with the following:
# powershell.exe -executionpolicy bypass -file "EnRouteBackup.ps1"
#
###

######################## VARIABLES ########################
# Define base directories to search in
$searchDirs = @("C:\")
$folderPatterns = "EnRoute*", "EzyNest*"

# List of file types to ignore
$excludedExtensions = @(".exe", ".dll")
###########################################################


# Progress tracking
$global:copyTotal = 0
$global:copyCurrent = 0
$global:copyStartTime = Get-Date


# Search and output a list of found Install Paths
function Get-InstallationPaths {
    param (
        [string[]]$baseDirs,
        [string[]]$patterns
    )

    $installPaths = @()

    # This will search in the folders, down to 2 folder levels
    foreach ($baseDir in $baseDirs) {
        if (-not (Test-Path $baseDir)) { continue }

        $level1 = Get-ChildItem -Path $baseDir -Directory -ErrorAction SilentlyContinue
        foreach ($dir1 in $level1) {
            foreach ($pattern in $patterns) {
                if ($dir1.Name -like $pattern) {
                    $installPaths += $dir1
                }
            }

            $level2 = Get-ChildItem -Path $dir1.FullName -Directory -ErrorAction SilentlyContinue
            foreach ($dir2 in $level2) {
                foreach ($pattern in $patterns) {
                    if ($dir2.Name -like $pattern) {
                        $installPaths += $dir2
                    }
                }
            }
        }
    }

    return $installPaths
}


# Create a copy of the selected version to Backup
function Backup-Version {
    param (
        [string]$sourcePath,
        [string]$destinationFolder,
        [string[]]$excludedExtensions,
        [string]$mode  # "1" = important only, "2" = full
    )

    $importantItems = @(
        "AutoTP",
        "NDrivers",
        "EnRoutePreferences.xml",
        "StrategyTemplates.ini",
        "ToolLibrary.ini"
    )

    $itemsToBackup = if ($mode -eq "1") {
        $importantItems
    } else {
        Get-ChildItem -Path $sourcePath -Recurse -File | ForEach-Object {
            $_.FullName.Substring($sourcePath.Length).TrimStart('\').Split('\')[0]
        } | Sort-Object -Unique
    }

    $folderName = Split-Path $sourcePath -Leaf
    $versionBackupPath = Join-Path $destinationFolder $folderName
    New-Item -Path $versionBackupPath -ItemType Directory -Force | Out-Null

    # Pre-count eligible files
    $localFiles = @()
    foreach ($item in $itemsToBackup) {
        $source = Join-Path $sourcePath $item
        if (Test-Path $source -PathType Container) {
            $localFiles += Get-ChildItem -Path $source -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
                $excludedExtensions -notcontains $_.Extension.ToLower()
            }
        } elseif (Test-Path $source -PathType Leaf) {
            if ($excludedExtensions -notcontains ([System.IO.Path]::GetExtension($source).ToLower())) {
                $localFiles += Get-Item $source
            }
        }
    }

    $global:copyTotal += $localFiles.Count

    # Copy files with progress
    foreach ($file in $localFiles) {
        $relativePath = $file.FullName.Substring($sourcePath.Length).TrimStart('\')
        $destPath = Join-Path $versionBackupPath $relativePath
        $destDir = Split-Path $destPath -Parent

        if (-not (Test-Path $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }

        Copy-Item -Path $file.FullName -Destination $destPath -Force

        $global:copyCurrent++
        $percent = [math]::Round(($global:copyCurrent / $global:copyTotal) * 100, 0)

        $elapsed = (Get-Date) - $global:copyStartTime
        $remaining = if ($global:copyCurrent -gt 0) {
            $avg = $elapsed.TotalSeconds / $global:copyCurrent
            [timespan]::FromSeconds(($global:copyTotal - $global:copyCurrent) * $avg)
        } else {
            [timespan]::Zero
        }

        Write-Progress -Activity "Backing up: $folderName" `
                       -Status "$global:copyCurrent of $global:copyTotal files" `
                       -PercentComplete $percent `
                       -SecondsRemaining $remaining.TotalSeconds `
                       -CurrentOperation $file.Name
    }
}



# MAIN EXECUTION

$installPaths = Get-InstallationPaths -baseDirs $searchDirs -patterns $folderPatterns

if ($installPaths.Count -eq 0) {
    Write-Host "No installations found for EnRoute or EzyNest." -ForegroundColor Red
    exit
}

# Display menu
Write-Host "`nFound Installations:" -ForegroundColor Cyan
for ($i = 0; $i -lt $installPaths.Count; $i++) {
    Write-Host "[$i] $($installPaths[$i].FullName)"
}
Write-Host "[A] Backup ALL versions found"

# Ask user for input
$choice = Read-Host "`nEnter the number of the version to back up, or 'A' for all"


Write-Host "`nBackup Type:"
Write-Host "[1] Only important files/folders"
Write-Host "[2] Everything in the folder"
$backupMode = Read-Host "Select backup mode (1 or 2)"
if ($backupMode -ne '1' -and $backupMode -ne '2') {
    Write-Host "Invalid choice. Exiting." -ForegroundColor Red
    exit
}



# Prepare paths
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupFolder = "C:\AllMasterSoftware\Backups"
$tempStaging = Join-Path $env:TEMP "BackupTemp_$timestamp"

if (-not (Test-Path $backupFolder)) {
    New-Item -Path $backupFolder -ItemType Directory -Force
}

New-Item -Path $tempStaging -ItemType Directory -Force | Out-Null

if ($choice -eq 'A' -or $choice -eq 'a') {
    # Backup all versions
    foreach ($install in $installPaths) {
        Backup-Version -sourcePath $install.FullName -destinationFolder $tempStaging -excludedExtensions $excludedExtensions -mode $backupMode

    }

    $zipName = "FullBackup_EZER_$timestamp.zip"
    $zipPath = Join-Path $backupFolder $zipName
    Compress-Archive -Path "$tempStaging\*" -DestinationPath $zipPath -Force
    Write-Host "`nAll versions backed up to: $zipPath" -ForegroundColor Green
} elseif ($choice -match '^\d+$' -and [int]$choice -lt $installPaths.Count) {
    $selectedPath = $installPaths[$choice].FullName
    $folderName = Split-Path $selectedPath -Leaf

    Backup-Version -sourcePath $selectedPath -destinationFolder $tempStaging -excludedExtensions $excludedExtensions -mode $backupMode


    $zipName = "$folderName" + "_$timestamp.zip"
    $zipPath = Join-Path $backupFolder $zipName
    Compress-Archive -Path "$tempStaging\*" -DestinationPath $zipPath -Force
    Write-Host "`nBackup complete: $zipPath" -ForegroundColor Green
} else {
    Write-Host "Invalid input. Exiting." -ForegroundColor Red
    Remove-Item -Path $tempStaging -Recurse -Force -ErrorAction SilentlyContinue
    exit
}

# Cleanup
Remove-Item -Path $tempStaging -Recurse -Force -ErrorAction SilentlyContinue
