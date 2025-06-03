###########
# EnRoute Backup v1.1
# By Robert Kosmac
#
# This script will create a backup of EnRoute or EzyNest core files.
# These files are typically used in migrations and are required should
# anything go wrong.
#
# If Execution Policy is an issue, run from the CMD with the following:
# powershell.exe -executionpolicy bypass -file “EnRouteBackup.ps1”
#
###########

# Define base directories to search in
$searchDirs = @("C:\")
$folderPatterns = "EnRoute*", "EzyNest*"

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

function Backup-Version {
    param (
        [string]$sourcePath,
        [string]$destinationFolder
    )
    
    # These are the IMPORTANT files and folders.
    # More can be added here if required later.
    $itemsToBackup = @(
        "AutoTP",
        "NDrivers",
        "EnRoutePreferences.xml",
        "StrategyTemplates.ini",
        "ToolLibrary.ini"
    )

    $folderName = Split-Path $sourcePath -Leaf
    $versionBackupPath = Join-Path $destinationFolder $folderName
    New-Item -Path $versionBackupPath -ItemType Directory -Force | Out-Null

    foreach ($item in $itemsToBackup) {
        $source = Join-Path $sourcePath $item
        if (Test-Path $source) {
            Copy-Item -Path $source -Destination $versionBackupPath -Recurse -Force
        } else {
            Write-Host "Warning: $item not found in $sourcePath" -ForegroundColor Yellow
        }
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
        Backup-Version -sourcePath $install.FullName -destinationFolder $tempStaging
    }

    $zipName = "FullBackup_EZER_$timestamp.zip"
    $zipPath = Join-Path $backupFolder $zipName
    Compress-Archive -Path "$tempStaging\*" -DestinationPath $zipPath -Force
    Write-Host "`nAll versions backed up to: $zipPath" -ForegroundColor Green
} elseif ($choice -match '^\d+$' -and [int]$choice -lt $installPaths.Count) {
    $selectedPath = $installPaths[$choice].FullName
    $folderName = Split-Path $selectedPath -Leaf

    Backup-Version -sourcePath $selectedPath -destinationFolder $tempStaging

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
