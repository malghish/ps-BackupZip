# Enhanced PowerShell Backup Script
# Purpose: Automates the backup of folders listed in a CSV file to a compressed archive

# Set error action preference and enable verbose output
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"
$settingsFile = "backup_settings.json"

# Function to read JSON settings
function Get-BackupSettings {
    if (Test-Path $settingsFile) {
        try {
            return Get-Content $settingsFile | ConvertFrom-Json
        } catch {
            Write-Host "Error reading settings file. It might be corrupted." -ForegroundColor Red
        }
    }
    return $null
}

# Function to save settings to JSON
function Save-BackupSettings($destinationFolder, $logExtensions, $defaultLogAction) {
    $settings = @{ 
        DestinationFolder = $destinationFolder
        LogExtensions = $logExtensions
        DefaultLogAction = $defaultLogAction
    } | ConvertTo-Json
    $settings | Out-File -FilePath $settingsFile
    Write-Host "Settings saved to $settingsFile" -ForegroundColor Yellow
}

function Get-FolderSize {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        $size = (Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue | 
                Measure-Object -Property Length -Sum).Sum
        
        if ($size -gt 1GB) {
            return "{0:N2} GB" -f ($size / 1GB)
        }
        else {
            return "{0:N2} MB" -f ($size / 1MB)
        }
    }
    catch {
        Write-Warning "Could not calculate size for path: $Path"
        return "Unknown size"
    }
}

function Get-LogFilesInfo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        
        [Parameter(Mandatory = $true)]
        [string[]]$LogExtensions
    )
    
    try {
        $logFiles = Get-ChildItem -Path $Path -Include $LogExtensions -Recurse -File -ErrorAction SilentlyContinue
        
        $logCount = $logFiles.Count
        $logSize = ($logFiles | Measure-Object -Property Length -Sum).Sum
        
        if ($logSize -gt 1GB) {
            $logSizeFormatted = "{0:N2} GB" -f ($logSize / 1GB)
        }
        else {
            $logSizeFormatted = "{0:N2} MB" -f ($logSize / 1MB)
        }
        
        return @{
            Count = $logCount
            Size = $logSizeFormatted
            Files = $logFiles
        }
    }
    catch {
        Write-Warning "Could not get log files info for path: $Path"
        return @{
            Count = 0
            Size = "0 MB"
            Files = @()
        }
    }
}

function New-BackupArchive {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ArchiveName,
        
        [Parameter(Mandatory = $false)]
        [string[]]$ExcludeFiles = @(),
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipLockedFiles = $true
    )
    try {
        # Ensure source path exists
        if (-not (Test-Path -Path $SourcePath)) {
            Write-Error "Source path does not exist: $SourcePath"
            return $false
        }

        # Create destination directory if it doesn't exist
        if (-not (Test-Path -Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        # Create a temporary directory for filtered content
        $tempDir = Join-Path -Path $env:TEMP -ChildPath ([Guid]::NewGuid().ToString())
        New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
        
        # Convert excluded files to a hash set for fast lookup
        $excludeFilter = @{}
        foreach ($file in $ExcludeFiles) {
            $excludeFilter[[System.IO.Path]::GetFileName($file)] = $true
        }
        
        # Copy items except excluded files
        $lockedFiles = @()
        Get-ChildItem -Path $SourcePath -Recurse | ForEach-Object {
            if ($_.PSIsContainer) {
                # Create directory structure
                $relativePath = $_.FullName.Substring($SourcePath.Length).TrimStart('\')
                $targetDir = Join-Path -Path $tempDir -ChildPath $relativePath
                if (-not (Test-Path $targetDir)) {
                    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                }
            } else {
                # For files, check if they should be excluded
                if (-not $excludeFilter.ContainsKey($_.Name)) {
                    try {
                        $relativePath = $_.FullName.Substring($SourcePath.Length).TrimStart('\')
                        $targetFile = Join-Path -Path $tempDir -ChildPath $relativePath
                        
                        # Create parent directory if it doesn't exist
                        $parentDir = Split-Path -Path $targetFile -Parent
                        if (-not (Test-Path $parentDir)) {
                            New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
                        }
                        
                        # Test if file is locked by attempting to open it with read access
                        $fileIsLocked = $false
                        try {
                            $fileStream = [System.IO.File]::Open($_.FullName, 'Open', 'Read', 'None')
                            $fileStream.Close()
                            $fileStream.Dispose()
                        } catch {
                            $fileIsLocked = $true
                        }
                        
                        if ($fileIsLocked) {
                            $lockedFiles += $_.FullName
                            Write-Warning "Skipping locked file: $($_.FullName)"
                        } else {
                            Copy-Item -Path $_.FullName -Destination $targetFile -Force -ErrorAction Stop
                        }
                    } catch {
                        Write-Warning "Failed to copy file: $($_.FullName). Error: $_"
                        $lockedFiles += $_.FullName
                    }
                }
            }
        }
        
        # Check if we have any files to archive
        $filesToArchive = Get-ChildItem -Path $tempDir -Recurse -File
        if ($filesToArchive.Count -eq 0) {
            Write-Warning "No files to archive. All files may be locked or excluded."
            return $false
        }
        
        # Create the archive from the temporary directory
        $archivePath = Join-Path -Path $DestinationPath -ChildPath "$ArchiveName.zip"
        Compress-Archive -Path "$tempDir\*" -DestinationPath $archivePath -CompressionLevel Optimal -Force
        
        # Report results
        Write-Host "Backup archive created successfully at: $archivePath" -ForegroundColor Green
        
        if ($lockedFiles.Count -gt 0) {
            Write-Warning "The following $($lockedFiles.Count) files were locked and not included in the backup:"
            $lockedFiles | ForEach-Object { Write-Warning "  - $_" }
        }
    } catch {
        Write-Error "Failed to create archive: $_"
        return $false
    } finally {
        # Cleanup the temporary directory
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    return $true
}

function Start-BackupProcess {
    # Display start banner
    Write-Host "`n===== BACKUP PROCESS STARTED $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====" -ForegroundColor Cyan

    # Load settings
    $settings = Get-BackupSettings
    
    # Get destination folder with validation
    # If no valid settings, ask the user
    if (-not $settings -or -not $settings.DestinationFolder) {
        do {
            $destinationFolder = Read-Host -Prompt "Please specify the destination folder for the backup"
            if (-not $destinationFolder) {
                Write-Host "Destination cannot be empty. Please try again." -ForegroundColor Red
            }
        } while (-not $destinationFolder)
        
        # Define log file extensions
        Write-Host "`nPlease specify which file extensions should be considered as log files (comma separated):" -ForegroundColor Yellow
        $logExtensions = (Read-Host -Prompt "For example: *.log,*.txt,*.trace").Split(',').Trim()
        
        if (-not $logExtensions -or $logExtensions.Count -eq 0 -or $logExtensions[0] -eq '') {
            $logExtensions = @("*.log", "*.txt")
            Write-Host "Using default log extensions: $($logExtensions -join ', ')" -ForegroundColor Yellow
        }
        
        # Ask for default log action
        Write-Host "`nWhat's the default action for log files?" -ForegroundColor Yellow
        Write-Host "1. Include logs in backup (default)" -ForegroundColor Green
        Write-Host "2. Exclude logs from backup" -ForegroundColor Yellow
        Write-Host "3. Delete logs after backup" -ForegroundColor Red
        
        $logAction = Read-Host -Prompt "Choose an option (1-3)"
        
        switch ($logAction) {
            "2" { $defaultLogAction = "exclude" }
            "3" { $defaultLogAction = "delete" }
            default { $defaultLogAction = "include" }
        }

        # Save the new settings
        Save-BackupSettings -destinationFolder $destinationFolder -logExtensions $logExtensions -defaultLogAction $defaultLogAction
    } else {
        $destinationFolder = $settings.DestinationFolder
        
        if ($settings.LogExtensions) {
            $logExtensions = $settings.LogExtensions
        } else {
            $logExtensions = @("*.log", "*.txt")
        }
        
        if ($settings.DefaultLogAction) {
            $defaultLogAction = $settings.DefaultLogAction
        } else {
            $defaultLogAction = "include"
        }
        
        Write-Host "Using destination folder from settings: $destinationFolder" -ForegroundColor Green
        Write-Host "Using log extensions from settings: $($logExtensions -join ', ')" -ForegroundColor Green
        Write-Host "Using default log action from settings: $defaultLogAction" -ForegroundColor Green
    }

    $name = $destinationFolder
    
    # Create timestamped backup folder
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $DestinationLog = Join-Path -Path $name -ChildPath "backup_$timestamp"
    $DestinationLog = $DestinationLog -replace '\\$', ''
    
    if (-not (Test-Path $DestinationLog)) {
        try {
            $null = New-Item -Type Directory -Path $DestinationLog -Force
            Write-Host "Created backup directory: $DestinationLog" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to create directory: $DestinationLog" -ForegroundColor Red
            Write-Host "Error: $_" -ForegroundColor Red
            exit 1
        }
    }
    
    # Check if CSV file exists
    $csvPath = "Backup-folder.csv"
    if (-not (Test-Path $csvPath)) {
        Write-Host "Error: Backup-folder.csv not found in the current directory." -ForegroundColor Red
        exit 1
    }
    
    # Check CSV structure and add LogAction column if needed
    $csvData = Import-Csv $csvPath
    $firstRow = $csvData | Select-Object -First 1
    
    if (-not ($firstRow.PSObject.Properties.Name -contains "LogAction")) {
        # Add LogAction column with default value
        $csvData | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "LogAction" -Value $defaultLogAction
        }
        
        # Save updated CSV
        $csvData | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Added LogAction column to CSV file with default value: $defaultLogAction" -ForegroundColor Yellow
    }
    
    # Create array for storing backup data
    $dataColl = @()
    $totalFolders = 0
    Import-Csv $csvPath | ForEach-Object {
        $totalFolders = $totalFolders + 1
    }
    
    $currentFolder = 0
    $successCount = 0
    $failureCount = 0
    
    # Create log file
    $logFile = Join-Path -Path $DestinationLog -ChildPath "backup_log_$timestamp.txt"
    "Backup started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $logFile
    
    # Process each folder from CSV
    Import-Csv $csvPath | ForEach-Object {
        $currentFolder++
        $folderName = $_.name
        $folderApp = $_.app -replace '\\$', ''
        $logAction = if ($_.LogAction) { $_.LogAction } else { $defaultLogAction }
        
        Write-Progress -Activity "Processing Backups" -Status "Folder $currentFolder of $totalFolders" -PercentComplete (($currentFolder / $totalFolders) * 100)
        
        try {
            # Validate folder exists
            if (-not (Test-Path $folderApp)) {
                throw "Source folder not found: $folderApp"
            }
            
            # Calculate folder size
            $folderSize = Get-FolderSize -Path $folderApp
            
            # Check for log files
            $logInfo = Get-LogFilesInfo -Path $folderApp -LogExtensions $logExtensions
            
            # Create data object for reporting
            $dataObject = [PSCustomObject]@{
                FolderName = $folderName
                Application = $folderApp
                FolderSize = $folderSize
                LogCount = $logInfo.Count
                LogSize = $logInfo.Size
                LogAction = $logAction
                BackupTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Status = "Success"
            }
            
            # Display progress
            Write-Host "[$currentFolder/$totalFolders] Processing: $folderApp" -NoNewline
            Write-Host " - Size: $folderSize" -ForegroundColor Green
            
            if ($logInfo.Count -gt 0) {
                Write-Host "  Found $($logInfo.Count) log file(s) - Size: $($logInfo.Size)" -ForegroundColor Cyan
                
                # Ask user about log action if not specified in CSV
                if (-not $_.LogAction) {
                    Write-Host "  Log action options:" -ForegroundColor Yellow
                    Write-Host "    1. Include logs in backup (default)" -ForegroundColor Green
                    Write-Host "    2. Exclude logs from backup" -ForegroundColor Yellow
                    Write-Host "    3. Delete logs after backup" -ForegroundColor Red
                    Write-Host "    4. Use default setting ($defaultLogAction)" -ForegroundColor Blue
                    
                    $action = Read-Host -Prompt "  Choose an option (1-4)"
                    
                    switch ($action) {
                        "1" { $logAction = "include" }
                        "2" { $logAction = "exclude" }
                        "3" { $logAction = "delete" }
                        default { $logAction = $defaultLogAction }
                    }
                    
                    $dataObject.LogAction = $logAction
                }
                
                # Report action
                Write-Host "  Log action: $logAction" -ForegroundColor $(
                    switch ($logAction) {
                        "include" { "Green" }
                        "exclude" { "Yellow" }
                        "delete" { "Red" }
                        default { "White" }
                    }
                )
            }
            
            # Create the zip archive
            $excludeFiles = @()
            if ($logAction -eq "exclude" -and $logInfo.Count -gt 0) {
                $excludeFiles = $logInfo.Files.FullName
            }

            
            $archiveResult = New-BackupArchive -SourcePath $folderApp -DestinationPath $DestinationLog -ArchiveName $folderName -ExcludeFiles $excludeFiles
            

            if ($archiveResult) {
                Write-Host "  Archive created: $DestinationLog\$folderName.zip" -ForegroundColor Yellow
                $successCount++
                
                # Delete log files if requested
                if ($logAction -eq "delete" -and $logInfo.Count -gt 0) {
                    try {
                        $logInfo.Files | ForEach-Object {
                            Remove-Item -Path $_.FullName -Force
                        }
                        Write-Host "  Deleted $($logInfo.Count) log file(s)" -ForegroundColor Red
                        $dataObject | Add-Member -MemberType NoteProperty -Name "LogsDeleted" -Value $true
                    }
                    catch {
                        Write-Host "  Failed to delete some log files: $_" -ForegroundColor Red
                        $dataObject | Add-Member -MemberType NoteProperty -Name "LogsDeleted" -Value $false
                    }
                }
                
                # Verify zip file was created and has content
                $zipPath = "$DestinationLog\$folderName.zip"
                if (Test-Path $zipPath) {
                    $zipSize = "{0:N2} MB" -f ((Get-Item $zipPath).Length / 1MB)
                    Write-Host "  Archive size: $zipSize" -ForegroundColor Cyan
                    $dataObject | Add-Member -MemberType NoteProperty -Name "ZipSize" -Value $zipSize
                }
            } else {
                $dataObject.Status = "Failed"
                $failureCount++
            }
            
            # Add to collection
            $dataColl += $dataObject
            
            # Add to log file
            "$($dataObject.BackupTime) - $folderApp - $folderSize - LogFiles: $($logInfo.Count) - LogAction: $logAction - Status: $($dataObject.Status)" | Out-File -FilePath $logFile -Append
        }
        catch {
            # Log error
            $errorMessage = $_.Exception.Message
            Write-Host "    Error processing folder: $folderName" -ForegroundColor Red
            Write-Host "    $errorMessage" -ForegroundColor Red
            
            $dataObject = [PSCustomObject]@{
                FolderName = $folderName
                Application = $folderApp
                FolderSize = "N/A"
                LogCount = 0
                LogSize = "N/A"
                LogAction = $logAction
                BackupTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Status = "Failed"
                Error = $errorMessage
            }
            
            $dataColl += $dataObject
            $failureCount++
            
            # Add to log file
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - ERROR - $folderApp - $errorMessage" | Out-File -FilePath $logFile -Append
        }
    }
    
    # Complete the progress bar
    Write-Progress -Activity "Processing Backups" -Completed
    
    # Export the data collection to CSV
    $reportPath = Join-Path -Path $DestinationLog -ChildPath "backup_report_$timestamp.csv"
    $dataColl | Export-Csv -Path $reportPath -NoTypeInformation
    
    # Display backup summary
    Write-Host "`n===== BACKUP SUMMARY =====" -ForegroundColor Cyan
    Write-Host "Total folders processed: $totalFolders" -ForegroundColor White
    Write-Host "Successful backups: $successCount" -ForegroundColor Green
    Write-Host "Failed backups: $failureCount" -ForegroundColor $(if ($failureCount -gt 0) {"Red"} else {"Green"})
    Write-Host "Backup location: $DestinationLog" -ForegroundColor Yellow
    Write-Host "Report saved to: $reportPath" -ForegroundColor Yellow
    Write-Host "Log saved to: $logFile" -ForegroundColor Yellow
    Write-Host "===== BACKUP COMPLETED $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====" -ForegroundColor Cyan
    
    # If we have data, show a summary
    if ($dataColl.Count -gt 0) {
        try {
            $dataColl | Format-Table -Property Application, FolderSize, LogCount, LogAction, Status -AutoSize
        }
        catch {
            Write-Host "Could not display summary table: $_" -ForegroundColor Red
        }
    }
    
    # Pause to show results
    Write-Host "`nPress any key to exit..." -ForegroundColor Magenta
    Read-Host | Out-Null
}

# Start the backup process
Start-BackupProcess
