# Enhanced PowerShell Backup Script
# Purpose: Automates the backup of folders listed in a CSV file to a compressed archive

# Set error action preference and enable verbose output
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

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

function New-BackupArchive {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ArchiveName
    )
    
    try {
        Write-Verbose "Creating archive: $DestinationPath\$ArchiveName.zip"
        Compress-Archive -Path $SourcePath -DestinationPath "$DestinationPath\$ArchiveName.zip" -CompressionLevel Optimal -Force
        return $true
    }
    catch {
        Write-Error "Failed to create archive: $_"
        return $false
    }
}

function Start-BackupProcess {
    # Display start banner
    Write-Host "`n===== BACKUP PROCESS STARTED $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') =====" -ForegroundColor Cyan
    
    # Get destination folder with validation
    do {
        $name = Read-Host -Prompt "Please specify the destination folder for the backup"
        if (-not $name) {
            Write-Host "Destination cannot be empty. Please try again." -ForegroundColor Red
        }
    } while (-not $name)
    
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
        
        Write-Progress -Activity "Processing Backups" -Status "Folder $currentFolder of $totalFolders" -PercentComplete (($currentFolder / $totalFolders) * 100)
        
        try {
            
            # Validate folder exists
            if (-not (Test-Path $folderApp)) {
                throw "Source folder not found: $folderApp"
            }
            
            # Calculate folder size
            $folderSize = Get-FolderSize -Path $folderApp
            
            # Create data object for reporting
            $dataObject = [PSCustomObject]@{
                FolderName = $folderName
                Application = $folderApp
                FolderSize = $folderSize
                BackupTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Status = "Success"
            }
            
            # Display progress
            Write-Host "[$currentFolder/$totalFolders] Processing: $folderApp" -NoNewline
            Write-Host " - Size: $folderSize" -ForegroundColor Green
            
            # Create the zip archive
            $archiveResult = New-BackupArchive -SourcePath $folderApp  -DestinationPath $DestinationLog -ArchiveName $folderName
            
            if ($archiveResult) {
                Write-Host "  ✓ Archive created: $DestinationLog\$folderApp.zip" -ForegroundColor Yellow
                $successCount++
                
                # Verify zip file was created and has content
                $zipPath = "$DestinationLog\$folderApp.zip"
                if (Test-Path $zipPath) {
                    $zipSize = "{0:N2} MB" -f ((Get-Item $zipPath).Length / 1MB)
                    Write-Host "  ℹ Archive size: $zipSize" -ForegroundColor Cyan
                    $dataObject | Add-Member -MemberType NoteProperty -Name "ZipSize" -Value $zipSize
                }
            } else {
                $dataObject.Status = "Failed"
                $failureCount++
            }
            
            # Add to collection
            $dataColl += $dataObject
            
            # Add to log file
            "$($dataObject.BackupTime) - $folderApp - $folderSize - Status: $($dataObject.Status)" | Out-File -FilePath $logFile -Append
        }
        catch {
            # Log error
            $errorMessage = $_.Exception.Message
            Write-Host "  ✗ Error processing folder: $folderName" -ForegroundColor Red
            Write-Host "    $errorMessage" -ForegroundColor Red
            
            $dataObject = [PSCustomObject]@{
                FolderName = $folderName
                Application = $folderApp
                FolderSize = "N/A"
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
            $dataColl | Format-Table -Property Application, FolderSize, Status -AutoSize
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

