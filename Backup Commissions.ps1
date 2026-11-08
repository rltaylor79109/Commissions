# Make backup copy of develmental version of "Training Plan 2.xlsm"

# Define source file and destination directory
$sourceFile = "C:\Users\User\source\repos\Commissions\Commissions (development).xlsm"
$backupDir  = "C:\Users\User\source\repos\Commissions\Backups\"

# Ensure the backup directory exists
if (-not (Test-Path -Path $backupDir)) {
    New-Item -ItemType Directory -Path $backupDir | Out-Null
}

# Generate timestamp (Format: YYYYMMDD_HHMMSS)
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Extract filename parts
$fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($sourceFile)
$extension          = [System.IO.Path]::GetExtension($sourceFile)

# Construct new backup filename
$backupFileName = "${fileNameWithoutExt}_${timestamp}${extension}"
$destinationPath = Join-Path -Path $backupDir -ChildPath $backupFileName

# Copy the file
Copy-Item -Path $sourceFile -Destination $destinationPath -Force

Write-Host "Backup created successfully at: $destinationPath" -ForegroundColor Green