# ==============================================================================
# PREREQUISITE: You must install the ImportExcel module if you haven't already.
# Run this once in an elevated PowerShell window: Install-Module -Name ImportExcel
# ==============================================================================

# --- Configuration Variables ---
$CsvFilePath   = "D:\Sync-Agent\Devices.csv"
$CsvColumnName = "DeviceName" # The header name in your Excel sheet containing the PCs

# Directory and Script Configuration
$LocalDirectory  = "D:\Sync-Agent" # The entire folder to be copied
$MainScriptName  = "Snipe-IT_Agent.ps1"                        # The name of the script INSIDE that folder to execute
$RemoteDrive     = "D:\"

# Calculate the remote paths dynamically
$FolderName      = Split-Path $LocalDirectory -Leaf
$RemoteDirectory = Join-Path -Path $RemoteDrive -ChildPath $FolderName
$RemoteScriptPath= Join-Path -Path $RemoteDirectory -ChildPath $MainScriptName

# --- 1. Fetch device names from Excel ---
Write-Host "Reading device list from Excel..." -ForegroundColor Cyan
try {
    $Devices = Import-Csv -Path $CsvFilePath | Select-Object -ExpandProperty $CsvColumnName
}
catch {
    Write-Error "Failed to read the Excel file. Please ensure the ImportExcel module is installed and the path is correct."
    Exit
}

# --- Loop through each device ---
foreach ($Device in $Devices) {
    if ([string]::IsNullOrWhiteSpace($Device)) { continue }

    Write-Host "`n========================================"
    Write-Host "Targeting Device: ${Device}" -ForegroundColor Yellow

    $Session = $null

    try {
        # --- 2. Open a session with the remote device ---
        Write-Host "  -> Opening PSSession..."
        $Session = New-PSSession -ComputerName $Device -ErrorAction Stop

        # --- 3. Copy the entire directory to Drive D ---
        Write-Host "  -> Copying directory '${FolderName}' to ${RemoteDrive}..."
        # Using -Recurse to ensure all files and subfolders are copied
        Copy-Item -Path $LocalDirectory -Destination $RemoteDrive -ToSession $Session -Recurse -Force -ErrorAction Stop

        # --- 4. Run the script from inside the copied directory ---
        Write-Host "  -> Executing ${MainScriptName} (Bypass Mode)..."
        Invoke-Command -Session $Session -ScriptBlock {
            param ($ScriptToRun, $WorkingDir)
            
            # Set location to the new directory in case the script relies on relative paths
            Set-Location -Path $WorkingDir

            # Launch a new powershell instance locally on the remote machine to bypass policy
            & powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File $ScriptToRun
            
        } -ArgumentList $RemoteScriptPath, $RemoteDirectory -ErrorAction Stop
        Write-Host "  -> Execution complete." -ForegroundColor Green

        # --- 5. Delete the entire directory after it finishes ---
        Write-Host "  -> Cleaning up directory ${RemoteDirectory}..."
        Invoke-Command -Session $Session -ScriptBlock {
            param ($DirToDelete)
            
            # Set location back to C:\ to avoid "file in use" locks before deleting
            Set-Location -Path "C:\" 
            
            if (Test-Path $DirToDelete) {
                # Using -Recurse and -Force to delete the folder and all its contents
                Remove-Item -Path $DirToDelete -Recurse -Force -ErrorAction Stop
            }
        } -ArgumentList $RemoteDirectory -ErrorAction Stop
        Write-Host "  -> Cleanup successful." -ForegroundColor Green

    }
    catch {
        Write-Host "  [X] Error processing ${Device}: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        # --- 6. Close the session ---
        if ($Session) {
            Write-Host "  -> Closing PSSession..."
            Remove-PSSession -Session $Session
        }
    }
}

Write-Host "`nAll operations completed." -ForegroundColor Cyan