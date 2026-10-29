# Define file paths
$InputXLSX     = "D:\Scripts\devices.xlsx"
$OutputXLSX    = "D:\Scripts\Inventory_Report.xlsx"
$ThrottleLimit = 20  # Max concurrent threads

# Helper functions shared with runspaces via InitialSessionState
$HelperFunctions = {
    function Get-RoundedRAM ($RawGB) {
        $Ceil = [math]::Ceiling($RawGB)
        $p = 1; while ($p -lt $Ceil) { $p *= 2 }
        return $p
    }

    function Get-RoundedDisk ($RawGB) {
        $thresholds = @(120,240,256,480,512,960,1024,2048,4096)
        foreach ($t in $thresholds) { if ($RawGB -le ($t * 1.05)) { return $t } }
        return [math]::Ceiling($RawGB / 1024) * 1024
    }

    function Format-DiskSize ($GB) {
        if ($GB -ge 1024) { return "$([int]($GB / 1024)) TB" } else { return "$GB GB" }
    }

    function Get-DiskType ($CimDisk, $DiskModel) {
        switch ($CimDisk.MediaType) {
            3 { return "HDD" } 4 { return "SSD" } 5 { return "SCM" }
            default {
                if ($DiskModel -match "SSD|NVMe|Solid|Hynix|Samsung|Micron|Kingston|WD.*Green|WD.*Blue SN|Crucial") { return "SSD" } else { return "HDD" }
            }
        }
    }
}

# Dot-source helpers in the main thread
. $HelperFunctions

# Scriptblock executed by each runspace thread
$WorkerScript = {
    param($PC, $ExistingRecord, $Mode)

    # Re-define helpers inside the runspace
    function Get-RoundedRAM ($RawGB) {
        $Ceil = [math]::Ceiling($RawGB)
        $p = 1; while ($p -lt $Ceil) { $p *= 2 }
        return $p
    }
    function Get-RoundedDisk ($RawGB) {
        $thresholds = @(120,240,256,480,512,960,1024,2048,4096)
        foreach ($t in $thresholds) { if ($RawGB -le ($t * 1.05)) { return $t } }
        return [math]::Ceiling($RawGB / 1024) * 1024
    }
    function Format-DiskSize ($GB) {
        if ($GB -ge 1024) { return "$([int]($GB / 1024)) TB" } else { return "$GB GB" }
    }
    function Get-DiskType ($CimDisk, $DiskModel) {
        switch ($CimDisk.MediaType) {
            3 { return "HDD" } 4 { return "SSD" } 5 { return "SCM" }
            default {
                if ($DiskModel -match "SSD|NVMe|Solid|Hynix|Samsung|Micron|Kingston|WD.*Green|WD.*Blue SN|Crucial") { return "SSD" } else { return "HDD" }
            }
        }
    }

    if ($Mode -eq "DiskOnly") {
        try {
            $Disk    = Get-WmiObject -Class Win32_DiskDrive -ComputerName $PC -ErrorAction Stop | Select-Object -First 1
            $CimDisk = Get-CimInstance -ClassName MSFT_PhysicalDisk -Namespace "root\Microsoft\Windows\Storage" -ComputerName $PC -ErrorAction SilentlyContinue | Select-Object -First 1
            $DiskGB  = Get-RoundedDisk ([math]::Round($Disk.Size / 1GB, 0))
            return [PSCustomObject]@{
                ComputerName = $ExistingRecord.ComputerName
                Model        = $ExistingRecord.Model
                SerialNumber = $ExistingRecord.SerialNumber
                CPU          = $ExistingRecord.CPU
                RAM_GB       = $ExistingRecord.RAM_GB
                DiskSize     = Format-DiskSize $DiskGB
                DiskType     = Get-DiskType $CimDisk $Disk.Model
                IPAddress    = $ExistingRecord.IPAddress
                MACAddress   = $ExistingRecord.MACAddress
                LoggedUser   = $ExistingRecord.LoggedUser
                Status       = $ExistingRecord.Status
                Log          = " -> $PC Disk updated!"
                LogColor     = "Green"
            }
        } catch {
            return [PSCustomObject]@{ ComputerName = $PC; Status = $null; Log = " -> $PC Failed to get disk info: $($_.Exception.Message)"; LogColor = "Red" }
        }
    } else {
        try {
            $Sys     = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $PC -ErrorAction Stop
            $BIOS    = Get-WmiObject -Class Win32_BIOS -ComputerName $PC -ErrorAction Stop
            $CPU     = Get-WmiObject -Class Win32_Processor -ComputerName $PC -ErrorAction Stop | Select-Object -First 1
            $Net     = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $PC -Filter "IPEnabled='True'" -ErrorAction Stop | Select-Object -First 1
            $Disk    = Get-WmiObject -Class Win32_DiskDrive -ComputerName $PC -ErrorAction Stop | Select-Object -First 1
            $CimDisk = Get-CimInstance -ClassName MSFT_PhysicalDisk -Namespace "root\Microsoft\Windows\Storage" -ComputerName $PC -ErrorAction SilentlyContinue | Select-Object -First 1
            $DiskGB  = Get-RoundedDisk ([math]::Round($Disk.Size / 1GB, 0))
            return [PSCustomObject]@{
                ComputerName = $Sys.DNSHostName
                Model        = $Sys.Model
                SerialNumber = $BIOS.SerialNumber
                CPU          = $CPU.Name
                RAM_GB       = Get-RoundedRAM ([math]::Round($Sys.TotalPhysicalMemory / 1GB, 2))
                DiskSize     = Format-DiskSize $DiskGB
                DiskType     = Get-DiskType $CimDisk $Disk.Model
                IPAddress    = $Net.IPAddress[0]
                MACAddress   = $Net.MACAddress
                LoggedUser   = $Sys.UserName
                Status       = "Success"
                Log          = " -> $PC Success!"
                LogColor     = "Green"
            }
        } catch {
            return [PSCustomObject]@{
                ComputerName = $PC; Model = "N/A"; SerialNumber = "N/A"; CPU = "N/A"; RAM_GB = "N/A"
                DiskSize = "N/A"; DiskType = "N/A"; IPAddress = "N/A"; MACAddress = "N/A"
                LoggedUser = "N/A"; Status = "Failed or Offline"
                Log = " -> $PC Failed: $($_.Exception.Message)"; LogColor = "Red"
            }
        }
    }
}

# Read input and load existing results into a hashtable keyed by ComputerName
$Computers = Import-Excel -Path $InputXLSX
$ExistingResults = @{}
if (Test-Path $OutputXLSX) {
    try {
        $Loaded = Import-Excel -Path $OutputXLSX -WorksheetName "Inventory"
        foreach ($row in $Loaded) { $ExistingResults[$row.ComputerName] = $row }
        Write-Host "Loaded existing inventory ($($ExistingResults.Count) records)." -ForegroundColor Yellow
    } catch {
        Write-Host "Existing report has schema issues. Starting fresh..." -ForegroundColor Yellow
    }
}

# Build the job list, skipping already-complete devices
$Jobs = [System.Collections.Generic.List[hashtable]]@()
foreach ($Device in $Computers) {
    $PC = $Device.ComputerName
    if ([string]::IsNullOrWhiteSpace($PC)) {
        Write-Host "Skipping row with empty ComputerName. Available columns: $($Device.PSObject.Properties.Name -join ', ')" -ForegroundColor Yellow
        continue
    }
    $Existing = if ($ExistingResults.ContainsKey($PC)) { $ExistingResults[$PC] } else { $null }

    if ($Existing -and $Existing.Status -eq "Success" -and $Existing.DiskSize -and $Existing.DiskSize -ne "N/A") {
        Write-Host "Skipping $PC -> Already successful." -ForegroundColor DarkGray
    } elseif ($Existing -and $Existing.Status -eq "Success" -and ($null -eq $Existing.DiskSize -or $Existing.DiskSize -eq "N/A")) {
        $Jobs.Add(@{ PC = $PC; Mode = "DiskOnly"; Existing = $Existing })
    } else {
        $Jobs.Add(@{ PC = $PC; Mode = "Full"; Existing = $null })
    }
}

# Create RunspacePool and dispatch all jobs
$Pool = [RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit)
$Pool.Open()

$RunningJobs = foreach ($Job in $Jobs) {
    $PS = [PowerShell]::Create()
    $PS.RunspacePool = $Pool
    [void]$PS.AddScript($WorkerScript)
    [void]$PS.AddParameter("PC", $Job.PC)
    [void]$PS.AddParameter("ExistingRecord", $Job.Existing)
    [void]$PS.AddParameter("Mode", $Job.Mode)
    @{ PS = $PS; Handle = $PS.BeginInvoke() }
}

# Collect results as jobs complete
foreach ($Job in $RunningJobs) {
    $Result = $Job.PS.EndInvoke($Job.Handle) | Select-Object -First 1
    $Job.PS.Dispose()
    if ($null -ne $Result) {
        Write-Host $Result.Log -ForegroundColor $Result.LogColor
        if ($null -ne $Result.Status) { $ExistingResults[$Result.ComputerName] = $Result }
    }
}

$Pool.Close()
$Pool.Dispose()

# Strip internal log fields before exporting
$ExistingResults.Values | Select-Object ComputerName, Model, SerialNumber, CPU, RAM_GB, DiskSize, DiskType, IPAddress, MACAddress, LoggedUser, Status |
    Export-Excel -Path $OutputXLSX -FreezeTopRow -BoldTopRow -WorksheetName "Inventory"

Write-Host "Inventory complete! Saved to $OutputXLSX" -ForegroundColor Yellow
