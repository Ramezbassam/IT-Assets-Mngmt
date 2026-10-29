<#
    Device Inventory Script

    Features:
    - Reads device names from an input Excel file.
    - Queries devices remotely by ComputerName.
    - Collects:
        ComputerName
        Model
        SerialNumber
        CPU
        RAM_GB
        DiskSize_GB
        DiskType
        IPAddress
        MACAddress
        LoggedUser
        Status
    - Rounds RAM and disk sizes before writing them.
    - Existing records are matched by ComputerName.
    - Successful records with complete disk information are skipped.
    - Successful records missing disk information are repaired.
    - Failed records are updated instead of duplicated.
    - Handles empty/missing existing Excel reports safely.
#>

# ============================================================
# CONFIGURATION
# ============================================================

$InputXLSX  = "D:\Scripts\devices.xlsx"
$OutputXLSX = "D:\Scripts\Inventory_Report.xlsx"

# Column in the input Excel file containing the hostname.
$DeviceNameColumn = "ComputerName"

# Rounding steps in GB.
$RAMRoundStep  = 1
$DiskRoundStep = 50

# ============================================================
# VALIDATE DEPENDENCIES
# ============================================================

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ERROR: ImportExcel module is not installed." -ForegroundColor Red
    Write-Host "Install it with:" -ForegroundColor Yellow
    Write-Host "Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

Import-Module ImportExcel -ErrorAction Stop

# ============================================================
# HELPER FUNCTIONS
# ============================================================

function Get-RoundedSize {
    param(
        [Parameter(Mandatory)]
        [double]$Value,

        [int]$Step = 1
    )

    if ($Step -le 0) {
        $Step = 1
    }

    return [int](
        [math]::Round(
            $Value / $Step,
            0,
            [System.MidpointRounding]::AwayFromZero
        ) * $Step
    )
}


function Get-DiskInfo {
    param(
        [Parameter(Mandatory)]
        [string]$DeviceName
    )

    try {
        # Get physical disk information.
        $Disk = Get-WmiObject `
            -Class Win32_DiskDrive `
            -ComputerName $DeviceName `
            -ErrorAction Stop |
            Select-Object -First 1

        if ($null -eq $Disk) {
            throw "No physical disk was found."
        }

        # Try to get modern Storage module information.
        $CimDisk = Get-CimInstance `
            -ClassName MSFT_PhysicalDisk `
            -Namespace "root\Microsoft\Windows\Storage" `
            -ComputerName $DeviceName `
            -ErrorAction SilentlyContinue |
            Select-Object -First 1

        # Calculate disk size.
        $DiskSizeGB = Get-RoundedSize `
            -Value ($Disk.Size / 1GB) `
            -Step $DiskRoundStep

        # Determine disk type.
        if ($null -ne $CimDisk) {

            switch ($CimDisk.MediaType) {
                3 {
                    $DiskType = "HDD"
                    break
                }

                4 {
                    $DiskType = "SSD"
                    break
                }

                5 {
                    $DiskType = "SCM"
                    break
                }

                default {
                    $DiskType = $null
                }
            }
        }

        # Fall back to model-based detection.
        if ([string]::IsNullOrWhiteSpace($DiskType)) {

            $Model = [string]$Disk.Model

            if ($Model -match "SSD|NVMe|Solid|Hynix|Samsung|Micron|Kingston|WD.*Green|WD.*Blue SN|Crucial") {
                $DiskType = "SSD"
            }
            else {
                $DiskType = "HDD"
            }
        }

        return [PSCustomObject]@{
            DiskSize = $DiskSizeGB
            DiskType = $DiskType
        }
    }
    catch {
        throw "Disk query failed: $($_.Exception.Message)"
    }
}


function Set-ResultRecord {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.ArrayList]$Results,

        [Parameter(Mandatory)]
        [hashtable]$Index,

        [Parameter(Mandatory)]
        [string]$DeviceName,

        [Parameter(Mandatory)]
        [PSCustomObject]$Record
    )

    if ($Index.ContainsKey($DeviceName)) {

        $ExistingIndex = $Index[$DeviceName]

        $Results[$ExistingIndex] = $Record
    }
    else {

        [void]$Results.Add($Record)

        $Index[$DeviceName] = $Results.Count - 1
    }
}


function New-FailedRecord {
    param(
        [Parameter(Mandatory)]
        [string]$DeviceName
    )

    return [PSCustomObject]@{
        ComputerName = $DeviceName
        Model        = "N/A"
        SerialNumber = "N/A"
        CPU          = "N/A"
        RAM_GB       = "N/A"
        DiskSize_GB  = "N/A"
        DiskType     = "N/A"
        IPAddress    = "N/A"
        MACAddress   = "N/A"
        LoggedUser   = "N/A"
        Status       = "Failed or Offline"
    }
}


function Test-InventoryComplete {
    param(
        [Parameter(Mandatory)]
        $Record
    )

    if ($null -eq $Record) {
        return $false
    }

    if ($Record.Status -ne "Success") {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace([string]$Record.DiskSize_GB)) {
        return $false
    }

    if ([string]$Record.DiskSize_GB -eq "N/A") {
        return $false
    }

    return $true
}

# ============================================================
# VALIDATE INPUT FILE
# ============================================================

if (-not (Test-Path -LiteralPath $InputXLSX)) {
    Write-Host "ERROR: Input file not found:" -ForegroundColor Red
    Write-Host $InputXLSX -ForegroundColor Red
    exit 1
}

# ============================================================
# LOAD INPUT LIST
# ============================================================

try {

    $Computers = @(Import-Excel `
        -Path $InputXLSX `
        -ErrorAction Stop)

}
catch {

    Write-Host "ERROR: Failed to read input Excel file." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

if ($Computers.Count -eq 0) {
    Write-Host "ERROR: Input Excel file contains no records." -ForegroundColor Red
    exit 1
}

# Validate that the requested column exists.
$FirstComputer = $Computers | Select-Object -First 1

if ($null -eq $FirstComputer.PSObject.Properties[$DeviceNameColumn]) {

    Write-Host "ERROR: Column '$DeviceNameColumn' was not found in the input file." -ForegroundColor Red

    Write-Host "Available columns:" -ForegroundColor Yellow

    $FirstComputer.PSObject.Properties.Name |
        ForEach-Object {
            Write-Host " - $_" -ForegroundColor Yellow
        }

    exit 1
}

Write-Host "Loaded $($Computers.Count) devices from input file." -ForegroundColor Cyan

# ============================================================
# INITIALIZE RESULTS
# ============================================================

# IMPORTANT:
# These are initialized BEFORE attempting to load the existing
# Excel report. Therefore they can never accidentally remain null.

$Results = [System.Collections.ArrayList]::new()

$ResultIndex = @{}

# ============================================================
# LOAD EXISTING RESULTS
# ============================================================

if (Test-Path -LiteralPath $OutputXLSX) {

    Write-Host "Existing inventory report found." -ForegroundColor Yellow

    try {

        $ImportedResults = @(Import-Excel `
            -Path $OutputXLSX `
            -WorksheetName "Inventory" `
            -ErrorAction Stop)

        foreach ($Record in $ImportedResults) {

            if ($null -ne $Record) {
                [void]$Results.Add($Record)
            }
        }

        Write-Host "Loaded $($Results.Count) existing inventory records." -ForegroundColor Yellow
    }
    catch {

        Write-Host "Could not load existing inventory records." -ForegroundColor DarkYellow
        Write-Host "Reason: $($_.Exception.Message)" -ForegroundColor DarkYellow
        Write-Host "Starting with an empty inventory." -ForegroundColor Yellow
    }
}
else {

    Write-Host "No existing inventory report found." -ForegroundColor Yellow
    Write-Host "Starting with an empty inventory." -ForegroundColor Yellow
}

# ============================================================
# BUILD RESULT INDEX
# ============================================================

for ($i = 0; $i -lt $Results.Count; $i++) {

    $ExistingName = [string]$Results[$i].ComputerName

    if (-not [string]::IsNullOrWhiteSpace($ExistingName)) {

        $ExistingName = $ExistingName.Trim()

        # If duplicate ComputerNames somehow already exist,
        # keep the first one and don't overwrite its index.
        if (-not $ResultIndex.ContainsKey($ExistingName)) {
            $ResultIndex[$ExistingName] = $i
        }
    }
}

Write-Host "Indexed $($ResultIndex.Count) existing devices." -ForegroundColor DarkGray

# ============================================================
# INVENTORY COUNTERS
# ============================================================

$TotalDevices   = $Computers.Count
$Processed      = 0
$Successful     = 0
$Failed         = 0
$Skipped        = 0
$DiskUpdated    = 0

# ============================================================
# MAIN LOOP
# ============================================================

foreach ($Device in $Computers) {

    $Processed++

    # --------------------------------------------------------
    # Get device name
    # --------------------------------------------------------

    $DeviceName = [string]$Device.$DeviceNameColumn

    if ([string]::IsNullOrWhiteSpace($DeviceName)) {

        Write-Host ""
        Write-Host "[$Processed/$TotalDevices] Skipping row -> no device name found." `
            -ForegroundColor DarkYellow

        $Skipped++

        continue
    }

    $DeviceName = $DeviceName.Trim()

    Write-Host ""
    Write-Host "============================================================" `
        -ForegroundColor DarkGray

    Write-Host "[$Processed/$TotalDevices] Processing: $DeviceName" `
        -ForegroundColor Cyan

    # --------------------------------------------------------
    # Find existing record
    # --------------------------------------------------------

    $Existing = $null

    if ($ResultIndex.ContainsKey($DeviceName)) {

        $ExistingIndex = $ResultIndex[$DeviceName]

        if ($ExistingIndex -ge 0 -and $ExistingIndex -lt $Results.Count) {
            $Existing = $Results[$ExistingIndex]
        }
    }

    # --------------------------------------------------------
    # CASE 1:
    # Already successful and disk information exists.
    # --------------------------------------------------------

    if (Test-InventoryComplete -Record $Existing) {

        Write-Host " -> Already successfully inventoried. Skipping." `
            -ForegroundColor DarkGray

        $Skipped++

        continue
    }

    # --------------------------------------------------------
    # CASE 2:
    # Successful record exists but disk information is missing.
    # Only query the disk.
    # --------------------------------------------------------

    if (
        $null -ne $Existing -and
        $Existing.Status -eq "Success"
    ) {

        Write-Host " -> Existing record is successful but disk information is missing." `
            -ForegroundColor Cyan

        Write-Host " -> Querying disk information..." `
            -ForegroundColor Cyan

        try {

            $DiskInfo = Get-DiskInfo -DeviceName $DeviceName

            $UpdatedRecord = [PSCustomObject]@{
                ComputerName = $Existing.ComputerName
                Model        = $Existing.Model
                SerialNumber = $Existing.SerialNumber
                CPU          = $Existing.CPU
                RAM_GB       = $Existing.RAM_GB
                DiskSize_GB  = $DiskInfo.DiskSize
                DiskType     = $DiskInfo.DiskType
                IPAddress    = $Existing.IPAddress
                MACAddress   = $Existing.MACAddress
                LoggedUser   = $Existing.LoggedUser
                Status       = "Success"
            }

            Set-ResultRecord `
                -Results $Results `
                -Index $ResultIndex `
                -DeviceName $DeviceName `
                -Record $UpdatedRecord

            Write-Host " -> Disk information updated successfully." `
                -ForegroundColor Green

            $DiskUpdated++
            $Successful++
        }
        catch {

            Write-Host " -> Failed to update disk information:" `
                -ForegroundColor Red

            Write-Host " -> $($_.Exception.Message)" `
                -ForegroundColor Red

            # Keep existing successful record.
            # We do NOT overwrite it with a failed record because
            # the rest of the inventory is still valid.
        }

        continue
    }

    # --------------------------------------------------------
    # CASE 3:
    # New device or previously failed device.
    # Perform full inventory.
    # --------------------------------------------------------

    Write-Host " -> Querying device..." -ForegroundColor Cyan

    try {

        # ----------------------------------------------------
        # Computer system
        # ----------------------------------------------------

        $Sys = Get-WmiObject `
            -Class Win32_ComputerSystem `
            -ComputerName $DeviceName `
            -ErrorAction Stop

        # ----------------------------------------------------
        # BIOS
        # ----------------------------------------------------

        $BIOS = Get-WmiObject `
            -Class Win32_BIOS `
            -ComputerName $DeviceName `
            -ErrorAction Stop

        # ----------------------------------------------------
        # CPU
        # ----------------------------------------------------

        $CPU = Get-WmiObject `
            -Class Win32_Processor `
            -ComputerName $DeviceName `
            -ErrorAction Stop |
            Select-Object -First 1

        # ----------------------------------------------------
        # Network
        # ----------------------------------------------------

        $Net = Get-WmiObject `
            -Class Win32_NetworkAdapterConfiguration `
            -ComputerName $DeviceName `
            -Filter "IPEnabled='True'" `
            -ErrorAction Stop |
            Select-Object -First 1

        if ($null -eq $Net) {
            throw "No enabled network adapter was found."
        }

        # ----------------------------------------------------
        # Disk
        # ----------------------------------------------------

        $DiskInfo = Get-DiskInfo -DeviceName $DeviceName

        # ----------------------------------------------------
        # Safely extract IP address
        # ----------------------------------------------------

        $IPAddress = "N/A"

        if ($null -ne $Net.IPAddress) {

            $IPv4 = @(
                $Net.IPAddress |
                Where-Object {
                    $_ -match '^\d{1,3}(\.\d{1,3}){3}$'
                }
            )

            if ($IPv4.Count -gt 0) {
                $IPAddress = $IPv4[0]
            }
        }

        # ----------------------------------------------------
        # Safely extract MAC address
        # ----------------------------------------------------

        $MACAddress = if (
            [string]::IsNullOrWhiteSpace([string]$Net.MACAddress)
        ) {
            "N/A"
        }
        else {
            [string]$Net.MACAddress
        }

        # ----------------------------------------------------
        # CPU
        # ----------------------------------------------------

        $CPUName = if ($null -eq $CPU) {
            "N/A"
        }
        else {
            [string]$CPU.Name
        }

        # ----------------------------------------------------
        # Build successful record
        # ----------------------------------------------------

        $DeviceInfo = [PSCustomObject]@{
            ComputerName = $DeviceName
            Model        = if ($Sys.Model) { $Sys.Model } else { "N/A" }
            SerialNumber = if ($BIOS.SerialNumber) { $BIOS.SerialNumber } else { "N/A" }
            CPU          = $CPUName
            RAM_GB       = Get-RoundedSize `
                -Value ($Sys.TotalPhysicalMemory / 1GB) `
                -Step $RAMRoundStep
            DiskSize_GB  = $DiskInfo.DiskSize
            DiskType     = $DiskInfo.DiskType
            IPAddress    = $IPAddress
            MACAddress   = $MACAddress
            LoggedUser   = if ($Sys.UserName) { $Sys.UserName } else { "N/A" }
            Status       = "Success"
        }

        # ----------------------------------------------------
        # Add or update result
        # ----------------------------------------------------

        Set-ResultRecord `
            -Results $Results `
            -Index $ResultIndex `
            -DeviceName $DeviceName `
            -Record $DeviceInfo

        Write-Host " -> Success!" -ForegroundColor Green

        $Successful++
    }
    catch {

        Write-Host " -> Failed: $($_.Exception.Message)" `
            -ForegroundColor Red

        # ----------------------------------------------------
        # Store failure in results.
        # This means failed devices are visible in the report
        # and will be retried on the next execution.
        # ----------------------------------------------------

        $FailedRecord = New-FailedRecord -DeviceName $DeviceName

        Set-ResultRecord `
            -Results $Results `
            -Index $ResultIndex `
            -DeviceName $DeviceName `
            -Record $FailedRecord

        $Failed++
    }
}

# ============================================================
# EXPORT RESULTS
# ============================================================

Write-Host ""
Write-Host "============================================================" `
    -ForegroundColor DarkGray

Write-Host "Exporting inventory..." -ForegroundColor Cyan

try {

    # Export only if there are records.
    if ($Results.Count -eq 0) {

        Write-Host "WARNING: No inventory records were generated." `
            -ForegroundColor Yellow
    }
    else {

        $Results |
            Export-Excel `
                -Path $OutputXLSX `
                -FreezeTopRow `
                -BoldTopRow `
                -AutoSize `
                -WorksheetName "Inventory" `
                -ClearSheet `
                -ErrorAction Stop

        Write-Host ""
        Write-Host "Inventory complete!" -ForegroundColor Green
        Write-Host "Saved to: $OutputXLSX" -ForegroundColor Green
    }
}
catch {

    Write-Host ""
    Write-Host "ERROR: Failed to export inventory report." `
        -ForegroundColor Red

    Write-Host $_.Exception.Message -ForegroundColor Red

    exit 1
}

# ============================================================
# SUMMARY
# ============================================================

Write-Host ""
Write-Host "================ INVENTORY SUMMARY ==========================" `
    -ForegroundColor Cyan

Write-Host "Total devices       : $TotalDevices"
Write-Host "Successfully queried: $Successful"
Write-Host "Disk records updated: $DiskUpdated"
Write-Host "Failed              : $Failed"
Write-Host "Skipped             : $Skipped"
Write-Host "Records in report   : $($Results.Count)"

Write-Host "============================================================" `
    -ForegroundColor Cyan