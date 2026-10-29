# Deployed via Group Policy - runs locally on each device
# GPO: Computer Configuration > Policies > Windows Settings > Scripts > Startup
# Or: GPO Software/Script deployment pointing to this script on a network share

$NewPassword  = "FLXBioS2026"
$UNCLog       = "\\dc01\logs\BiosPassword\$env:COMPUTERNAME.log"
$LocalLog     = "$env:ProgramData\BiosPassword\$env:COMPUTERNAME.log"

function Write-Log {
    param($Message)
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $env:COMPUTERNAME | $Message"
    # Try UNC first, fall back to local
    try { Add-Content -Path $UNCLog -Value $entry -ErrorAction Stop }
    catch {
        $localDir = Split-Path $LocalLog
        if (-not (Test-Path $localDir)) { New-Item -ItemType Directory -Path $localDir -Force | Out-Null }
        Add-Content -Path $LocalLog -Value $entry
    }
    Write-Output $entry
}

$Manufacturer = (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer

try {
    if ($Manufacturer -match "Lenovo") {
        Import-Module Lenovo_BIOSSettings -ErrorAction Stop   # LVFS / Lenovo BIOS Settings module

        # Set supervisor password (SVP) - empty string = no current password
        Set-LenovoBiosSetting -SettingName "SVP" -Value $NewPassword -CurrentPassword "" -ErrorAction Stop
        Write-Log "Success | Lenovo"

    } elseif ($Manufacturer -match "Dell") {
        # Locate DellBIOSProvider regardless of Dell Command Configure install path
        $dellModule = Get-ChildItem "C:\Program Files\Dell","C:\Program Files (x86)\Dell" `
            -Recurse -Filter "DellBIOSProvider.psd1" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $dellModule) { throw "DellBIOSProvider module not found. Install Dell Command | Configure." }
        Import-Module $dellModule.FullName -ErrorAction Stop

        Set-Item -Path "DellSmbios:\Security\AdminPassword" -Value $NewPassword `
            -Password "" -ErrorAction Stop
        Write-Log "Success | Dell"

    } elseif ($Manufacturer -match "HP|Hewlett") {
        # Locate HP CMSL module regardless of install path
        $hpModule = Get-ChildItem "C:\Program Files\HP","C:\Program Files (x86)\HP" `
            -Recurse -Filter "HP.ClientManagement.psd1" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $hpModule) { throw "HP.ClientManagement module not found. Install HP CMSL." }
        Import-Module $hpModule.FullName -ErrorAction Stop

        Set-HPBIOSSetupPassword -NewPassword $NewPassword -CurrentPassword "" -ErrorAction Stop
        Write-Log "Success | HP"

    } else {
        Write-Log "Skipped | Unsupported manufacturer: $Manufacturer"
    }

} catch {
    Write-Log "Failed | $Manufacturer | $($_.Exception.Message)"
    exit 1
}
