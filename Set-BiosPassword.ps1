# Run with: powershell.exe -ExecutionPolicy Bypass -File Set-BiosPassword.ps1

$InputXLSX   = "D:\Scripts\devices.xlsx"
$NewPassword  = "Fl3x!EGypt$"
$LogFile      = "D:\Scripts\Set-BiosPassword_Log.csv"

# Add wildcard to TrustedHosts to allow IP-based WinRM connections
$currentTrusted = (Get-Item WSMan:\localhost\Client\TrustedHosts).Value
if ($currentTrusted -ne '*') {
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*' -Force
    Write-Host "TrustedHosts set to *." -ForegroundColor Yellow
}

$Creds    = Get-Credential -Message "Enter admin credentials for remote sessions"
$Devices  = Import-Excel -Path $InputXLSX
$Results  = [System.Collections.ArrayList]@()

function Set-LenovoBiosPassword {
    param($Session, $NewPass)
    $result = Invoke-Command -Session $Session -ScriptBlock {
        param($pass)
        try {
            # Format: "SettingName,NewValue,ascii,CurrentPassword" (empty current password)
            $set = Get-WmiObject -Namespace "root\WMI" -Class "Lenovo_SetBiosSetting" -ErrorAction Stop
            $set.SetBiosSetting("SVP,$pass,ascii,") | Out-Null

            # Save using Lenovo_SaveBiosSettings (plural)
            $save = Get-WmiObject -Namespace "root\WMI" -Class "Lenovo_SaveBiosSettings" -ErrorAction Stop
            $save.SaveBiosSettings("$pass,ascii") | Out-Null
            return "Success"
        } catch {
            return "Failed: $($_.Exception.Message)"
        }
    } -ArgumentList $NewPass
    return $result
}

function Set-DellBiosPassword {
    param($Session, $NewPass)
    $result = Invoke-Command -Session $Session -ScriptBlock {
        param($pass)
        try {
            # Dell Command Monitor must be installed for DCIM namespace
            # Fallback: use root\WMI Dell BIOS provider
            $ns = "root\WMI"
            $svc = Get-WmiObject -Namespace $ns -Class "BIOSAttributeInterface" -ErrorAction Stop
            $ret = $svc.SetBIOSAttributeInteger("AdminPwd", $pass, "")
            if ($ret.ReturnValue -eq 0) { return "Success" }
            else { return "Failed: ReturnValue=$($ret.ReturnValue)" }
        } catch {
            # Second attempt via Dell Command Monitor namespace
            try {
                $svc2 = Get-WmiObject -Namespace "root\dcim\sysman" -Class "Dell_BIOSPassword" -ErrorAction Stop | Where-Object { $_.AttributeName -eq "AdminPwd" }
                $svc2.SetPassword($pass, "") | Out-Null
                return "Success"
            } catch {
                return "Failed: $($_.Exception.Message)"
            }
        }
    } -ArgumentList $NewPass
    return $result
}

foreach ($Device in $Devices) {
    $IP = $Device.IPAddress
    Write-Host "Connecting to $IP..." -ForegroundColor Cyan

    $Status = $null
    $Manufacturer = "Unknown"

    try {
        $Session = New-PSSession -ComputerName $IP -Credential $Creds -ErrorAction Stop

        $Manufacturer = Invoke-Command -Session $Session -ScriptBlock {
            (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer
        }

        if ($Manufacturer -match "Lenovo") {
            $Status = Set-LenovoBiosPassword -Session $Session -NewPass $NewPassword
        } elseif ($Manufacturer -match "Dell") {
            $Status = Set-DellBiosPassword -Session $Session -NewPass $NewPassword
        } else {
            $Status = "Skipped: Unsupported manufacturer ($Manufacturer)"
        }

        Write-Host " -> $Status" -ForegroundColor $(if ($Status -eq "Success") { "Green" } elseif ($Status -like "Skipped*") { "Yellow" } else { "Red" })

    } catch {
        $Status = "Failed: $($_.Exception.Message)"
        Write-Host " -> $Status" -ForegroundColor Red
    } finally {
        if ($Session) { Remove-PSSession $Session }
    }

    [void]$Results.Add([PSCustomObject]@{
        IPAddress    = $IP
        Manufacturer = $Manufacturer
        Status       = $Status
    })
}

$Results | Export-Csv -Path $LogFile -NoTypeInformation -Force
Write-Host "`nDone! Log saved to $LogFile" -ForegroundColor Yellow
