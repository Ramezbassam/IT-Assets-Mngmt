# Load Excel and read device names from Column A, Sheet 1
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open("D:\Scripts\Sync-Agent\Test.xlsx")
$sheet = $workbook.Sheets.Item(1)

$devices = @()
$row = 1
while ($sheet.Cells.Item($row, 1).Value2) {
    $devices += $sheet.Cells.Item($row, 1).Value2
    $row++
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# Read local script content
$scriptContent = Get-Content "D:\Scripts\Sync-Agent\Snipe-IT_Agent.ps1" -Raw

# Force-enable WinRM on remote device via sc.exe
function Enable-WinRM {
    param($device)
    Write-Host "[$device] Attempting to enable WinRM via sc.exe..." -ForegroundColor Yellow
    sc.exe \\$device config WinRM start= auto | Out-Null
    sc.exe \\$device start WinRM | Out-Null
    Start-Sleep -Seconds 5
}

# Execute script on each remote device
foreach ($device in $devices) {
    Write-Host "Connecting to $device..."

    # Check WinRM reachability, enable it if not responding
    if (-not (Test-WSMan -ComputerName $device -ErrorAction SilentlyContinue)) {
        Enable-WinRM -device $device
    }

    try {
        $session = New-PSSession -ComputerName $device -ErrorAction Stop
        Invoke-Command -Session $session -ScriptBlock {
            param($script)
            Invoke-Expression "& { $script }"
        } -ArgumentList $scriptContent -ErrorAction Stop
        Write-Host "[$device] Script executed successfully." -ForegroundColor Green
    } catch {
        Write-Host "[$device] Failed: $_" -ForegroundColor Red
    } finally {
        if ($session) { Remove-PSSession $session }
    }
}
