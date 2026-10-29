<#
.SYNOPSIS
    Audits a self-hosted Snipe-IT installation for any possible telemetry / data-gathering.

.PARAMETER Path
    Root path to the Snipe-IT source code (where composer.json lives).

.EXAMPLE
    .\Check-SnipeIT-Telemetry.ps1 -Path "C:\inetpub\snipe-it"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Path
)

Write-Host "=== Snipe-IT Telemetry & Network Audit ===" -ForegroundColor Cyan
Write-Host "Scanning path: $Path`n"

# region --- 1. Search source for telemetry keywords --------------------------
$keywords = @(
    'sentry','segment','mixpanel','amplitude','google-analytics','gtag','telemetry',
    'analytics','tracking','phone-home','phone.home','usage','stats','reporting'
)

Write-Host "[1] Searching source for telemetry / analytics keywords..."
$matches = @()
foreach ($k in $keywords) {
    $results = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
               Select-String -Pattern $k -SimpleMatch
    if ($results) { $matches += $results }
}
if ($matches) {
    Write-Host "Potential telemetry references found:`n" -ForegroundColor Yellow
    $matches | Select-Object FileName,LineNumber,Line | Format-Table -AutoSize
} else {
    Write-Host "No telemetry keywords found in source." -ForegroundColor Green
}

# endregion -------------------------------------------------------------------

# region --- 2. Search for hardcoded external HTTP/HTTPS URLs -----------------
Write-Host "`n[2] Searching for hardcoded external HTTP/HTTPS URLs..."
$urls = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue |
        Select-String -Pattern 'https?://[a-zA-Z0-9._:/?&=%-]+' 
if ($urls) {
    Write-Host "External URLs detected:`n" -ForegroundColor Yellow
    $urls | Select-Object FileName,LineNumber,Line | Format-Table -AutoSize
} else {
    Write-Host "No external URLs found in code." -ForegroundColor Green
}
# endregion -------------------------------------------------------------------

# region --- 3. Check for telemetry env vars in .env or config ----------------
Write-Host "`n[3] Checking .env and config for telemetry-related variables..."
$envFiles = Get-ChildItem -Path $Path -Recurse -Include ".env*", "config*.php","config*.js" -ErrorAction SilentlyContinue
$envHits = $envFiles | Select-String -Pattern 'SENTRY|ANALYTICS|TELEMETRY|UPDATE|CHECK_FOR_UPDATES|DSN'
if ($envHits) {
    Write-Host "Possible telemetry-related env/config entries:`n" -ForegroundColor Yellow
    $envHits | Select-Object Path,LineNumber,Line | Format-Table -AutoSize
} else {
    Write-Host "No telemetry-related env/config variables found." -ForegroundColor Green
}
# endregion -------------------------------------------------------------------

# region --- 4. Inspect composer files for telemetry dependencies -------------
Write-Host "`n[4] Inspecting composer.json / composer.lock for telemetry packages..."
$composerFiles = Get-ChildItem -Path $Path -Include "composer.json","composer.lock" -ErrorAction SilentlyContinue
if ($composerFiles) {
    $composerHits = $composerFiles | Select-String -Pattern 'sentry|segment|mixpanel|amplitude|google|telemetry|analytics'
    if ($composerHits) {
        Write-Host "Suspicious dependencies:`n" -ForegroundColor Yellow
        $composerHits | Select-Object Path,LineNumber,Line | Format-Table -AutoSize
    } else {
        Write-Host "No telemetry dependencies detected in composer files." -ForegroundColor Green
    }
} else {
    Write-Host "composer.json not found under $Path" -ForegroundColor Red
}
# endregion -------------------------------------------------------------------

# region --- 5. Optional live outbound connection check -----------------------
Write-Host "`n[5] Checking for active outbound Internet connections..."
$processes = @('php','apache','nginx','httpd','node')
$net = Get-NetTCPConnection -State Established |
       Where-Object { $_.RemoteAddress -notmatch '^(127\.|10\.|172\.1[6-9]\.|172\.2[0-9]\.|172\.3[0-1]\.|192\.168\.)' }
if ($net) {
    Write-Host "Active external connections found:`n" -ForegroundColor Yellow
    $net | Select-Object OwningProcess,LocalAddress,LocalPort,RemoteAddress,RemotePort |
          Sort-Object OwningProcess | Format-Table -AutoSize
    foreach ($p in $processes) {
        Get-Process -Name $p -ErrorAction SilentlyContinue | 
        ForEach-Object { Write-Host "  Process: $($_.Name) PID=$($_.Id)" }
    }
} else {
    Write-Host "No established outbound Internet connections from this host." -ForegroundColor Green
}
# endregion -------------------------------------------------------------------

Write-Host "`n=== Audit complete ===" -ForegroundColor Cyan
