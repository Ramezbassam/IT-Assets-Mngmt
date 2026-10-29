# =================================================================================
# SCRIPT: Sync-Assets_DropFolder.ps1
# VERSION: 1.2 (Drop Folder Version)
# AUTHOR: Gemini
#
# DESCRIPTION:
# Watches a specified "drop folder" for new Freshservice CSV exports. When a file
# is found, it is processed and synced with Snipe-IT, then moved to an archive
# folder. This script is intended to be run on a schedule.
#
# USAGE:
# 1. Manually download the CSV from your Freshservice email into the 'dropFolder'.
# 2. Run this script (or have a Scheduled Task run it).
# =================================================================================

#region Initialization and Configuration Loading
$ErrorActionPreference = 'Stop'
$scriptPath = $PSScriptRoot
$configPath = Join-Path $scriptPath "config.json"

if (-not (Test-Path $configPath)) {
    throw "CRITICAL: config.json not found in the script directory '$scriptPath'."
}

try {
    $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
} catch {
    throw "CRITICAL: Failed to parse config.json. Please ensure it is valid JSON. Error: $($_.Exception.Message)"
}

# Assign variables from config
$snipeItUrl = $config.snipeIt.apiUrl.TrimEnd('/')
$snipeItApiKey = $config.snipeIt.apiKey
$uniqueIdFieldSnipe = $config.scriptSettings.uniqueIdentifier.snipeItApiField
$uniqueIdFieldFreshservice = $config.scriptSettings.uniqueIdentifier.freshserviceCsvHeader
$fieldMapping = $config.fieldMapping
$dropFolder = $config.scriptSettings.dropFolder
$archiveFolder = $config.scriptSettings.archiveFolder

# Prepare Snipe-IT API headers
$script:snipeHeaders = @{ "Authorization" = "Bearer $snipeItApiKey"; "Accept" = "application/json"; "Content-Type" = "application/json" }
#endregion

#region Helper Functions
$snipeIdCache = @{ models = @{}; statuslabels = @{}; locations = @{}; users = @{}; departments = @{}; companies = @{}; categories = @{} }

function Get-SnipeItId {
    param([string]$Endpoint, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if ($snipeIdCache[$Endpoint].ContainsKey($Name)) { return $snipeIdCache[$Endpoint][$Name] }

    Write-Host "Fetching ID for $Endpoint named '$Name'..."
    # Try without search parameter first (get all and filter)
    $apiUrl = "$snipeItUrl/api/v1/$Endpoint?limit=100"
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Get -Headers $script:snipeHeaders
        Start-Sleep -Milliseconds 200
        if ($response.total -gt 0) {
            $foundNames = $response.rows | ForEach-Object { $_.name }
            Write-Host "DEBUG: Available $Endpoint names: $($foundNames -join ', ')"
            $id = ($response.rows | Where-Object { $_.name -ieq $Name } | Select-Object -First 1).id
            if ($id) {
                Write-Host "DEBUG: Found exact match with ID: $id"
                $snipeIdCache[$Endpoint][$Name] = $id
                return $id
            }
        }
    } catch {
        Write-Warning "DEBUG: API call failed: $($_.Exception.Message)"
        Start-Sleep -Milliseconds 500
    }
    
    # Only create for specific endpoints
    if ($Endpoint -eq 'statuslabels') {
        try {
            Write-Host "Creating status '$Name'..."
            $payload = @{ name = $Name; status_type = 'deployable' } | ConvertTo-Json
            $created = Invoke-RestMethod -Uri "$snipeItUrl/api/v1/statuslabels" -Method Post -Headers $script:snipeHeaders -Body $payload
            Start-Sleep -Milliseconds 300
            $snipeIdCache[$Endpoint][$Name] = $created.id
            return $created.id
        } catch {
            Write-Warning "Failed to create status '$Name': $($_.Exception.Message)"
        }
    }
    elseif ($Endpoint -eq 'locations') {
        try {
            Write-Host "Creating location '$Name'..."
            $payload = @{ name = $Name } | ConvertTo-Json
            $created = Invoke-RestMethod -Uri "$snipeItUrl/api/v1/locations" -Method Post -Headers $script:snipeHeaders -Body $payload
            Start-Sleep -Milliseconds 300
            $snipeIdCache[$Endpoint][$Name] = $created.id
            return $created.id
        } catch {
            Write-Warning "Failed to create location '$Name': $($_.Exception.Message)"
        }
    }
    
    Write-Warning "Could not find $Endpoint '$Name'."
    return $null
}

function Get-OrCreateModel {
    param([string]$ModelName, [string]$AssetType)
    if ([string]::IsNullOrWhiteSpace($ModelName)) { return $null }
    if ($snipeIdCache.models.ContainsKey($ModelName)) { return $snipeIdCache.models[$ModelName] }

    # Try to find existing model first
    $modelId = Get-SnipeItId -Endpoint 'models' -Name $ModelName
    if ($modelId) { return $modelId }

    # Get category ID for the asset type
    $categoryId = Get-SnipeItId -Endpoint 'categories' -Name $AssetType
    if (-not $categoryId) {
        Write-Warning "Cannot create model '$ModelName' - category '$AssetType' not found"
        return $null
    }

    # Create the model
    try {
        Write-Host "Creating model '$ModelName' under category '$AssetType'..."
        $payload = @{ name = $ModelName; category_id = $categoryId } | ConvertTo-Json
        $created = Invoke-RestMethod -Uri "$snipeItUrl/api/v1/models" -Method Post -Headers $script:snipeHeaders -Body $payload
        Start-Sleep -Milliseconds 300
        if ($created.id) {
            Write-Host "Successfully created model '$ModelName' with ID: $($created.id)"
            $snipeIdCache.models[$ModelName] = $created.id
            return $created.id
        }
    } catch {
        Write-Warning "Failed to create model '$ModelName': $($_.Exception.Message)"
    }
    
    return $null
}
#endregion

# --- Main Script Execution ---
try {
    # Test API connectivity first
    Write-Host "Testing Snipe-IT API connectivity..."
    try {
        $testResponse = Invoke-RestMethod -Uri "$snipeItUrl/api/v1/hardware?limit=1" -Method Get -Headers $script:snipeHeaders
        Write-Host "API connectivity successful. Found $($testResponse.total) total assets."
    } catch {
        Write-Error "API connectivity failed: $($_.Exception.Message)"
        Write-Host "API URL: $snipeItUrl"
        Write-Host "Headers: $($script:snipeHeaders | ConvertTo-Json)"
        throw "Cannot connect to Snipe-IT API"
    }

    # Ensure folders exist
    if (-not (Test-Path $dropFolder)) { New-Item -ItemType Directory -Path $dropFolder | Out-Null }
    if (-not (Test-Path $archiveFolder)) { New-Item -ItemType Directory -Path $archiveFolder | Out-Null }

    # Find the newest CSV file in the drop folder
    $latestCsv = Get-ChildItem -Path $dropFolder -Filter "*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $latestCsv) {
        Write-Host "No new CSV files found in '$dropFolder'. Exiting."
        exit
    }

    Write-Host "Processing file: $($latestCsv.FullName)"
    
    # --- STEP 1: Process CSV and Sync to Snipe-IT ---
    $assetsToSync = Import-Csv -Path $latestCsv.FullName
    
    $Global:createdCount = 0; $Global:updatedCount = 0; $Global:failedCount = 0; $Global:skippedCount = 0
    Write-Host "Starting sync process for $($assetsToSync.Count) assets..."
    Write-Host ("-"*50)

    foreach ($asset in $assetsToSync) {
        $uniqueIdValue = $asset.$uniqueIdFieldFreshservice
        if ([string]::IsNullOrWhiteSpace($uniqueIdValue)) {
            Write-Warning "Skipping row because the unique identifier '$uniqueIdFieldFreshservice' is empty."
            $Global:skippedCount++; continue
        }
        Write-Host "Processing Asset: $uniqueIdValue"

        $existingAsset = $null
        $searchApiEndpoint = switch ($uniqueIdFieldSnipe) {
            'asset_tag' { "hardware/bytag/$([System.Web.HttpUtility]::UrlEncode($uniqueIdValue))" }
            'serial'    { "hardware/byserial/$([System.Web.HttpUtility]::UrlEncode($uniqueIdValue))" }
            default     { throw "Invalid uniqueIdFieldSnipe in config.json: '$uniqueIdFieldSnipe'" }
        }
        try {
            $existingAsset = Invoke-RestMethod -Uri "$snipeItUrl/api/v1/$searchApiEndpoint" -Method Get -Headers $script:snipeHeaders
        } catch {
            if ($_.Exception.Response.StatusCode -ne "NotFound") {
                Write-Error "Error searching for asset ${$uniqueIdValue}: $($_.Exception.Message)"
                $Global:failedCount++; continue
            }
        }

        $payload = @{}
        foreach ($csvHeader in $fieldMapping.PSObject.Properties.Name) {
            $snipeField = $fieldMapping.$csvHeader
            $csvValue = $asset.$csvHeader
            if ([string]::IsNullOrWhiteSpace($csvValue)) { continue }

            switch ($snipeField) {
                "model" { 
                    $modelId = Get-OrCreateModel -ModelName $csvValue -AssetType $asset."Asset Type"
                    $payload["model_id"] = $modelId
                }
                "status_label"{ $payload["status_id"] = Get-SnipeItId -Endpoint 'statuslabels' -Name $csvValue }
                "location"    { $payload["location_id"] = Get-SnipeItId -Endpoint 'locations' -Name $csvValue }
                "company"     { $payload["company_id"] = Get-SnipeItId -Endpoint 'companies' -Name $csvValue }
                "department"  { $payload["department_id"] = Get-SnipeItId -Endpoint 'departments' -Name $csvValue }
                "assigned_to" {
                    $userId = Get-SnipeItId -Endpoint 'users' -Name $csvValue
                    if ($userId) { $payload["assigned_to"] = $userId; $payload["assigned_type"] = 'App\Models\User' }
                }
                default { $payload[$snipeField] = $csvValue }
            }
        }

        if ($existingAsset -and $existingAsset.id) {
            $assetId = $existingAsset.id
            try {
                Invoke-RestMethod -Uri "$snipeItUrl/api/v1/hardware/$assetId" -Method Patch -Headers $script:snipeHeaders -Body ($payload | ConvertTo-Json -Depth 5) | Out-Null
                Write-Host "  -> SUCCESS: Updated existing asset ID $assetId." -ForegroundColor Cyan
                $Global:updatedCount++
                Start-Sleep -Milliseconds 300
            } catch { Write-Warning "  -> FAILED to update asset $uniqueIdValue (ID $assetId): $($_.Exception.Message)"; $Global:failedCount++ }
        } else {
            if (-not $payload.ContainsKey("serial")) { $payload["serial"] = $uniqueIdValue }
            if (-not $payload.ContainsKey("asset_tag") -or [string]::IsNullOrWhiteSpace($payload["asset_tag"])) { 
                $payload["asset_tag"] = $uniqueIdValue 
            }
            if (-not $payload.ContainsKey("model_id") -or $payload["model_id"] -eq $null) { 
                Write-Warning "  -> SKIPPED: Asset $uniqueIdValue is missing a valid model_id."
                $Global:skippedCount++; continue 
            }
            if (-not $payload.ContainsKey("status_id") -or $payload["status_id"] -eq $null) {
                $defaultStatusId = Get-SnipeItId -Endpoint 'statuslabels' -Name 'Stock'
                if ($defaultStatusId) { 
                    $payload["status_id"] = $defaultStatusId 
                } else {
                    Write-Warning "  -> SKIPPED: Asset $uniqueIdValue is missing a status."; $Global:skippedCount++; continue
                }
            }
            try {
                $response = Invoke-RestMethod -Uri "$snipeItUrl/api/v1/hardware" -Method Post -Headers $script:snipeHeaders -Body ($payload | ConvertTo-Json -Depth 5)
                if ($response.status -eq 'success' -or $response.id) {
                    Write-Host "  -> SUCCESS: Created new asset with ID $($response.id)." -ForegroundColor Green
                    $Global:createdCount++
                } else {
                    Write-Warning "  -> Asset creation failed: $($response | ConvertTo-Json)"
                    $Global:failedCount++
                }
                Start-Sleep -Milliseconds 300
            } catch { 
                Write-Error "  -> FAILED to create asset $($uniqueIdValue): $($_.Exception.Message)"
                $Global:failedCount++ 
            }
        }
    }

    # --- STEP 2: Archive the processed file ---
    $archiveFileName = "$($latestCsv.BaseName)_$(Get-Date -Format 'yyyyMMddHHmmss')$($latestCsv.Extension)"
    $archiveFilePath = Join-Path $archiveFolder $archiveFileName
    Move-Item -Path $latestCsv.FullName -Destination $archiveFilePath
    Write-Host "Successfully processed and archived file to '$archiveFilePath'" -ForegroundColor Green
    
} catch {
    Write-Error "A critical error occurred: $($_.Exception.Message)"
} finally {
    # --- Final Summary ---
    if ($latestCsv) {
        Write-Host ("="*50); Write-Host "SYNC COMPLETE for file $($latestCsv.Name)"; Write-Host ("-"*50)
        Write-Host "Assets Created: $Global:createdCount" -ForegroundColor Green
        Write-Host "Assets Updated: $Global:updatedCount" -ForegroundColor Cyan
        Write-Host "Assets Skipped: $Global:skippedCount" -ForegroundColor Yellow
        Write-Host "Sync Failures: $Global:failedCount" -ForegroundColor Red
        Write-Host ("="*50)
    }
}