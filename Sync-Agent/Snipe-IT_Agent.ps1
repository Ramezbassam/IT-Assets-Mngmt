#Requires -Version 5.0
<#
.SYNOPSIS
    Syncs local device information to Snipe-IT via the API.
.DESCRIPTION
    This script gathers information from the local host machine and creates or updates
    the asset in Snipe-IT. It uses the Serial Number as the primary key for matching.
    It dynamically creates new Models if they don't exist and matches existing 
    Categories and Status Labels.
.NOTES
    File Name: Snipe-IT_Agent.ps1
    Author:    Ramez Bassam
    Version:   2.0 (Local device sync)
    Requires:  PowerShell 5.0+, config.json
#>

# --- Script Configuration ---
$ErrorActionPreference = 'Stop'
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ConfigFile = Join-Path $ScriptPath "config.json"

# --- Counters ---
$Counters = @{
    Success = 0
    Skipped = 0
    Failed  = 0
}

# --- Asset Tag Counter for Stock Items ---
$NextAssetTagNumber = 1

# --- Helper Functions ---
function Get-LocalDeviceInfo {
    [CmdletBinding()]
    param()
    
    try {
        $ComputerInfo = Get-ComputerInfo
        
        $DeviceInfo = [PSCustomObject]@{
            Name = $env:COMPUTERNAME
            SerialNumber = $ComputerInfo.BiosSeralNumber
            Model = "$($ComputerInfo.CsManufacturer) $($ComputerInfo.CsModel)"
            AssetTag = $null  # Will be generated if needed
            LastLoginBy = $env:USERNAME
            Status = "Ready to Deploy"
            Category = "Laptop"
        }
        
        return $DeviceInfo
    }
    catch {
        Write-Error "Failed to gather local device information: $($_.Exception.Message)"
        return $null
    }
}

function Get-AllSnipeitPages {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$BaseUrl,
        [Parameter(Mandatory = $true)]
        [hashtable]$Headers
    )

    $AllRows = [System.Collections.Generic.List[PSObject]]@()
    $Offset = 0
    $Limit = 500

    do {
        $CurrentUrl = "$BaseUrl" + "?limit=$Limit" + "&offset=$Offset"
        try {
            $Response = Invoke-RestMethod -Uri $CurrentUrl -Method Get -Headers $Headers
            if ($Response.rows) {
                foreach ($Row in $Response.rows) {
                    $AllRows.Add($Row)
                }
                $Offset += $Response.rows.Count
            }
        }
        catch {
            Write-Error "Failed to fetch data from $CurrentUrl. Error: $($_.Exception.Message)"
            throw "API pagination failed. Stopping script."
        }
    } while ($Offset -lt $Response.total)

    return $AllRows
}

# --- 1. Load Config and Prepare Environment ---
Write-Host "Initializing..."
if (-not (Test-Path $ConfigFile)) {
    Write-Error "Configuration file not found at $ConfigFile. Please create it."
    return
}
$Config = Get-Content $ConfigFile -Raw | ConvertFrom-Json

# Prepare API Headers
$Headers = @{
    "Authorization" = "Bearer $($Config.ApiKey)"
    "Accept"        = "application/json"
    "Content-Type"  = "application/json"
}

# --- 2. Fetch Models, Categories, Statuses, and Users ---
$ModelCache = @{}
$CategoryCache = @{}
$StatusCache = @{}
$UserCache = @{}
$AssetTagCache = @{}  # Cache for existing asset tags

try {
    Write-Host "Fetching all models from Snipe-IT..."
    $AllModels = Get-AllSnipeitPages -BaseUrl "$($Config.ApiUrl)/api/v1/models" -Headers $Headers
    foreach ($Model in $AllModels) {
        if (-not $ModelCache.ContainsKey($Model.name)) {
            $ModelCache.Add($Model.name, $Model.id)
        }
    }
    Write-Host "Loaded $($ModelCache.Count) models."

    Write-Host "Fetching all categories from Snipe-IT..."
    $AllCategories = Get-AllSnipeitPages -BaseUrl "$($Config.ApiUrl)/api/v1/categories" -Headers $Headers
    foreach ($Category in $AllCategories) {
        if (-not $CategoryCache.ContainsKey($Category.name)) {
            $CategoryCache.Add($Category.name, $Category.id)
        }
    }
    Write-Host "Loaded $($CategoryCache.Count) categories."

    Write-Host "Fetching all status labels from Snipe-IT..."
    $AllStatuses = Get-AllSnipeitPages -BaseUrl "$($Config.ApiUrl)/api/v1/statuslabels" -Headers $Headers
    foreach ($Status in $AllStatuses) {
        if (-not $StatusCache.ContainsKey($Status.name)) {
            $StatusCache.Add($Status.name, $Status.id)
        }
    }
    Write-Host "Loaded $($StatusCache.Count) status labels."
    
    Write-Host "Fetching all users from Snipe-IT..."
    $AllUsers = Get-AllSnipeitPages -BaseUrl "$($Config.ApiUrl)/api/v1/users" -Headers $Headers
    foreach ($User in $AllUsers) {
        if (-not [string]::IsNullOrWhiteSpace($User.username) -and -not $UserCache.ContainsKey($User.username)) {
            $UserCache.Add($User.username, $User.id)
        }
    }
    Write-Host "Loaded $($UserCache.Count) users."
    
    # Fetch all existing asset tags
    Write-Host "Fetching all asset tags from Snipe-IT..."
    $AllAssets = Get-AllSnipeitPages -BaseUrl "$($Config.ApiUrl)/api/v1/hardware" -Headers $Headers
    foreach ($Asset in $AllAssets) {
        if (-not [string]::IsNullOrWhiteSpace($Asset.asset_tag)) {
            $AssetTagCache[$Asset.asset_tag] = $true
        }
    }
    Write-Host "Loaded $($AssetTagCache.Count) existing asset tags."
    
    # Find highest existing STOCK-* tag for Stock items
    if ($StatusCache.ContainsKey("Stock")) {
        $StockStatusId = $StatusCache["Stock"]
        $StockAssetsUrl = "$($Config.ApiUrl)/api/v1/hardware?status_id=$StockStatusId"
        $StockAssets = Get-AllSnipeitPages -BaseUrl $StockAssetsUrl -Headers $Headers
        $HighestNumber = 0
        foreach ($Asset in $StockAssets) {
            if ($Asset.asset_tag -match '^STOCK-(\d+)$') {
                $Number = [int]$matches[1]
                if ($Number -gt $HighestNumber) { $HighestNumber = $Number }
            }
        }
        $NextAssetTagNumber = $HighestNumber + 1
        Write-Host "Next stock tag will be: STOCK-$NextAssetTagNumber"
    }
}
catch {
    Write-Error "Failed to load initial data from Snipe-IT API: $($_.Exception.Message)"
    Write-Error "Please check APIUrl and ApiKey in config.json."
    return
}

# --- 3. Process Local Device Information ---
Write-Host "Gathering local device information..."
$LocalDevice = Get-LocalDeviceInfo
if (-not $LocalDevice) {
    Write-Error "Failed to gather local device information. Exiting."
    return
}

Write-Host "Processing device: $($LocalDevice.Name)..."
try {
    $Serial = $LocalDevice.SerialNumber
    $ModelName = $LocalDevice.Model
    $CategoryName = $LocalDevice.Category
    $AssetTag = $LocalDevice.AssetTag
    $StatusName = $LocalDevice.Status
    $AssignedUser = $LocalDevice.LastLoginBy

    # --- Validation ---
    if ([string]::IsNullOrWhiteSpace($Serial)) {
        Write-Warning "Serial number is missing."
        $Counters.Skipped++
        return
    }
    if ([string]::IsNullOrWhiteSpace($ModelName)) {
        Write-Warning "Model is missing."
        $Counters.Skipped++
        return
    }
    if ([string]::IsNullOrWhiteSpace($CategoryName)) {
        Write-Warning "Category is missing."
        $Counters.Skipped++
        return
    }

    # Asset tag will be generated only for new assets

    # --- Find Status ID ---
    if (-not $StatusCache.ContainsKey($StatusName)) {
        Write-Warning "Status '$StatusName' not found in Snipe-IT. Please create the Status Label."
        $Counters.Skipped++
        return
    }
    $StatusId = $StatusCache[$StatusName]

    # --- Find Category ID ---
    if (-not $CategoryCache.ContainsKey($CategoryName)) {
        Write-Warning "Category '$CategoryName' not found in Snipe-IT. Please create it manually."
        $Counters.Skipped++
        return
    }
    $CategoryId = $CategoryCache[$CategoryName]

    # --- Find or Create Model ID ---
    $ModelId = $null
    if ($ModelCache.ContainsKey($ModelName)) {
        $ModelId = $ModelCache[$ModelName]
    }
    else {
        Write-Host "Model '$ModelName' not found. Creating new model..."
        $NewModelPayload = @{
            name        = $ModelName
            category_id = $CategoryId
        }
        
        if ($Config.NewModelDefaults.manufacturer_id) {
            $NewModelPayload.Add("manufacturer_id", $Config.NewModelDefaults.manufacturer_id)
        }
        if ($Config.NewModelDefaults.fieldset_id) {
            $NewModelPayload.Add("fieldset_id", $Config.NewModelDefaults.fieldset_id)
        }

        $NewModelBody = $NewModelPayload | ConvertTo-Json
        
        try {
            $CreateModelUrl = "$($Config.ApiUrl)/api/v1/models"
            $NewModelResponse = Invoke-RestMethod -Uri $CreateModelUrl -Method Post -Headers $Headers -Body $NewModelBody
            
            if ($NewModelResponse.status -eq 'success') {
                $ModelId = $NewModelResponse.payload.id
                $ModelCache.Add($ModelName, $ModelId)
                Write-Host "Successfully created model '$ModelName' with ID $ModelId."
            } else {
                Write-Warning "Failed to create model '$ModelName': $($NewModelResponse.messages | ConvertTo-Json -Depth 3)"
                $Counters.Failed++
                return
            }
        } catch {
            Write-Error "API error creating model '$ModelName': $($_.Exception.Message)"
            $Counters.Failed++
            return
        }
    }

    # Check for existing asset by serial
    $ExistingAsset = $null
    $CheckUrl = "$($Config.ApiUrl)/api/v1/hardware/byserial/$([System.Web.HttpUtility]::UrlEncode($Serial))"
    
    try {
        $ExistingAssetResponse = Invoke-RestMethod -Uri $CheckUrl -Method Get -Headers $Headers
        if ($ExistingAssetResponse.total -gt 0) {
            $ExistingAsset = $ExistingAssetResponse.rows[0]
        }
    }
    catch {
        if ($_.Exception.Response.StatusCode -ne [System.Net.HttpStatusCode]::NotFound) {
            Write-Error "Error checking for existing asset '$Serial': $($_.Exception.Message)"
            $Counters.Failed++
            return
        }
    }

    # --- Create or Update Asset ---
    if ($ExistingAsset) {
        # UPDATE existing asset
        $Payload = @{
            name      = $LocalDevice.Name
            serial    = $Serial
            status_id = $StatusId
            model_id  = $ModelId
        }
        
        # Check if stock asset is being assigned to user (going live)
        $NeedsNewAssetTag = $false
        if (-not [string]::IsNullOrWhiteSpace($AssignedUser) -and $ExistingAsset.asset_tag -match '^STOCK-\d+$') {
            $NeedsNewAssetTag = $true
        }
        
        # Generate new ASSET tag if stock is going live
        if ($NeedsNewAssetTag) {
            do {
                $RandomNumber = Get-Random -Minimum 10000 -Maximum 99999
                $NewAssetTag = "ASSET-$RandomNumber"
            } while ($AssetTagCache.ContainsKey($NewAssetTag))
            $AssetTagCache[$NewAssetTag] = $true
            $Payload.Add("asset_tag", $NewAssetTag)
            Write-Host "Converting stock asset to live: $($ExistingAsset.asset_tag) -> $NewAssetTag"
        }
        
        # Handle user assignment
        if (-not [string]::IsNullOrWhiteSpace($AssignedUser)) {
            if ($UserCache.ContainsKey($AssignedUser)) {
                $UserId = $UserCache[$AssignedUser]
                $Payload.Add("assigned_to", $UserId)
                $Payload.Add("assigned_type", "App\Models\User")
            } else {
                Write-Warning "User '$AssignedUser' not found in Snipe-IT. Asset will not be assigned."
            }
        }
        
        $PayloadJson = $Payload | ConvertTo-Json
        $UpdateUrl = "$($Config.ApiUrl)/api/v1/hardware/$($ExistingAsset.id)"
        try {
            $UpdateResponse = Invoke-RestMethod -Uri $UpdateUrl -Method Patch -Headers $Headers -Body $PayloadJson
            if ($UpdateResponse.status -eq 'success') {
                Write-Host "Updated asset with serial '$Serial' (ID: $($ExistingAsset.id))"
                $Counters.Success++
            } else {
                Write-Warning "Failed to update asset '$Serial': $($UpdateResponse.messages | ConvertTo-Json -Depth 3)"
                $Counters.Failed++
            }
        } catch {
            Write-Error "API error updating asset '$Serial': $($_.Exception.Message)"
            $Counters.Failed++
        }
    }
    else {
        # CREATE new asset (generate asset tag for new assets only)
        if ($StatusName -eq "Stock") {
            $AssetTag = "STOCK-$NextAssetTagNumber"
            $NextAssetTagNumber++
        } else {
            # Generate unique random asset tag for new assets
            do {
                $RandomNumber = Get-Random -Minimum 10000 -Maximum 99999
                $AssetTag = "ASSET-$RandomNumber"
            } while ($AssetTagCache.ContainsKey($AssetTag))
            $AssetTagCache[$AssetTag] = $true  # Add to cache to prevent duplicates
        }
        Write-Host "Generated asset tag '$AssetTag' for new device with serial '$Serial'"
        
        $Payload = @{
            name      = $LocalDevice.Name
            serial    = $Serial
            asset_tag = $AssetTag
            status_id = $StatusId
            model_id  = $ModelId
        }
        
        # Handle user assignment
        if (-not [string]::IsNullOrWhiteSpace($AssignedUser)) {
            if ($UserCache.ContainsKey($AssignedUser)) {
                $UserId = $UserCache[$AssignedUser]
                $Payload.Add("assigned_to", $UserId)
                $Payload.Add("assigned_type", "App\Models\User")
            } else {
                Write-Warning "User '$AssignedUser' not found in Snipe-IT. Asset will not be assigned."
            }
        }
        
        $PayloadJson = $Payload | ConvertTo-Json
        $CreateUrl = "$($Config.ApiUrl)/api/v1/hardware"
        try {
            $CreateResponse = Invoke-RestMethod -Uri $CreateUrl -Method Post -Headers $Headers -Body $PayloadJson
            if ($CreateResponse.status -eq 'success') {
                Write-Host "Created new asset with serial '$Serial' (ID: $($CreateResponse.payload.id))"
                $Counters.Success++
            } else {
                Write-Warning "Failed to create asset '$Serial': $($CreateResponse.messages | ConvertTo-Json -Depth 3)"
                $Counters.Failed++
            }
        } catch {
            Write-Error "API error creating asset '$Serial': $($_.Exception.Message)"
            $Counters.Failed++
        }
    }
}
catch {
    Write-Error "Unhandled error processing device '$($LocalDevice.Name)': $($_.Exception.Message)"
    $Counters.Failed++
}

# --- Display Counters ---
Write-Host "----------------------"
Write-Host "--- Sync Summary ---"
Write-Host "Success: $($Counters.Success)"
Write-Host "Skipped: $($Counters.Skipped)"
Write-Host "Failed:  $($Counters.Failed)"
Write-Host "----------------------"
Write-Host "Script finished."