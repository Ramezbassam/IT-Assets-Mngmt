# Simple API test script
$configPath = "C:\Users\ramezab\Downloads\config.json"

if (Test-Path $configPath) {
    $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
    $snipeItUrl = $config.snipeIt.apiUrl.TrimEnd('/')
    $snipeItApiKey = $config.snipeIt.apiKey
} else {
    # Fallback - enter your details here
    $snipeItUrl = "http://10.42.100.10:50101"  # Replace with your URL
    $snipeItApiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiNmY0MDliZDYyMGY4Y2VlN2Y1ZmViNGJjMGM2OTI1NzQyZjlkNThiZTM4N2Y2M2I2ZWY1MTRiZThlMWRmMjIzNTU3ODRkMzJlZjg1ODFiMjciLCJpYXQiOjE3NTk4MjcxMDYuMTYxNjczLCJuYmYiOjE3NTk4MjcxMDYuMTYxNjc4LCJleHAiOjIyMzMyMTI3MDYuMDg0MjE5LCJzdWIiOiIzMDciLCJzY29wZXMiOltdfQ.nsdYwsnBttbcdTEKGx7NVUP9VUlWo45IiBjoQxqUcZdmjfVRPSLL6t4DZsEAHu7A_JNOShaUFUAY9xCBbyw8uQYq_aFgEVPWzyOA26hVyka729qNJ3hsKlC5RUuXcAU7xv8mNSxUeC-7AadOHnx9QpggleF3pKSIGmGbhWY1Mu8ACnzd9rDURKpMa8yNEsPjBGPwV-ViwEYTHFPurP4yRNIx0vHYLO_RVJdB9U8AOSpsf0wt4SAiU2MhiHvzjYAjHTo9h8ySxMiA6kAI_uCjDFkKXXgrY-_H_t2DXQdZ9JL3oeg1bBmdO_nf9_6tXmqHtolSjkOy7oqBfnKKftXP4Nn3lheEztOlozM6FGiWsaWAUIUBdSvXQqbfSsorlAQcHWcVWNT9f4S3qElQdV1KnEjf4A60IxKay_ZlKllMx1vxk5MKJ6jdr_XHA-tNFOe__9hLszxUS-95oUsSasLCHvi_IpDgUMv1kMut0PZKA4ha4GITPnq-w0exUCH3i4oTC8ukMgpN1BkrRv89kNlmt79-Z177dMdw-1vzqzC3-h7Plv1e9XVduH6435sraisp9MwbLm56qUDZYog9YRV5xM2pH8pt76kJcfFFgLzv2tHr3bm_IzLVTfstFWjuy3m7oURud-jNe_uHm4TrAc4zJhoN2svIAbMJPlwcZKeE8AU"  # Replace with your API key
}

$headers = @{ 
    "Authorization" = "Bearer $snipeItApiKey"
    "Accept" = "application/json"
    "Content-Type" = "application/json"
}

Write-Host "Testing Snipe-IT API..."
Write-Host "URL: $snipeItUrl"

# Test different endpoints
$endpoints = @("hardware", "categories", "models", "statuslabels", "locations")

foreach ($endpoint in $endpoints) {
    $cleanUrl = $snipeItUrl.TrimEnd('/').TrimEnd('\\')
    $url = "$cleanUrl/api/v1/$endpoint?limit=1"
    Write-Host "`nTesting: $url"
    
    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers
        Write-Host "SUCCESS: $endpoint - Total: $($response.total)" -ForegroundColor Green
        if ($response.rows -and $response.rows.Count -gt 0) {
            Write-Host "Sample item: $($response.rows[0].name)" -ForegroundColor Cyan
        }
    } catch {
        Write-Host "FAILED: $endpoint - $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception.Response) {
            Write-Host "Status Code: $($_.Exception.Response.StatusCode)" -ForegroundColor Yellow
        }
    }
}