<#
.SYNOPSIS
    Sends the sample payload to the InsertContent function (local or deployed).

.EXAMPLE
    # Local (default URL http://localhost:7071)
    ./Invoke-InsertContent.ps1

.EXAMPLE
    # Deployed
    ./Invoke-InsertContent.ps1 -BaseUrl "https://your-func-app.azurewebsites.net" -FunctionKey "abc123..."
#>
[CmdletBinding()]
param(
    [string]$BaseUrl = "http://localhost:7071",
    [string]$FunctionKey,
    [string]$PayloadPath = (Join-Path $PSScriptRoot "sample-payload.json")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $PayloadPath)) {
    throw "Payload file not found: $PayloadPath"
}

$uri = "$BaseUrl/api/InsertContent"
$headers = @{ "Content-Type" = "application/json" }
if ($FunctionKey) { $headers["x-functions-key"] = $FunctionKey }

$body = Get-Content -Path $PayloadPath -Raw

Write-Host "POST $uri" -ForegroundColor Cyan
try {
    $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body
    Write-Host "Success:" -ForegroundColor Green
    $response | ConvertTo-Json -Depth 8
}
catch {
    Write-Host "Request failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.ErrorDetails) { Write-Host $_.ErrorDetails.Message }
    throw
}
