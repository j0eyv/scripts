$targetFile = Join-Path $env:LocalAppData "Publishers\8wekyb3d8bbwe\TeamsSharedConfig\app_switcher_settings.json"

if (Test-Path -Path $targetFile -PathType Leaf) {
    Write-Host "File exists"
    exit 0
}
else {
    Write-Host "File missing"
    exit 1
}