$targetDirectory = Join-Path $env:LocalAppData "Publishers\8wekyb3d8bbwe\TeamsSharedConfig"
$targetFile = Join-Path $targetDirectory "app_switcher_settings.json"

if (-not (Test-Path -Path $targetDirectory)) {
	New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
}

$jsonContent = '{"defaultApp":1,"cohort":"","webAccountId_AAD":"","cohortStage":"","userId_AAD":"","previousT1MachineId":"","previousT1SessionId":""}'

Set-Content -Path $targetFile -Value $jsonContent -Encoding utf8