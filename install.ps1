# PrintCloud Install Script
$InstallDir  = "$env:APPDATA\PrintCloud"
$AgentUrl    = "https://raw.githubusercontent.com/meta-1216/printcloud/main/agent.ps1"
$StartupDir  = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
$ShortcutPath = "$StartupDir\PrintCloud.lnk"

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host " PrintCloud Agent Install" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir | Out-Null
}
Write-Host "Install dir: $InstallDir" -ForegroundColor Green

Write-Host "Downloading agent..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $AgentUrl -OutFile "$InstallDir\agent.ps1" -UseBasicParsing
    Write-Host "Download OK" -ForegroundColor Green
} catch {
    Write-Host "Download failed: $_" -ForegroundColor Red
    exit 1
}

$configContent = "# PrintCloud config`r`nCLIENT_ID=$cid`r`nBRANCH=$br"
[System.IO.File]::WriteAllText("$InstallDir\config.txt", $configContent, [System.Text.Encoding]::UTF8)
Write-Host "Config created" -ForegroundColor Green

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallDir\agent.ps1`""
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.Save()
Write-Host "Startup registered" -ForegroundColor Green

Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallDir\agent.ps1`"" -WindowStyle Hidden
Write-Host "Agent started" -ForegroundColor Green

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host " Install complete!" -ForegroundColor Cyan
Write-Host " Auto-starts on next reboot" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 2
