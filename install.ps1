# PrintCloud インストールスクリプト

$InstallDir  = "$env:APPDATA\PrintCloud"
$AgentUrl    = "https://raw.githubusercontent.com/meta-1216/printcloud/main/agent.ps1"
$StartupDir  = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
$ShortcutPath = "$StartupDir\PrintCloud.lnk"

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host " PrintCloud エージェント インストール" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# フォルダ作成
if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir | Out-Null
}
Write-Host "インストール先: $InstallDir" -ForegroundColor Green

# agent.ps1をダウンロード
Write-Host "エージェントをダウンロード中..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $AgentUrl -OutFile "$InstallDir\agent.ps1" -UseBasicParsing
    Write-Host "ダウンロード完了" -ForegroundColor Green
} catch {
    Write-Host "ダウンロード失敗: $_" -ForegroundColor Red
    exit 1
}

# config.txtを生成
$configContent = "# PrintCloud config`r`nCLIENT_ID=$cid`r`nBRANCH=$br"
[System.IO.File]::WriteAllText("$InstallDir\config.txt", $configContent, [System.Text.Encoding]::UTF8)
Write-Host "設定ファイルを作成しました" -ForegroundColor Green

# スタートアップにショートカットを作成
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallDir\agent.ps1`""
$Shortcut.WorkingDirectory = $InstallDir
$Shortcut.Save()
Write-Host "スタートアップに登録しました" -ForegroundColor Green

# 今すぐ起動
Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallDir\agent.ps1`"" -WindowStyle Hidden
Write-Host "エージェントを起動しました" -ForegroundColor Green

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host " インストール完了！" -ForegroundColor Cyan
Write-Host " PC再起動後も自動で起動します" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 2
