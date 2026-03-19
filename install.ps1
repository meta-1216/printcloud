# PrintCloud インストールスクリプト
# 使い方: powershell -ExecutionPolicy Bypass -Command "& {$url='...';$cid='顧客ID';$br='拠点名';iex (irm $url)}"

$InstallDir = "C:\PrintCloud"
$AgentUrl   = "https://raw.githubusercontent.com/meta-1216/printcloud/main/agent.ps1"
$TaskName   = "PrintCloudAgent"

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host " PrintCloud エージェント インストール" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# フォルダ作成
if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir | Out-Null
    Write-Host "フォルダを作成しました: $InstallDir" -ForegroundColor Green
}

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
$configContent = @"
# PrintCloud エージェント設定ファイル
CLIENT_ID=$cid
BRANCH=$br
"@
$configContent | Set-Content "$InstallDir\config.txt" -Encoding UTF8
Write-Host "設定ファイルを作成しました" -ForegroundColor Green

# タスクスケジューラに登録（PC起動時に自動実行）
$action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstallDir\agent.ps1`""
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit 0 -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

# 既存タスクを削除
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal | Out-Null
Write-Host "タスクスケジューラに登録しました" -ForegroundColor Green

# 今すぐ起動
Start-ScheduledTask -TaskName $TaskName
Write-Host "エージェントを起動しました" -ForegroundColor Green

Write-Host ""
Write-Host "=================================" -ForegroundColor Cyan
Write-Host " インストール完了！" -ForegroundColor Cyan
Write-Host " PC再起動後も自動で起動します" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 3
