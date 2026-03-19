# PrintCloud Agent - PowerShell

$ConfigPath = "$PSScriptRoot\config.txt"
$LogPath    = "$PSScriptRoot\agent_log.txt"
$FirebaseUrl = "https://firestore.googleapis.com/v1/projects/printcloud-22933/databases/(default)/documents/print_logs"

function Load-Config {
    $config = @{ CLIENT_ID = "unset"; BRANCH = "unset"; PRINTER_IP = ""; PRINTER_PASS = "" }
    if (Test-Path $ConfigPath) {
        Get-Content $ConfigPath -Encoding UTF8 | ForEach-Object {
            if ($_ -match "^([^#=]+)=(.+)$") {
                $config[$Matches[1].Trim()] = $Matches[2].Trim()
            }
        }
    }
    return $config
}

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts $msg"
    Write-Host $line
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
}

function Send-ToFirebase($config, $user, $printer, $pages) {
    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ts  = [int][double]::Parse((Get-Date -UFormat %s))
    $body = @{
        fields = @{
            date        = @{ stringValue  = $now }
            username    = @{ stringValue  = $user }
            printer     = @{ stringValue  = $printer }
            docname     = @{ stringValue  = "unknown" }
            pages       = @{ integerValue = "$pages" }
            color       = @{ stringValue  = "unknown" }
            client_id   = @{ stringValue  = $config.CLIENT_ID }
            branch      = @{ stringValue  = $config.BRANCH }
            os          = @{ stringValue  = "Windows" }
            timestamp   = @{ integerValue = "$ts" }
        }
    } | ConvertTo-Json -Depth 5

    try {
        Invoke-RestMethod -Uri $FirebaseUrl -Method POST -Body $body -ContentType "application/json" | Out-Null
        Write-Log "[OK] $now | $($config.CLIENT_ID) | $($config.BRANCH) | $user | $printer | ${pages}p -> Firebase"
    } catch {
        Write-Log "[ERR] Firebase: $_"
    }
}

function Parse-PrintMessage($message) {
    $user    = "unknown"
    $printer = "unknown"
    $pages   = 1

    # Japanese: 所有者 \\菊地良 の cmind) は Canon GX7130 series (ポート
    if ($message -match "所有者 \\\\(.+?) の") {
        $user = $Matches[1]
    }
    if ($message -match "\) は (.+?) \(ポート") {
        $printer = $Matches[1]
    }

    # English: owned by cmind on \\菊地良 was printed on Canon GX7130
    if ($user -eq "unknown" -and $message -match "owned by \S+ on \\\\(.+?) was printed on (.+?) through") {
        $user    = $Matches[1]
        $printer = $Matches[2]
    }

    if ($message -match "印刷したページ数: (\d+)") {
        $pages = [int]$Matches[1]
    } elseif ($message -match "Pages printed: (\d+)") {
        $pages = [int]$Matches[1]
    }

    return @{ user = $user; printer = $printer; pages = $pages }
}


# Main
$config = Load-Config
Write-Log "PrintCloud Agent start CLIENT_ID=$($config.CLIENT_ID) BRANCH=$($config.BRANCH)"

$seen = @{}
$events = Get-WinEvent -LogName "Microsoft-Windows-PrintService/Operational" -MaxEvents 100 -ErrorAction SilentlyContinue |
    Where-Object { $_.Id -eq 307 }
foreach ($e in $events) {
    $key = $e.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss") + $e.Message.Substring(0, [Math]::Min(40, $e.Message.Length))
    $seen[$key] = $true
}
Write-Log "Skip $($seen.Count) existing logs"

while ($true) {
    try {
        $events = Get-WinEvent -LogName "Microsoft-Windows-PrintService/Operational" -MaxEvents 20 -ErrorAction SilentlyContinue |
            Where-Object { $_.Id -eq 307 }

        foreach ($e in $events) {
            $key = $e.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss") + $e.Message.Substring(0, [Math]::Min(40, $e.Message.Length))
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $parsed = Parse-PrintMessage $e.Message
            Send-ToFirebase $config $parsed.user $parsed.printer $parsed.pages
        }
    } catch {
        Write-Log "[ERR] $_"
    }
    Start-Sleep -Seconds 5
}
