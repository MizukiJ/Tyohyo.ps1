# ================================================================
# 　　　　　　　　　　　　　　　　ログ出力関数
# ================================================================
function outputLogg {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [string]$Level = "INFO", # デフォルトはINFO
        [string]$Source = "Powershell" # ログ出力元（例: DB_CONNECTION, SP_EXECUTIONなど）
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" # ミリ秒まで含む
    
    # ログファイルに出力する形式
    # [日時][レベル][ソース] メッセージ
    $logEntry = "[$timestamp][$(($Level.ToUpper()).PadRight(5))][$(($Source.ToUpper()).PadRight(15))] $Message"
    Add-Content -Path $logFile -Value $logEntry

    # コンソール出力はログレベルによって色を変える
    $consoleOutput = "[$timestamp][$(($Level.ToUpper()).PadRight(5))][$(($Source.ToUpper()).PadRight(15))] $Message"
    switch ($Level.ToUpper()) {
        "INFO"  { Write-Host $consoleOutput -ForegroundColor Green }    # 情報メッセージは緑
        "WARN"  { Write-Host $consoleOutput -ForegroundColor Yellow }   # 警告メッセージは黄
        "ERROR" { Write-Host $consoleOutput -ForegroundColor Red }      # エラーメッセージは赤
        "FATAL" { Write-Host $consoleOutput -ForegroundColor White -BackgroundColor Red } # 致命的エラーは赤背景に白文字
        default { Write-Host $consoleOutput } # その他のレベルはデフォルト色
    }
}