# 変更履歴
# Ver1.00 20250724

# ================================================================
#             グローバルエラーハンドリング (trapブロック)
# ================================================================
trap {
    Write-Host "`n"
    Write-Host "FATAL ERROR: スクリプトで致命的なエラーが発生しました！" -ForegroundColor Red -BackgroundColor Black
    Write-Host "---------------------------------------------------------" -ForegroundColor Red
    Write-Host "エラーメッセージ: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "エラータイプ: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-Host "発生場所: $($_.InvocationInfo.ScriptName) (行 $($_.InvocationInfo.ScriptLineNumber))" -ForegroundColor Red
    Write-Host "詳細: $($_.Exception.ToString())" -ForegroundColor DarkRed
    Write-Host "---------------------------------------------------------" -ForegroundColor Red
    
    # outputLogg 関数が定義されていればログにも出力
    if (Get-Command -Name outputLogg -ErrorAction SilentlyContinue) {
        outputLogg "FATAL ERROR: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" -Level "FATAL" -Source "GLOBAL_TRAP"
    } else {
        # outputLogg が定義されていない場合は、直接ファイルに書き込む (logFileが定義されていれば)
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        # $logFile 変数が定義されていることを前提とする
        Add-Content -Path $logFile -Value "[$timestamp][FATAL][GLOBAL_TRAP    ] FATAL ERROR (trap): $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" -ErrorAction SilentlyContinue
    }
    Exit # エラー発生時はスクリプトを終了
}



# ================================================================
#   　　　　　　　　　　　パス/Function読込
# ================================================================

# 現在実行中のスクリプトファイルが存在するディレクトリのパスを取得
# PS1では $PSScriptRoot が使えないため、この方法を使用
$scriptDirectory = (Get-Item $MyInvocation.MyCommand.Path).Directory.FullName
$functionsPath = Join-Path -Path $scriptDirectory -ChildPath "Functions"

# 関数を定義したファイルをドットソーシングで読み込み
# これにより、関数がこのスクリプト内で利用可能になります
. (Join-Path -Path $functionsPath -ChildPath "Get-IniContent.ps1")
. (Join-Path -Path $functionsPath -ChildPath "outputLogg.ps1")



# ================================================================
#   　　　　　　　 iniファイル読み込み/変数定義
# ================================================================

# 同じフォルダにある settings.ini ファイルのフルパスを生成
$iniFilePath = Join-Path -Path $scriptDirectory -ChildPath "settings.ini"
# Get-IniContent 関数を使って INIファイルから設定を読み込む
$mySettings = Get-IniContent -Path $iniFilePath

# ログファイルのパスを設定
$logFile = $mySettings.filePath.logFile

# データベース接続情報
$serverName = $mySettings.settingSQL.serverName
$databaseName = $mySettings.settingSQL.databaseName
$commandTimeout = $mySettings.settingSQL.commandTimeout

# マクロのパスを設定
$WorkbookPath = $mySettings.filePath.WorkbookPath
$MacroName = $null # Excel VBAのマクロ名を指定
$para = $null # Excel VBAのマクロのパラメータ

# 接続文字列
$connectionString = "Server=$serverName;Database=$databaseName;Integrated Security=True;" # Windows認証
# SQL認証の場合の例: "Server=$serverName;Database=$databaseName;User ID=your_user;Password=your_password;"

# ADOオブジェクト変数（一応nullしておく）
$connection = $null
$command = $null
# $reader は結果セットがないため不要です

# Excelアプリケーションのオブジェクトを初期化
$excel = $null # 初期化
$workbook = $null # 初期化



# ================================================================
#                      ログディレクトリ確認
# ================================================================
# ログディレクトリが存在しない場合は作成する
$logDirectory = Split-Path -Path $logFile -Parent
if (-not (Test-Path -Path $logDirectory -PathType Container)) {
    try {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        Write-Host "ログディレクトリ '$logDirectory' を作成しました。"
    } catch {
        Write-Error "ログディレクトリの作成に失敗しました: $($_.Exception.Message)"
        # エラー発生時はExitでスクリプトを終了。Read-Hostは削除。
        Exit # ディレクトリ作成失敗時は終了
    }
}


# ================================================================
#                       メイン処理の開始
# ================================================================
try { # sp_PropertyImportは失敗することが考えられる為、別のTry-Catchブロックとする
    　# ★ただエラーが出たときに、値がnullになってしまうため、もしかしたらこれも問題がない初期CSVを用意する必要ある?
    outputLogg "実行開始" -Level "INFO" -Source "Powershell"

    # 1. SqlConnection オブジェクトの作成とオープン
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()


    # ================================================================
    #                         SQL：プロパティインポート
    # ================================================================
    # ストアドプロシージャ名
    $storedProcedureName_PropImport = "dbo.sp_PropertyImport" # 変数名を明確化

    # 2. SqlCommand オブジェクトの作成 (初回のみ)
    $command = New-Object System.Data.SqlClient.SqlCommand($storedProcedureName_PropImport, $connection)
    $command.CommandType = [System.Data.CommandType]::StoredProcedure # ストアドプロシージャを指定
    $command.CommandTimeout = $commandTimeout # タイムアウト時間を設定

    # 3. SqlParameter オブジェクトの作成と追加 (入力パラメータ) がある場合
    # 例: $paramExample = $command.Parameters.Add("@YourParameterName", [System.Data.SqlDbType]::NVarChar, 100)
    # $paramExample.Value = "YourParameterValue"
    # $paramExample.Direction = [System.Data.ParameterDirection]::Input

    # 4. コマンドの実行 (結果セットがないため ExecuteNonQuery() を使用)
    $command.ExecuteNonQuery() 

    # 5.完了通知
    #outputLogg "ストアドプロシージャ '$($storedProcedureName_PropImport)' の実行が完了しました。" -Level "INFO" -Source "SP_EXECUTION"
}
# SQLエラー時 (System.Data.SqlClient.SqlException)
catch [System.Data.SqlClient.SqlException] {
    $errorMessage = "SQL Error: $($_.Exception.Number) - $($_.Exception.Message)"
    outputLogg $errorMessage -Level "ERROR" -Source "DB_OPERATION"
}
# その他のエラー処理
catch {
    $errorMessage = "予期せぬエラー: $($_.Exception.Message)"
    outputLogg $errorMessage -Level "ERROR" -Source "SCRIPT_GENERAL" # Sourceを 'SCRIPT_GLOBAL' から 'SCRIPT_GENERAL' に変更
}


try {
    # ================================================================
    #                         SQL：日報出力
    # ================================================================
    # 1. 新しいストアドプロシージャ名を設定
    $storedProcedureName_TyohyoOut = "dbo.sp_Tyohyo_Out" # 変数名を明確化
    $command.CommandText = "$storedProcedureName_TyohyoOut"

    # 2. 以前のパラメータをクリアする（非常に重要！）
    $command.Parameters.Clear()

    # 3. 2つ目のストアドプロシージャのパラメータを追加
    $command.Parameters.Add("@NipOrGep", [System.Data.SqlDbType]::NVarChar,255).Value = "Nippo"

    # 4. コマンドの実行 (結果セットがないため ExecuteNonQuery() を使用)
    $command.ExecuteNonQuery() 

    # 5.完了通知
    outputLogg "ストアドプロシージャ '$($storedProcedureName_TyohyoOut)' の実行（日報）が完了しました。" -Level "INFO" -Source "SP_EXECUTION"

    
    # ================================================================
    #                         Excel：日報出力
    # ================================================================
    # Excelアプリケーションの起動（新しいオブジェクトでやるので、現在とオブジェクトとは独立）
    $excel = New-Object -ComObject Excel.Application

    # 警告ダイアログを表示しない（読み取り専用で開きますかを制御）
    $excel.DisplayAlerts = $false 
    # ★Excelを非表示にする場合は $false に設定 (デバッグ中は $true)
    $excel.Visible = $false 

    # ワークブックを開く
    $workbook = $excel.Workbooks.Open($WorkbookPath)

    # パラメータなしでマクロを呼び出す
    $MacroName = "testCall_Nip" # Excel VBA側のマクロ名を指定
    $excel.Run($MacroName)
  


    # ================================================================
    #                     月の初めの場合の月報出力
    # ================================================================
    $currentDay = (Get-Date).Day # 現在の日付の「日」を取得
    if ($currentDay -eq 1) { # 月の1日かどうかを判定

        # ================================================================
        #                         SQL：月報出力
        # ================================================================
        # 1. 新しいストアドプロシージャ名を設定
        #$storedProcedureName_TyohyoOut = "dbo.sp_Tyohyo_Out" # 変数名を明確化
        #$command.CommandText = "$storedProcedureName_TyohyoOut"

        # 2. 以前のパラメータをクリアする（非常に重要！）
        $command.Parameters.Clear() # 日報のパラメータを削除

        # 3. 2つ目のストアドプロシージャのパラメータを追加
        $command.Parameters.Add("@NipOrGep", [System.Data.SqlDbType]::NVarChar,255).Value = "Geppo"

        # 4. コマンドの実行 (結果セットがないため ExecuteNonQuery() を使用)
        $command.ExecuteNonQuery() 

        # 5.完了通知
        outputLogg "ストアドプロシージャ '$($storedProcedureName_TyohyoOut)' の実行（月報）が完了しました。" -Level "INFO" -Source "SP_EXECUTION"


        # ================================================================
        #                         Excel：月報出力
        # ================================================================
        # Excelアプリケーションの起動
        #$excel = New-Object -ComObject Excel.Application
        #$excelWasNewlyCreated = $true # 新しく作成した場合はフラグをtrueにする

        # 警告ダイアログを表示しない（読み取り専用で開きますかを制御） 
        $excel.DisplayAlerts = $false 
        $excel.Visible = $true 

        # ワークブックを開く
        $workbook = $excel.Workbooks.Open($WorkbookPath)

        # パラメータなしでマクロを呼び出す
        $MacroName = "testCall_Gep" # Excel VBA側のマクロ名を指定
        $excel.Run($MacroName)
    } else {
    }
    
}
# SQLエラー時 (System.Data.SqlClient.SqlException)
catch [System.Data.SqlClient.SqlException] {
    $errorMessage = "SQL Error: $($_.Exception.Number) - $($_.Exception.Message)"
    outputLogg $errorMessage -Level "ERROR" -Source "DB_OPERATION"
}
# その他のエラー処理
catch {
    $errorMessage = "予期せぬエラー: $($_.Exception.Message)"
    outputLogg $errorMessage -Level "ERROR" -Source "SCRIPT_GENERAL" # Sourceを 'SCRIPT_GLOBAL' から 'SCRIPT_GENERAL' に変更
}
finally {
    # 6. オブジェクトのクリーンアップと接続のクローズ (非常に重要)
    if ($command -ne $null) {
        $command.Dispose()
    }
    if ($connection -ne $null -and $connection.State -eq [System.Data.ConnectionState]::Open) {
        $connection.Close()
        $connection.Dispose()
    }

    # Excelオブジェクトのクリーンアップ
    if ($workbook -ne $null) {
        # ワークブックに意図しない変更がなければ保存せずに閉じる
        $workbook.Close($false) 
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    if ($excel -ne $null) {
        # 常にこのスクリプトが起動したインスタンスなので、無条件でQuit()する
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    Remove-Variable workbook, excel -ErrorAction SilentlyContinue
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

}