# ================================================================
# 　　　　　　　　　　　iniファイル読み込み関数
# ================================================================
function Get-IniContent {
    param (
        [string]$Path
    )

    $settings = @{}
    $currentSection = ""

    # ファイルが存在するか確認
    if (-not (Test-Path $Path)) {
        Write-Error "指定されたINIファイルが見つかりません: $Path"
        return $null
    }

    try {
        # INIファイルの内容を1行ずつ読み込む
        $iniContent = Get-Content $Path

        foreach ($line in $iniContent) {
            $line = $line.Trim() # 行頭・行末の空白を削除

            # コメント行 (セミコロンまたはシャープで始まる) や空行はスキップ
            if ($line.StartsWith(";") -or $line.StartsWith("#") -or [string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            # セクションの検出 (例: [SectionName])
            if ($line.StartsWith("[") -and $line.EndsWith("]")) {
                # セクション名を抽出 (例: "Message" を取得)
                $currentSection = $line.Substring(1, $line.Length - 2)
                # 設定を格納するハッシュテーブルに新しいセクションのハッシュテーブルを追加
                $settings.$currentSection = @{}
            }
            # キーと値のペアの検出 (例: Key = Value)
            elseif ($line.Contains("=")) {
                # 最初のイコールで分割
                $parts = $line.Split("=", 2)
                $key = $parts[0].Trim() # キーの空白を削除
                $value = $parts[1].Trim() # 値の空白を削除

                # 現在のセクションがある場合、そのセクションにキーと値を追加
                if ($currentSection) {
                    $settings.$currentSection.$key = $value
                }
                # セクションがない場合 (INIファイルの先頭に直接キーがある場合など)
                else {
                    $settings.$key = $value
                }
            }
        }
    } catch {
        # エラー発生時はメッセージを表示し、nullを返す
        Write-Error "INIファイルの読み込み中にエラーが発生しました: $($_.Exception.Message)"
        return $null
    }

    # 読み込んだ設定を返す
    return $settings
}