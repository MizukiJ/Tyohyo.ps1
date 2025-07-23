# ================================================================
# 　　　　　　　　　　　INIファイル読み込み関数
# ================================================================
<#
.SYNOPSIS
指定されたINIファイルの内容を読み込み、ハッシュテーブルとして返します。

.DESCRIPTION
この関数は、INIファイル（.ini）を解析し、セクションとキー-値のペアをPowerShellのハッシュテーブルに変換します。
コメント行（セミコロン ';' またはシャープ '#' で始まる行）と空行はスキップされます。
キー-値のペアは、セクション内またはINIファイルのルートレベルに配置できます。

.PARAMETER Path
読み込むINIファイルのパスを指定します。

.OUTPUTS
System.Collections.Hashtable
INIファイルの内容を表すネストされたハッシュテーブル。
ファイルが見つからない場合や読み込み中にエラーが発生した場合は、$nullを返します。

.EXAMPLE
# 例1: INIファイルを読み込み、内容を表示する
$iniFilePath = "C:\config.ini"
$iniSettings = Get-IniContent -Path $iniFilePath
if ($iniSettings) {
    $iniSettings | Format-List
}

# config.ini の内容例:
# [Database]
# Server = localhost
# Port = 3306
# User = admin
#
# [Application]
# Name = MyApp
# Version = 1.0
# ; This is a comment
# # Another comment
#
# GlobalKey = GlobalValue

.NOTES
コメントは行頭に ';' または '#' を付けることで認識されます。
キーと値の間は '=' で区切られます。
セクションは '[セクション名]' の形式で指定します。
#>
function Get-IniContent {
    param (
        [string]$Path # 読み込むINIファイルのパス
    )

    # 読み込んだ設定を格納するメインのハッシュテーブルを初期化
    # セクション名またはルートキーがキーとなり、その値が別のハッシュテーブル（セクションの場合）
    # または直接の値（ルートキーの場合）となります。
    $settings = @{}
    # 現在処理中のセクション名を保持する変数
    $currentSection = ""

    # ファイルが存在するかどうかを確認
    if (-not (Test-Path $Path)) {
        # ファイルが見つからない場合はエラーメッセージを表示し、$nullを返して終了
        Write-Error "指定されたINIファイルが見つかりません: $Path"
        return $null
    }

    try {
        # INIファイルの内容を1行ずつ読み込む
        $iniContent = Get-Content $Path

        # 読み込んだ各行をループ処理
        foreach ($line in $iniContent) {
            # 行頭・行末の空白を削除して処理を簡略化
            $line = $line.Trim()

            # コメント行 (セミコロン ';' またはシャープ '#' で始まる) や空行はスキップ
            if ($line.StartsWith(";") -or $line.StartsWith("#") -or [string]::IsNullOrWhiteSpace($line)) {
                continue # 次の行へ
            }

            # セクションの検出 (例: [SectionName])
            if ($line.StartsWith("[") -and $line.EndsWith("]")) {
                # セクション名を抽出 (例: "[Message]" から "Message" を取得)
                # Substring(1, $line.Length - 2) は、最初の '[' と最後の ']' を除外します。
                $currentSection = $line.Substring(1, $line.Length - 2)
                # 読み込んだ設定を格納するハッシュテーブルに、新しいセクション用のハッシュテーブルを追加
                # これにより、$settings.SectionName.Key のようにアクセスできるようになります。
                $settings.$currentSection = @{}
            }
            # キーと値のペアの検出 (例: Key = Value)
            elseif ($line.Contains("=")) {
                # 最初のイコール '=' で行を2つの部分に分割（キーと値）
                # Split("=", 2) は、最初の '=' のみで分割し、それ以降の '=' は値の一部として扱います。
                $parts = $line.Split("=", 2)
                # キーの行頭・行末の空白を削除
                $key = $parts[0].Trim()
                # 値の行頭・行末の空白を削除
                $value = $parts[1].Trim()

                # 現在のセクションが設定されている場合、そのセクションにキーと値を追加
                if ($currentSection) {
                    $settings.$currentSection.$key = $value
                }
                # セクションがない場合 (INIファイルの先頭に直接キーがある場合など)
                else {
                    # ルートレベルのキーとして設定に追加
                    $settings.$key = $value
                }
            }
        }
    } catch {
        # INIファイルの読み込み中にエラーが発生した場合、エラーメッセージを表示し、$nullを返す
        Write-Error "INIファイルの読み込み中にエラーが発生しました: $($_.Exception.Message)"
        return $null
    }

    # 読み込んだ設定をハッシュテーブルとして返す
    return $settings
}
