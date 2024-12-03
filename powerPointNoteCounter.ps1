$ErrorActionPreference = 'Stop'

$WORK_PATH = $PSScriptRoot

function Get-Path {
    param (
        [string]$Message,
        [switch]$FILE,
        [switch]$DIR
    )
    if ($null -eq $Message) {
        $Message = "絶対パスまたは PowerShell スクリプトからの相対パスを入力してください"
    }
    $value = Read-Host $Message

    if (($null -eq $value) -or ($value.Length -eq 0)) {
        return Get-Path -Message "未入力です。再度、入力してください"
    }

    if (-not (Test-Path -Path $value)) {
        $tmp = Join-Path -Path $WORK_PATH -ChildPath $value
        if (Test-Path -Path $tmp) {
            $value = $tmp
        }
        else {
            return Get-Path -Message "ファイルまたはフォルダが存在しません。再度、入力してください"
        }
    }
    if ($FILE -and -not (Get-ChildItem -Path $value -File)) {
        return Get-Path -Message "指定されたパスがファイルパスではありません。再度、入力してください"
    }
    if ($DIR -and -not (Get-ChildItem -Path $value -Directory)) {
        return Get-Path -Message "指定されたパスがフォルダパスではありません。再度、入力してください"
    }
    return $value
}

function Out-Log {
    param (
        [string]$Message,
        [switch]$INFO,
        [switch]$ERR
    )
    $file = "error-" + (Get-Date -Format "yyyyMMdd") + ".log"
    if (-Not($file)) {
        New-Item -Path $WORK_PATH -Name $file -ItemType "file"
    }
    if ($ERR) {
        $Message = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + " [ERROR] " + $Message
    }
    elseif ($INFO) {
        $Message = (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff") + " [INFO] " + $Message
    }

    Out-file -FilePath $file -InputObject $Message -Encoding UTF8 -Append
}

try {
    $pptPath = Get-Path -FILE -Message "PowerPointのパスを入力するか、ファイルをドラッグ＆ドロップしてください。"

    # ファイル拡張子を確認する
    $ext = [System.IO.Path]::GetExtension($pptPath);
    if ($ext -ne ".pptx" -and $ext -ne ".ppt") {
        throw "ファイル拡張子がpptxまたはpptではないため、処理を中断します"
    }

    # pptxをコピーしてzipに変更する
    $tmp = Join-Path -Path $WORK_PATH -ChildPath 'tmp.pptx'
    Copy-Item -Path $pptPath -Destination $tmp
    $zipFilePath = [System.IO.Path]::ChangeExtension($tmp, "zip")
    if (Test-Path $zipFilePath) {
        Remove-Item -Path $zipFilePath -Force
    }
    [System.IO.File]::Move($tmp, $zipFilePath)

    # 解凍処理
    $extractPath = $WORK_PATH + "/unzipped"
    $extractPath = Join-Path -path $WORK_PATH -ChildPath "unzipped"
    if (Test-Path $extractPath) {
        Remove-Item -Recurse -Force $extractPath
    }
    Expand-Archive -Path $zipFilePath -DestinationPath $extractPath

    # ノートのパス
    $folderPath = $PSScriptRoot + '\unzipped\ppt\notesSlides'

    # PowerPointから抽出したいノート部分を出力するファイルのパス
    $outputFilePath1 = $WORK_PATH + '\notes.txt'
    $outputFilePath2 = $WORK_PATH + '\strikes.txt'
    if (Test-Path -Path $outputFilePath1) {
        Remove-Item -Path $outputFilePath1 -Force
    }
    if (Test-Path -Path $outputFilePath2) {
        Remove-Item -Path $outputFilePath2 -Force
    }

    # フォルダ内のノート（XMLファイル）を取得
    Get-ChildItem -Path $folderPath -Filter 'notesSlide*.xml' | Sort-Object {
        # ファイル名にある数字で並び替え
        [regex]::Replace($_.BaseName, '\d+', { $args[0].Value.PadLeft(5) })
    } | ForEach-Object {
        $filePath = $_.FullName
        $fileContent = Get-Content -Path $filePath -Raw -Encoding UTF8

        # 正規表現でノートに書いてる文字を抽出
        $pattern = "<a:t>(.*?)</a:t>"
        [regex]::Matches($fileContent, $pattern) | ForEach-Object {
            $matchValue = $_.Groups[1].Value
            $totalCharCount += $matchValue.Length

            Out-File -FilePath $outputFilePath1 -InputObject $matchValue -Append -Encoding UTF8 -NoNewline
        }

        # 正規表現でノートに書いてある取り消し線の文字を抽出
        $strikePattern = "<a:rPr(.*?)strike=""sngStrike""(.*?)/>(.*?)<a:t>(.*?)</a:t>"
        [regex]::Matches($fileContent, $strikePattern) | ForEach-Object {
            $matchValue = $_.Groups[4].Value
            $totalCharCount -= $matchValue.Length

            Out-File -FilePath $outputFilePath2 -InputObject $matchValue -Append -Encoding UTF8 -NoNewline
        }
    }
    # 合計文字数（ノートの文字から取り消し線の文字を引いた数）を出力
    Write-Host "取り消し線を除く合計文字数は $totalCharCount 文字"
}
catch {
    Out-Log -ERR -Message "異常終了。処理を中断します。"
    Out-Log -Message $Error
    Write-Host "問題が発生したため処理を中断します。PowerShell スクリプト があるフォルダにある error.log にエラー内容を出力しましたので確認してください。"
    exit 1
}
finally {
    if ($null -ne $zipFilePath -and (Test-Path -Path $zipFilePath)) {
        Remove-Item -Path $zipFilePath -Force
    }
    if ($null -ne $extractPath -and (Test-Path -Path $extractPath)) {
        Remove-Item -Path $extractPath -Force -Recurse
    }
    if ($null -ne $totalCharCount) {
        $totalCharCount = $null
    }
    pause
}

