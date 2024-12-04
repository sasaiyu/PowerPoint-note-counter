$ErrorActionPreference = 'Stop'

$WORK_PATH = $PSScriptRoot

function Get-Path {
    param (
        [string]$Message,
        [switch]$FILE,
        [switch]$DIR
    )
    if ($null -eq $Message) {
        $Message = "��΃p�X�܂��� PowerShell �X�N���v�g����̑��΃p�X����͂��Ă�������"
    }
    $value = Read-Host $Message

    if (($null -eq $value) -or ($value.Length -eq 0)) {
        return Get-Path -Message "�����͂ł��B�ēx�A���͂��Ă�������"
    }

    if (-not (Test-Path -Path $value)) {
        $tmp = Join-Path -Path $WORK_PATH -ChildPath $value
        if (Test-Path -Path $tmp) {
            $value = $tmp
        }
        else {
            return Get-Path -Message "�t�@�C���܂��̓t�H���_�����݂��܂���B�ēx�A���͂��Ă�������"
        }
    }
    if ($FILE -and -not (Get-ChildItem -Path $value -File)) {
        return Get-Path -Message "�w�肳�ꂽ�p�X���t�@�C���p�X�ł͂���܂���B�ēx�A���͂��Ă�������"
    }
    if ($DIR -and -not (Get-ChildItem -Path $value -Directory)) {
        return Get-Path -Message "�w�肳�ꂽ�p�X���t�H���_�p�X�ł͂���܂���B�ēx�A���͂��Ă�������"
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
    $pptPath = Get-Path -FILE -Message "PowerPoint�̃p�X����͂��邩�A�t�@�C�����h���b�O���h���b�v���Ă��������B"

    # �t�@�C���g���q���m�F����
    $ext = [System.IO.Path]::GetExtension($pptPath);
    if ($ext -ne ".pptx" -and $ext -ne ".ppt") {
        throw "�t�@�C���g���q��pptx�܂���ppt�ł͂Ȃ����߁A�����𒆒f���܂�"
    }

    # pptx���R�s�[����zip�ɕύX����
    $tmp = Join-Path -Path $WORK_PATH -ChildPath 'tmp.pptx'
    Copy-Item -Path $pptPath -Destination $tmp
    $zipFilePath = [System.IO.Path]::ChangeExtension($tmp, "zip")
    if (Test-Path $zipFilePath) {
        Remove-Item -Path $zipFilePath -Force
    }
    [System.IO.File]::Move($tmp, $zipFilePath)

    # �𓀏���
    $extractPath = $WORK_PATH + "/unzipped"
    $extractPath = Join-Path -path $WORK_PATH -ChildPath "unzipped"
    if (Test-Path $extractPath) {
        Remove-Item -Recurse -Force $extractPath
    }
    Expand-Archive -Path $zipFilePath -DestinationPath $extractPath

    # �m�[�g�̃p�X
    $folderPath = $PSScriptRoot + '\unzipped\ppt\notesSlides'

    # PowerPoint���璊�o�������m�[�g�������o�͂���t�@�C���̃p�X
    $outputFilePath1 = $WORK_PATH + '\notes.txt'
    $outputFilePath2 = $WORK_PATH + '\strikes.txt'
    if (Test-Path -Path $outputFilePath1) {
        Remove-Item -Path $outputFilePath1 -Force
    }
    if (Test-Path -Path $outputFilePath2) {
        Remove-Item -Path $outputFilePath2 -Force
    }

    # �t�H���_���̃m�[�g�iXML�t�@�C���j���擾
    Get-ChildItem -Path $folderPath -Filter 'notesSlide*.xml' | Sort-Object {
        # �t�@�C�����ɂ��鐔���ŕ��ёւ�
        [regex]::Replace($_.BaseName, '\d+', { $args[0].Value.PadLeft(5) })
    } | ForEach-Object {
        $filePath = $_.FullName
        $fileContent = Get-Content -Path $filePath -Raw -Encoding UTF8

        # ���K�\���Ńm�[�g�ɏ����Ă镶���𒊏o
        $pattern = "<a:t>(.*?)</a:t>"
        [regex]::Matches($fileContent, $pattern) | ForEach-Object {
            $matchValue = $_.Groups[1].Value
            $totalCharCount += $matchValue.Length

            Out-File -FilePath $outputFilePath1 -InputObject $matchValue -Append -Encoding UTF8 -NoNewline
        }

        # ���K�\���Ńm�[�g�ɏ����Ă�����������̕����𒊏o
        $strikePattern = "<a:rPr(.*?)strike=""sngStrike""(.*?)/>(.*?)<a:t>(.*?)</a:t>"
        [regex]::Matches($fileContent, $strikePattern) | ForEach-Object {
            $matchValue = $_.Groups[4].Value
            $totalCharCount -= $matchValue.Length

            Out-File -FilePath $outputFilePath2 -InputObject $matchValue -Append -Encoding UTF8 -NoNewline
        }
    }
    # ���v�������i�m�[�g�̕���������������̕��������������j���o��
    Write-Host "�����������������v�������� $totalCharCount ����"
}
catch {
    Out-Log -ERR -Message "�ُ�I���B�����𒆒f���܂��B"
    Out-Log -Message $Error
    Write-Host "��肪�����������ߏ����𒆒f���܂��BPowerShell �X�N���v�g ������t�H���_�ɂ��� error.log �ɃG���[���e���o�͂��܂����̂Ŋm�F���Ă��������B"
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

