#Requires -Version 5.1
<#
.SYNOPSIS
    Fork Git Client에 XlsxMerge를 외부 Diff/Merge 도구로 등록합니다.

.DESCRIPTION
    Fork의 settings.json을 수정하여 XlsxMerge.exe를 xlsx 파일용
    외부 Diff/Merge 도구로 등록합니다.

.PARAMETER ExePath
    XlsxMerge.exe 전체 경로. 생략 시 스크립트와 같은 폴더를 사용합니다.

.PARAMETER Uninstall
    등록을 해제합니다.

.PARAMETER Silent
    Fork 실행 중일 때 자동 종료 (인스톨러/비대화형 환경용).

.EXAMPLE
    .\Register-ForkDiffTool.ps1
    .\Register-ForkDiffTool.ps1 -ExePath "C:\Tools\XlsxMerge\XlsxMerge.exe"
    .\Register-ForkDiffTool.ps1 -Uninstall
#>

[OutputType([void])]
param(
    [string]$ExePath,
    [switch]$Uninstall,
    [switch]$Silent
)

$ErrorActionPreference = "Stop"

# ── Fork settings.json 경로 ──
$forkSettingsPath = Join-Path $env:LOCALAPPDATA "Fork\settings.json"

if (-not (Test-Path $forkSettingsPath)) {
    Write-Warning "Fork settings.json을 찾을 수 없습니다: $forkSettingsPath"
    Write-Warning "Fork가 설치되어 있는지 확인해주세요."
    exit 1
}

# ── XlsxMerge.exe 경로 결정 ──
if (-not $ExePath) {
    $ExePath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "XlsxMerge.exe"
}

if (-not $Uninstall -and -not (Test-Path $ExePath)) {
    Write-Warning "XlsxMerge.exe를 찾을 수 없습니다: $ExePath"
    exit 1
}

# ── Fork 실행 여부 확인 ──
$forkProcess = Get-Process -Name "Fork" -ErrorAction SilentlyContinue
if ($forkProcess) {
    if ($Silent) {
        Write-Host "[INFO] Fork 실행 중 - 자동 종료합니다..."
        $forkProcess | Stop-Process -Force
        Start-Sleep -Seconds 2
    }
    else {
        Write-Host "[WARNING] Fork가 실행 중입니다." -ForegroundColor Yellow
        $answer = Read-Host "Fork를 종료하고 계속하시겠습니까? (Y/N)"
        if ($answer -match '^[Yy]') {
            $forkProcess | Stop-Process -Force
            Start-Sleep -Seconds 2
            Write-Host "Fork를 종료했습니다." -ForegroundColor Green
        }
        else {
            Write-Host "취소합니다. Fork 종료 후 다시 실행해주세요." -ForegroundColor Yellow
            exit 0
        }
    }
}

# ── settings.json 읽기 (BOM 포함 UTF-8 대응) ──
$rawJson = Get-Content $forkSettingsPath -Raw -Encoding UTF8
$json = $rawJson | ConvertFrom-Json

# ── 백업 ──
$backupPath = "$forkSettingsPath.bak"
Copy-Item $forkSettingsPath $backupPath -Force
Write-Host "[INFO] 백업 생성: $backupPath"

# ── 도구 항목 생성 헬퍼 ──
function New-ToolEntry([string]$appPath, [string]$arguments) {
    [PSCustomObject]@{
        Type            = "Custom"
        ApplicationPath = $appPath
        Arguments       = $arguments
    }
}

# ── JSON 프로퍼티 안전 설정 (Add-Member -Force로 없는 프로퍼티도 생성) ──
function Set-JsonProp($obj, [string]$name, $value) {
    if ($obj.PSObject.Properties[$name]) {
        $obj.PSObject.Properties[$name].Value = $value
    }
    else {
        $obj | Add-Member -NotePropertyName $name -NotePropertyValue $value -Force
    }
}

# ── 기존 XlsxMerge 항목 제거 ──
function Remove-ToolEntry($list) {
    if ($null -eq $list) { return @() }
    $result = @($list | Where-Object { $_.Name -ne "XlsxMerge" })
    return $result
}

if ($Uninstall) {
    # 목록에서 제거
    Set-JsonProp $json 'ExternalDiffTools'  (Remove-ToolEntry $json.ExternalDiffTools)
    Set-JsonProp $json 'ExternalMergeTools' (Remove-ToolEntry $json.ExternalMergeTools)

    # 활성 도구가 XlsxMerge였으면 초기화
    if ($json.PSObject.Properties['ExternalDiffTool'] -and $json.ExternalDiffTool.ApplicationPath -like '*XlsxMerge*') {
        Set-JsonProp $json 'ExternalDiffTool' (New-ToolEntry '' '')
    }
    if ($json.PSObject.Properties['MergeTool'] -and $json.MergeTool.ApplicationPath -like '*XlsxMerge*') {
        Set-JsonProp $json 'MergeTool' (New-ToolEntry '' '')
    }

    $json | ConvertTo-Json -Depth 20 | Set-Content $forkSettingsPath -Encoding UTF8
    Write-Host ""
    Write-Host "=== XlsxMerge Fork 등록 해제 완료 ===" -ForegroundColor Green
    Write-Host "Fork를 다시 시작하면 반영됩니다." -ForegroundColor Cyan
}
else {
    # ── Diff 인자: -b=기준파일 -d=비교파일 ──
    $diffArgs  = "-b=`"`$LOCAL`" -d=`"`$REMOTE`""
    # ── Merge 인자: -b=공통조상 -d=내것 -s=상대것 -r=결과 ──
    $mergeArgs = "-b=`"`$BASE`" -d=`"`$LOCAL`" -s=`"`$REMOTE`" -r=`"`$MERGED`""

    $diffEntry  = New-ToolEntry $ExePath $diffArgs  | Add-Member -PassThru -NotePropertyName Name -NotePropertyValue "XlsxMerge"
    $mergeEntry = New-ToolEntry $ExePath $mergeArgs | Add-Member -PassThru -NotePropertyName Name -NotePropertyValue "XlsxMerge"

    # 목록에 추가 (중복 방지 위해 기존 항목 제거 후 추가)
    $diffList  = @(Remove-ToolEntry $json.ExternalDiffTools)  + $diffEntry
    $mergeList = @(Remove-ToolEntry $json.ExternalMergeTools) + $mergeEntry

    Set-JsonProp $json 'ExternalDiffTools'  $diffList
    Set-JsonProp $json 'ExternalMergeTools' $mergeList

    # 활성 도구도 XlsxMerge로 설정
    Set-JsonProp $json 'ExternalDiffTool' (New-ToolEntry $ExePath $diffArgs)
    Set-JsonProp $json 'MergeTool'        (New-ToolEntry $ExePath $mergeArgs)

    $json | ConvertTo-Json -Depth 20 | Set-Content $forkSettingsPath -Encoding UTF8

    Write-Host ""
    Write-Host "=== XlsxMerge Fork 등록 완료 ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Diff 인자 : $diffArgs"  -ForegroundColor Cyan
    Write-Host "  Merge 인자: $mergeArgs" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Fork > File > Preferences > Integration 에서 확인할 수 있습니다." -ForegroundColor Yellow
    Write-Host "Fork를 다시 시작하면 반영됩니다." -ForegroundColor Yellow
}
