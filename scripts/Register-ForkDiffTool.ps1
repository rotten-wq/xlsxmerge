<#
.SYNOPSIS
    Fork Git Client에 XlsxMerge를 외부 Diff/Merge 도구로 등록합니다.

.DESCRIPTION
    이 스크립트는 Fork의 settings.json을 수정하여
    XlsxMerge.exe를 xlsx 파일용 외부 Diff 도구 및 Merge 도구로 등록합니다.

    - ExternalDiffTools: xlsx 파일 비교 시 사용
    - ExternalMergeTools: xlsx 파일 병합 충돌 해결 시 사용

    Fork가 실행 중이면 설정이 덮어씌워질 수 있으므로,
    Fork를 종료한 상태에서 실행하는 것을 권장합니다.

.PARAMETER ExePath
    XlsxMerge.exe의 전체 경로. 지정하지 않으면 이 스크립트와 같은 폴더의 XlsxMerge.exe를 사용합니다.

.PARAMETER Uninstall
    등록을 해제합니다.

.EXAMPLE
    .\Register-ForkDiffTool.ps1
    .\Register-ForkDiffTool.ps1 -ExePath "D:\Tools\XlsxMerge\XlsxMerge.exe"
    .\Register-ForkDiffTool.ps1 -Uninstall
#>

param(
    [string]$ExePath,
    [switch]$Uninstall,
    [switch]$Silent   # 인스톨러 등 비대화형 환경에서 사용: Fork 자동 종료, 프롬프트 없음
)

$ErrorActionPreference = "Stop"

# ── Fork settings.json 경로 ──
$forkSettingsPath = Join-Path $env:LOCALAPPDATA "Fork\settings.json"

if (-not (Test-Path $forkSettingsPath)) {
    Write-Host "[ERROR] Fork settings.json not found: $forkSettingsPath" -ForegroundColor Red
    Write-Host "        Fork가 설치되어 있는지 확인해주세요." -ForegroundColor Yellow
    exit 1
}

# ── XlsxMerge.exe 경로 결정 ──
if (-not $ExePath) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $ExePath = Join-Path $scriptDir "XlsxMerge.exe"
}

if (-not $Uninstall -and -not (Test-Path $ExePath)) {
    Write-Host "[ERROR] XlsxMerge.exe not found: $ExePath" -ForegroundColor Red
    Write-Host "        -ExePath 파라미터로 정확한 경로를 지정해주세요." -ForegroundColor Yellow
    exit 1
}

$exePathEscaped = $ExePath.Replace('\', '\\')

# ── Fork 실행 여부 확인 ──
$forkProcess = Get-Process -Name "Fork" -ErrorAction SilentlyContinue
if ($forkProcess) {
    if ($Silent) {
        # 비대화형 모드: Fork 자동 종료 (인스톨러 등)
        Write-Host "[INFO] Fork가 실행 중입니다. 자동으로 종료합니다." -ForegroundColor Yellow
        $forkProcess | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host "[INFO] Fork를 종료했습니다." -ForegroundColor Green
    } else {
        Write-Host "[WARNING] Fork가 실행 중입니다. 설정이 덮어씌워질 수 있습니다." -ForegroundColor Yellow
        $answer = Read-Host "Fork를 종료하고 계속할까요? (Y/N)"
        if ($answer -eq 'Y' -or $answer -eq 'y') {
            $forkProcess | Stop-Process -Force
            Start-Sleep -Seconds 2
            Write-Host "  Fork를 종료했습니다." -ForegroundColor Green
        } else {
            Write-Host "  스크립트를 중단합니다. Fork를 종료한 후 다시 실행해주세요." -ForegroundColor Yellow
            exit 0
        }
    }
}

# ── settings.json 백업 ──
$backupPath = "$forkSettingsPath.bak"
Copy-Item $forkSettingsPath $backupPath -Force
Write-Host "[INFO] 백업 생성: $backupPath" -ForegroundColor Cyan

# ── settings.json 로드 ──
$json = Get-Content $forkSettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json

# ── Diff Tool 정의 ──
$diffToolEntry = [PSCustomObject]@{
    Type      = "Custom"
    Name      = "XlsxMerge"
    Path      = $ExePath
    Arguments = '-order=bd "$REMOTE" "$LOCAL"'
}

# ── Merge Tool 정의 ──
$mergeToolEntry = [PSCustomObject]@{
    Type      = "Custom"
    Name      = "XlsxMerge"
    Path      = $ExePath
    Arguments = '-b="$BASE" -d="$LOCAL" -s="$REMOTE" -r="$MERGED"'
}

# ── 기존 XlsxMerge 항목 제거 (있으면) ──
function Remove-XlsxMergeEntries($list) {
    if ($null -eq $list) { return @() }
    return @($list | Where-Object { $_.Name -ne "XlsxMerge" })
}

if ($Uninstall) {
    # ── 등록 해제 ──
    $json.ExternalDiffTools  = Remove-XlsxMergeEntries $json.ExternalDiffTools
    $json.ExternalMergeTools = Remove-XlsxMergeEntries $json.ExternalMergeTools

    $json | ConvertTo-Json -Depth 10 | Set-Content $forkSettingsPath -Encoding UTF8
    Write-Host ""
    Write-Host "=== XlsxMerge 등록 해제 완료 ===" -ForegroundColor Green
    Write-Host "  Fork를 다시 시작하면 반영됩니다." -ForegroundColor Cyan
} else {
    # ── 등록 ──
    # 기존 항목 정리 후 추가
    $diffTools  = Remove-XlsxMergeEntries $json.ExternalDiffTools
    $mergeTools = Remove-XlsxMergeEntries $json.ExternalMergeTools

    $diffTools  += $diffToolEntry
    $mergeTools += $mergeToolEntry

    $json.ExternalDiffTools  = $diffTools
    $json.ExternalMergeTools = $mergeTools

    $json | ConvertTo-Json -Depth 10 | Set-Content $forkSettingsPath -Encoding UTF8

    Write-Host ""
    Write-Host "=== XlsxMerge Fork 등록 완료 ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "  [Diff Tool]"  -ForegroundColor White
    Write-Host "    Name: XlsxMerge" -ForegroundColor Cyan
    Write-Host "    Path: $ExePath" -ForegroundColor Cyan
    Write-Host "    Args: -order=bd `"`$REMOTE`" `"`$LOCAL`"" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [Merge Tool]" -ForegroundColor White
    Write-Host "    Name: XlsxMerge" -ForegroundColor Cyan
    Write-Host "    Path: $ExePath" -ForegroundColor Cyan
    Write-Host '    Args: -b="$BASE" -d="$LOCAL" -s="$REMOTE" -r="$MERGED"' -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Fork > File > Preferences > Integration 에서 확인할 수 있습니다." -ForegroundColor Yellow
    Write-Host "  Fork를 다시 시작하면 반영됩니다." -ForegroundColor Yellow
}
