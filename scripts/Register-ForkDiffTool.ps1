#Requires -Version 5.1
<#
.SYNOPSIS
    Registers XlsxMerge as an external Diff/Merge tool in Fork Git Client.

.PARAMETER ExePath
    Full path to XlsxMerge.exe. Defaults to the same folder as this script.

.PARAMETER Uninstall
    Removes XlsxMerge from Fork's tool list.

.PARAMETER Silent
    Auto-closes Fork if running, without prompting (for installer use).

.EXAMPLE
    .\Register-ForkDiffTool.ps1
    .\Register-ForkDiffTool.ps1 -ExePath "C:\Tools\XlsxMerge\XlsxMerge.exe"
    .\Register-ForkDiffTool.ps1 -Uninstall
#>
param(
    [string]$ExePath,
    [switch]$Uninstall,
    [switch]$Silent
)

$ErrorActionPreference = "Stop"

# Fork settings.json location
$forkSettingsPath = Join-Path $env:LOCALAPPDATA "Fork\settings.json"

if (-not (Test-Path $forkSettingsPath)) {
    Write-Warning "Fork settings.json not found: $forkSettingsPath"
    Write-Warning "Please make sure Fork is installed."
    exit 1
}

# Resolve XlsxMerge.exe path
if (-not $ExePath) {
    $ExePath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "XlsxMerge.exe"
}

if (-not $Uninstall -and -not (Test-Path $ExePath)) {
    Write-Warning "XlsxMerge.exe not found: $ExePath"
    exit 1
}

# Handle Fork process
$forkProcess = Get-Process -Name "Fork" -ErrorAction SilentlyContinue
if ($forkProcess) {
    if ($Silent) {
        Write-Host "[INFO] Fork is running - closing automatically..."
        $forkProcess | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host "[INFO] Fork closed."
    }
    else {
        Write-Host "[WARNING] Fork is currently running." -ForegroundColor Yellow
        $answer = Read-Host "Close Fork and continue? (Y/N)"
        if ($answer -match '^[Yy]') {
            $forkProcess | Stop-Process -Force
            Start-Sleep -Seconds 2
            Write-Host "Fork closed." -ForegroundColor Green
        }
        else {
            Write-Host "Cancelled. Close Fork first and run again." -ForegroundColor Yellow
            exit 0
        }
    }
}

# Read settings.json (handle UTF-8 BOM)
$rawJson = Get-Content $forkSettingsPath -Raw -Encoding UTF8
$json = $rawJson | ConvertFrom-Json

# Backup
$backupPath = "$forkSettingsPath.bak"
Copy-Item $forkSettingsPath $backupPath -Force
Write-Host "[INFO] Backup saved: $backupPath"

# Safely set a property (works whether the property exists or not)
function Set-JsonProp($obj, [string]$name, $value) {
    if ($obj.PSObject.Properties[$name]) {
        $obj.PSObject.Properties[$name].Value = $value
    }
    else {
        $obj | Add-Member -NotePropertyName $name -NotePropertyValue $value -Force
    }
}

# Remove existing XlsxMerge entries from a list
function Remove-ToolEntry($list) {
    if ($null -eq $list) { return @() }
    return @($list | Where-Object { $_.Name -ne "XlsxMerge" })
}

if ($Uninstall) {
    Set-JsonProp $json 'ExternalDiffTools'  (Remove-ToolEntry $json.ExternalDiffTools)
    Set-JsonProp $json 'ExternalMergeTools' (Remove-ToolEntry $json.ExternalMergeTools)

    if ($json.PSObject.Properties['ExternalDiffTool'] -and
        $json.ExternalDiffTool.ApplicationPath -like '*XlsxMerge*') {
        Set-JsonProp $json 'ExternalDiffTool' ([PSCustomObject]@{ Type = "Custom"; ApplicationPath = ""; Arguments = "" })
    }
    if ($json.PSObject.Properties['MergeTool'] -and
        $json.MergeTool.ApplicationPath -like '*XlsxMerge*') {
        Set-JsonProp $json 'MergeTool' ([PSCustomObject]@{ Type = "Custom"; ApplicationPath = ""; Arguments = "" })
    }

    $json | ConvertTo-Json -Depth 20 | Set-Content $forkSettingsPath -Encoding UTF8
    Write-Host ""
    Write-Host "=== XlsxMerge unregistered from Fork ===" -ForegroundColor Green
    Write-Host "Restart Fork to apply changes." -ForegroundColor Cyan
}
else {
    $diffArgs  = '-b="$LOCAL" -d="$REMOTE"'
    $mergeArgs = '-b="$BASE" -d="$LOCAL" -s="$REMOTE" -r="$MERGED"'

    # Build list entries as plain PSCustomObject literals (pipe+Add-Member drops ApplicationPath)
    $diffEntry = [PSCustomObject]@{
        Type            = "Custom"
        Name            = "XlsxMerge"
        ApplicationPath = $ExePath
        Arguments       = $diffArgs
    }
    $mergeEntry = [PSCustomObject]@{
        Type            = "Custom"
        Name            = "XlsxMerge"
        ApplicationPath = $ExePath
        Arguments       = $mergeArgs
    }

    # Update library lists
    $diffList  = @(Remove-ToolEntry $json.ExternalDiffTools)  + $diffEntry
    $mergeList = @(Remove-ToolEntry $json.ExternalMergeTools) + $mergeEntry

    Set-JsonProp $json 'ExternalDiffTools'  $diffList
    Set-JsonProp $json 'ExternalMergeTools' $mergeList

    # Set as active tool
    Set-JsonProp $json 'ExternalDiffTool' ([PSCustomObject]@{
        Type            = "Custom"
        ApplicationPath = $ExePath
        Arguments       = $diffArgs
    })
    Set-JsonProp $json 'MergeTool' ([PSCustomObject]@{
        Type            = "Custom"
        ApplicationPath = $ExePath
        Arguments       = $mergeArgs
    })

    $json | ConvertTo-Json -Depth 20 | Set-Content $forkSettingsPath -Encoding UTF8

    Write-Host ""
    Write-Host "=== XlsxMerge registered in Fork ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Diff  args: $diffArgs"  -ForegroundColor Cyan
    Write-Host "  Merge args: $mergeArgs" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Check: Fork > File > Preferences > Integration" -ForegroundColor Yellow
    Write-Host "Restart Fork to apply changes." -ForegroundColor Yellow
}
