#Requires -Version 5.1
<#
.SYNOPSIS
    Synchronizes MappingTable.xlsx with the current Word template and Excel data.

.DESCRIPTION
    Scans the Word template for labels and compares with existing mappings.
    - Keeps existing mappings if label still exists
    - Adds new labels (with empty ExcelColumn for manual mapping)
    - Removes mappings for labels no longer in template

.PARAMETER ConfigPath
    Path to the configuration file. Defaults to config.psd1 in script directory.

.PARAMETER ShowColumns
    Lists all available data columns with their mapped/unmapped status.

.EXAMPLE
    .\Update-MappingTable.ps1

.EXAMPLE
    .\Update-MappingTable.ps1 -ShowColumns
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "",
    [switch]$ShowColumns
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$sectionHeaders = @(
    'Architektur allgemein',
    'Architektur Fenster',
    'Architektur Sonnenschutz',
    'Architektur Boden',
    'Architektur Wand',
    'Architektur Decke',
    'Beleuchtung',
    'Elektro Schwachstrom',
    'Elektro Starkstrom',
    'Elektro Sicherheit',
    'Heizung / Kälte',
    'Lüftung / Klima',
    'Medizinalgas',
    'Medizinaltechnik',
    'Sanitär',
    'x'
)

function Load-Configuration {
    param([string]$Path)

    if ([string]::IsNullOrEmpty($Path)) {
        $Path = Join-Path $scriptDir "config.psd1"
    }

    if (-not (Test-Path $Path)) {
        throw "Configuration file not found: $Path"
    }

    $config = Import-PowerShellDataFile -Path $Path

    if (-not [System.IO.Path]::IsPathRooted($config.TemplateFile)) {
        $config.TemplateFile = Join-Path $scriptDir $config.TemplateFile
    }
    if (-not [System.IO.Path]::IsPathRooted($config.DataFile)) {
        $config.DataFile = Join-Path $scriptDir $config.DataFile
    }
    if (-not [System.IO.Path]::IsPathRooted($config.MappingFile)) {
        $config.MappingFile = Join-Path $scriptDir $config.MappingFile
    }

    return $config
}

function Get-TemplateLabels {
    param([string]$TemplatePath)

    $word = $null
    $doc = $null
    $labels = @()

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        $doc = $word.Documents.Open($TemplatePath, $false, $true)

        foreach ($table in $doc.Tables) {
            for ($row = 1; $row -le $table.Rows.Count; $row++) {
                try {
                    $cellText = $table.Cell($row, 1).Range.Text
                    $cellText = $cellText -replace "`r`a", "" -replace "`r", "" -replace "`a", ""
                    $cellText = $cellText.Trim()

                    if (-not [string]::IsNullOrWhiteSpace($cellText) -and
                        $cellText.Length -ge 3 -and
                        $cellText -notin $sectionHeaders) {
                        $labels += $cellText
                    }
                }
                catch {
                    continue
                }
            }
        }

        return $labels
    }
    finally {
        if ($null -ne $doc) {
            $doc.Close([ref]$false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        if ($null -ne $word) {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Get-ExcelHeaders {
    param([string]$DataPath)

    $data = Import-Excel -Path $DataPath -StartRow 1
    return $data[0].PSObject.Properties.Name
}

try {
    Write-Host "=" * 60 -ForegroundColor Yellow
    Write-Host "Update Mapping Table" -ForegroundColor Yellow
    Write-Host "=" * 60 -ForegroundColor Yellow
    Write-Host ""

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "ImportExcel module is required. Install with: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel

    Write-Host "Loading configuration..." -ForegroundColor Cyan
    $config = Load-Configuration -Path $ConfigPath
    Write-Host "  Template: $($config.TemplateFile)"
    Write-Host "  Data: $($config.DataFile)"
    Write-Host "  Mapping: $($config.MappingFile)"
    Write-Host ""

    Write-Host "Extracting template labels..." -ForegroundColor Cyan
    $templateLabels = Get-TemplateLabels -TemplatePath $config.TemplateFile
    Write-Host "  Found $($templateLabels.Count) labels in template"
    Write-Host ""

    Write-Host "Extracting Excel headers..." -ForegroundColor Cyan
    $excelHeaders = Get-ExcelHeaders -DataPath $config.DataFile
    Write-Host "  Found $($excelHeaders.Count) columns in data file"
    Write-Host ""

    Write-Host "Loading existing mappings..." -ForegroundColor Cyan
    $existingMappings = @()
    if (Test-Path $config.MappingFile) {
        $existingMappings = @(Import-Excel -Path $config.MappingFile)
        Write-Host "  Found $($existingMappings.Count) existing mappings"
    }
    else {
        Write-Host "  No existing mapping file, creating new one"
    }
    Write-Host ""

    Write-Host "Synchronizing..." -ForegroundColor Cyan

    $kept = 0
    $added = 0
    $removed = 0
    $updatedMappings = @()

    foreach ($label in $templateLabels) {
        $existing = $existingMappings | Where-Object { $_.WordLabel -eq $label } | Select-Object -First 1

        if ($existing) {
            $updatedMappings += [PSCustomObject]@{
                ExcelColumn = $existing.ExcelColumn
                WordLabel = $label
            }
            $kept++
        }
        else {
            $updatedMappings += [PSCustomObject]@{
                ExcelColumn = ""
                WordLabel = $label
            }
            $added++
            Write-Host "  + Added: $label" -ForegroundColor Green
        }
    }

    $removedLabels = $existingMappings | Where-Object { $_.WordLabel -notin $templateLabels }
    foreach ($r in $removedLabels) {
        if (-not [string]::IsNullOrWhiteSpace($r.WordLabel)) {
            Write-Host "  - Removed: $($r.WordLabel)" -ForegroundColor Red
            $removed++
        }
    }

    $unmapped = ($updatedMappings | Where-Object { [string]::IsNullOrWhiteSpace($_.ExcelColumn) }).Count

    Write-Host ""
    Write-Host "Saving mapping table..." -ForegroundColor Cyan
    $updatedMappings | Export-Excel -Path $config.MappingFile -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet
    Write-Host "  Saved to: $($config.MappingFile)"

    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Yellow
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  Kept: $kept" -ForegroundColor Green
    Write-Host "  Added: $added" -ForegroundColor $(if ($added -gt 0) { "Yellow" } else { "Green" })
    Write-Host "  Removed: $removed" -ForegroundColor $(if ($removed -gt 0) { "Yellow" } else { "Green" })
    Write-Host "  Unmapped: $unmapped" -ForegroundColor $(if ($unmapped -gt 0) { "Yellow" } else { "Green" })
    Write-Host "=" * 60 -ForegroundColor Yellow

    if ($ShowColumns) {
        Write-Host ""
        Write-Host "Available data columns ($($excelHeaders.Count) total):" -ForegroundColor Cyan
        foreach ($col in $excelHeaders) {
            $isMapped = $updatedMappings | Where-Object { $_.ExcelColumn -eq $col }
            $status = if ($isMapped) { "[MAPPED]  " } else { "[UNMAPPED]" }
            Write-Host "  $status $col"
        }
    }

    if ($unmapped -gt 0) {
        Write-Host ""
        Write-Host "ACTION REQUIRED: $unmapped label(s) need ExcelColumn mapping" -ForegroundColor Yellow
        Write-Host "Open MappingTable.xlsx and fill in the empty ExcelColumn cells"
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    exit 1
}
