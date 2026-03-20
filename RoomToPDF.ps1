#Requires -Version 5.1
<#
.SYNOPSIS
    Converts Excel room data to PDF specification sheets using a Word template.

.PARAMETER ConfigPath
    Path to the configuration file. Defaults to config.psd1 in script directory.

.PARAMETER RoomCode
    Process only a specific room by its code (e.g., "RT.017").

.EXAMPLE
    .\RoomToPDF.ps1

.EXAMPLE
    .\RoomToPDF.ps1 -RoomCode "RT.017"
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "",
    [string]$RoomCode = ""
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Load-Configuration {
    param([string]$Path)

    if ([string]::IsNullOrEmpty($Path)) {
        $Path = Join-Path $scriptDir "config.psd1"
    }

    if (-not (Test-Path $Path)) {
        throw "Configuration file not found: $Path"
    }

    $config = Import-PowerShellDataFile -Path $Path

    $config.DataFile = Resolve-ConfigPath $config.DataFile
    $config.TemplateFile = Resolve-ConfigPath $config.TemplateFile
    $config.MappingFile = Resolve-ConfigPath $config.MappingFile
    $config.OutputFolder = Resolve-ConfigPath $config.OutputFolder

    if (-not (Test-Path $config.DataFile)) {
        throw "Data file not found: $($config.DataFile)"
    }
    if (-not (Test-Path $config.TemplateFile)) {
        throw "Template file not found: $($config.TemplateFile)"
    }
    if (-not (Test-Path $config.MappingFile)) {
        throw "Mapping file not found: $($config.MappingFile)"
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $config.OutputFolder = Join-Path (Resolve-ConfigPath $config.OutputFolder) $timestamp
    New-Item -ItemType Directory -Path $config.OutputFolder -Force | Out-Null

    return $config
}

function Resolve-ConfigPath {
    param([string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }
    return Join-Path $scriptDir $Path
}

function Read-MappingTable {
    param([string]$Path)

    $mappings = Import-Excel -Path $Path
    $result = @()

    foreach ($row in $mappings) {
        if ([string]::IsNullOrWhiteSpace($row.ExcelColumn)) {
            continue
        }

        $result += [PSCustomObject]@{
            ExcelColumn = $row.ExcelColumn.Trim()
            WordLabel = $row.WordLabel.Trim()
            UnitSuffix = if ($row.UnitSuffix) { $row.UnitSuffix.Trim() } else { "" }
        }
    }

    Write-Host "Loaded $($result.Count) mappings from mapping table" -ForegroundColor Cyan
    return $result
}

function Read-RoomData {
    param(
        [string]$Path,
        [string]$RoomCodeColumn
    )

    $data = Import-Excel -Path $Path
    Write-Host "Loaded $($data.Count) rooms from data file" -ForegroundColor Cyan
    return $data
}

function Normalize-Text {
    param([string]$Text)

    # Handle German umlauts and encoding inconsistencies between Excel and Word
    $result = $Text.ToLower().Trim()
    $result = $result -replace 'ä', 'ae'
    $result = $result -replace 'ö', 'oe'
    $result = $result -replace 'ü', 'ue'
    $result = $result -replace 'ß', 'ss'
    $result = $result -replace '\xc3\xa4', 'ae'
    $result = $result -replace '\xc3\xb6', 'oe'
    $result = $result -replace '\xc3\xbc', 'ue'
    $result = $result -replace '\ufffd', ''
    $result = $result -replace '\?', ''
    $result = $result -replace '[^\x20-\x7E]', ''
    return $result
}

function Build-LabelPositionCache {
    param(
        [string]$TemplatePath,
        [array]$Mappings
    )

    $word = $null
    $doc = $null
    $cache = @{}

    try {
        Write-Host "Building label position cache..." -ForegroundColor Cyan

        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        $doc = $word.Documents.Open($TemplatePath, $false, $true)

        $labelsToFind = @{}
        foreach ($mapping in $Mappings) {
            $normalizedLabel = Normalize-Text $mapping.WordLabel
            $labelsToFind[$normalizedLabel] = $mapping.WordLabel
        }

        $tableIndex = 0
        foreach ($table in $doc.Tables) {
            $tableIndex++
            $rowCount = $table.Rows.Count
            $isSingleColumn = $table.Columns.Count -eq 1

            for ($row = 1; $row -le $rowCount; $row++) {
                try {
                    $cellText = $table.Cell($row, 1).Range.Text
                    $cellText = $cellText -replace "`r`a", "" -replace "`r", "" -replace "`a", ""
                    $cellText = $cellText.Trim()

                    $normalizedCell = Normalize-Text $cellText

                    if ($labelsToFind.ContainsKey($normalizedCell)) {
                        if ($isSingleColumn) {
                            $cache[$normalizedCell] = @{
                                TableIndex = $tableIndex
                                Row = $row + 1
                                Col = 1
                            }
                        }
                        else {
                            $cache[$normalizedCell] = @{
                                TableIndex = $tableIndex
                                Row = $row
                                Col = 2
                            }
                        }
                    }
                }
                catch {
                    continue
                }
            }
        }

        Write-Host "  Cached $($cache.Count) label positions from $tableIndex tables" -ForegroundColor Green

        $notFound = $labelsToFind.Keys | Where-Object { -not $cache.ContainsKey($_) }
        if ($notFound.Count -gt 0) {
            Write-Host "  Warning: $($notFound.Count) labels not found in template" -ForegroundColor Yellow
        }

        return $cache
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

function Set-CellValue {
    param(
        [object]$Cell,
        [string]$Value,
        [string]$UnitSuffix
    )

    if ($null -eq $Cell) { return $false }

    if ($Value -eq "TRUE" -or $Value -eq "True") { $Value = "ja" }
    elseif ($Value -eq "FALSE" -or $Value -eq "False") { $Value = "nein" }

    $existingText = $Cell.Range.Text -replace "`r`a", "" -replace "`r", "" -replace "`a", ""
    $existingText = $existingText.Trim()

    $finalValue = $Value
    if (-not [string]::IsNullOrEmpty($UnitSuffix)) {
        $finalValue = "$Value $UnitSuffix"
    }
    elseif ($existingText -match '^\s*(°C|lux|%|Stk\.|m²|m³|Pa|W|W/m²|kW|l/s|m³/h)$') {
        $finalValue = "$Value $existingText"
    }

    $Cell.Range.Text = $finalValue
    return $true
}

function Process-SingleRoom {
    param(
        [object]$WordApp,
        [string]$TemplatePath,
        [PSCustomObject]$RoomData,
        [array]$Mappings,
        [string]$OutputPath,
        [string]$RoomCode,
        [hashtable]$PositionCache
    )

    $doc = $null
    $success = $false
    $warnings = @()

    try {
        $doc = $WordApp.Documents.Open($TemplatePath, $false, $true)

        foreach ($mapping in $Mappings) {
            $excelValue = $RoomData.($mapping.ExcelColumn)

            if ([string]::IsNullOrWhiteSpace($excelValue)) {
                continue
            }

            $cell = $null
            $normalizedLabel = Normalize-Text $mapping.WordLabel

            if ($PositionCache.ContainsKey($normalizedLabel)) {
                $pos = $PositionCache[$normalizedLabel]
                try {
                    $cell = $doc.Tables.Item($pos.TableIndex).Cell($pos.Row, $pos.Col)
                }
                catch {
                    $warnings += "Cache miss: $($mapping.WordLabel)"
                }
            }

            if ($null -eq $cell) {
                $warnings += "Label not found: $($mapping.WordLabel)"
                continue
            }

            Set-CellValue -Cell $cell -Value $excelValue -UnitSuffix $mapping.UnitSuffix | Out-Null
        }

        $doc.ExportAsFixedFormat($OutputPath, 17) # wdExportFormatPDF
        $success = $true
    }
    catch {
        $warnings += "Error: $_"
    }
    finally {
        if ($null -ne $doc) {
            $doc.Close([ref]$false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
    }

    return @{ Success = $success; Warnings = $warnings }
}

function Release-ComObjects {
    param([object]$WordApp)

    if ($null -ne $WordApp) {
        try {
            $WordApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApp) | Out-Null
        }
        catch { }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$word = $null

try {
    Write-Host "=" * 60 -ForegroundColor Yellow
    Write-Host "Room Database to PDF Converter" -ForegroundColor Yellow
    Write-Host "=" * 60 -ForegroundColor Yellow
    Write-Host ""

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "ImportExcel module is required. Install with: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel

    Write-Host "Loading configuration..." -ForegroundColor Cyan
    $config = Load-Configuration -Path $ConfigPath
    Write-Host "  Data file: $($config.DataFile)"
    Write-Host "  Template: $($config.TemplateFile)"
    Write-Host "  Output: $($config.OutputFolder)"
    Write-Host ""

    $mappings = Read-MappingTable -Path $config.MappingFile
    $roomData = Read-RoomData -Path $config.DataFile -RoomCodeColumn $config.RoomCodeColumn
    Write-Host ""

    if (-not [string]::IsNullOrEmpty($RoomCode)) {
        $roomData = $roomData | Where-Object { $_.$($config.RoomCodeColumn) -eq $RoomCode }
        if ($roomData.Count -eq 0) {
            throw "Room not found: $RoomCode"
        }
        Write-Host "Filtering to room: $RoomCode" -ForegroundColor Cyan
        Write-Host ""
    }

    $startTotal = Get-Date
    $positionCache = Build-LabelPositionCache -TemplatePath $config.TemplateFile -Mappings $mappings
    $cacheTime = (Get-Date) - $startTotal
    Write-Host "  Cache built in $([math]::Round($cacheTime.TotalSeconds, 1))s" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Starting Word..." -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $totalRooms = @($roomData).Count
    $successCount = 0
    $failCount = 0
    $i = 0

    Write-Host ""
    Write-Host "Processing $totalRooms rooms..." -ForegroundColor Yellow
    Write-Host "-" * 60 -ForegroundColor Yellow

    $startProcess = Get-Date

    foreach ($room in $roomData) {
        $i++
        $roomCode = $room.$($config.RoomCodeColumn)

        if ([string]::IsNullOrWhiteSpace($roomCode)) {
            $failCount++
            continue
        }

        $safeRoomCode = $roomCode -replace '[\\/:*?"<>|]', '_'
        $outputPath = Join-Path $config.OutputFolder "$safeRoomCode.pdf"

        Write-Host "[$i/$totalRooms] $roomCode..." -NoNewline

        $result = Process-SingleRoom `
            -WordApp $word `
            -TemplatePath $config.TemplateFile `
            -RoomData $room `
            -Mappings $mappings `
            -OutputPath $outputPath `
            -RoomCode $roomCode `
            -PositionCache $positionCache

        if ($result.Success) {
            $successCount++
            Write-Host " OK" -ForegroundColor Green
        }
        else {
            $failCount++
            Write-Host " FAILED" -ForegroundColor Red
            foreach ($w in $result.Warnings) {
                Write-Host "    $w" -ForegroundColor Red
            }
        }
    }

    $processTime = (Get-Date) - $startProcess
    $totalTime = (Get-Date) - $startTotal

    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Yellow
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  Total rooms: $totalRooms"
    Write-Host "  Successful: $successCount" -ForegroundColor Green
    Write-Host "  Failed: $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
    Write-Host ""
    Write-Host "Performance:" -ForegroundColor Yellow
    Write-Host "  Cache build: $([math]::Round($cacheTime.TotalSeconds, 1))s"
    $perRoom = if ($successCount -gt 0) { [math]::Round($processTime.TotalSeconds / $successCount, 1) } else { 0 }
    Write-Host "  Processing: $([math]::Round($processTime.TotalMinutes, 2)) min ($perRoom s/room)"
    Write-Host "  Total time: $([math]::Round($totalTime.TotalMinutes, 2)) min"
    Write-Host "=" * 60 -ForegroundColor Yellow
}
catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
    exit 1
}
finally {
    Write-Host ""
    Write-Host "Cleaning up..." -ForegroundColor Cyan
    Release-ComObjects -WordApp $word
    Write-Host "Done." -ForegroundColor Green
}
