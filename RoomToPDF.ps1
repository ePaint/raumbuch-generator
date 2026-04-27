#Requires -Version 5.1
<#
.SYNOPSIS
    Converts Excel room data to PDF specification sheets using a Word template.

.PARAMETER ConfigPath
    Path to the configuration file. Defaults to config.psd1 in script directory.

.PARAMETER RoomCode
    Process specific room(s) by code. Comma-separated for multiple (e.g., "RT.017,RT.018").

.PARAMETER Source
    Data source: 'Excel' or 'API'. Overrides config.psd1 setting.

.PARAMETER Template
    Path to Word template file. Overrides config.psd1 setting.

.PARAMETER ExcelFile
    Path to Excel data file. Overrides config.psd1 setting.

.EXAMPLE
    .\RoomToPDF.ps1

.EXAMPLE
    .\RoomToPDF.ps1 -Source API

.EXAMPLE
    .\RoomToPDF.ps1 -Template "Input/my-template.docx"

.EXAMPLE
    .\RoomToPDF.ps1 -RoomCode "RT.017"

.EXAMPLE
    .\RoomToPDF.ps1 -RoomCode "RT.001,RT.017,RT.187"

.EXAMPLE
    .\RoomToPDF.ps1 -Merge
    # Combines all rooms into a single PDF file

.EXAMPLE
    .\RoomToPDF.ps1 -RoomCode "RT.001,RT.017,RT.187" -Merge
    # Combines specified rooms into a single PDF file
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "",
    [string]$RoomCode = "",
    [ValidateSet('Excel', 'API')]
    [string]$Source = "",
    [string]$Template = "",
    [string]$ExcelFile = "",
    [switch]$Merge
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load assemblies for OpenXML processing (docx is a zip file)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Load-Configuration {
    param(
        [string]$Path,
        [string]$SourceOverride = "",
        [string]$TemplateOverride = "",
        [string]$ExcelFileOverride = ""
    )

    if ([string]::IsNullOrEmpty($Path)) {
        $Path = Join-Path $scriptDir "config.psd1"
    }

    if (-not (Test-Path $Path)) {
        throw "Configuration file not found: $Path"
    }

    $config = Import-PowerShellDataFile -Path $Path

    # Template override
    if ($TemplateOverride) {
        $config.TemplateFile = Resolve-ConfigPath $TemplateOverride
    } else {
        $config.TemplateFile = Resolve-ConfigPath $config.TemplateFile
    }
    $config.OutputFolder = Resolve-ConfigPath $config.OutputFolder

    # Parameter overrides config, config defaults to Excel
    $dataSource = if ($SourceOverride) { $SourceOverride } elseif ($config.DataSource) { $config.DataSource } else { 'Excel' }
    $config.DataSource = $dataSource

    if ($dataSource -eq 'Excel') {
        if ($ExcelFileOverride) {
            $config.Excel.DataFile = Resolve-ConfigPath $ExcelFileOverride
        } else {
            $config.Excel.DataFile = Resolve-ConfigPath $config.Excel.DataFile
        }
        if (-not (Test-Path $config.Excel.DataFile)) {
            throw "Data file not found: $($config.Excel.DataFile)"
        }
    }

    if (-not (Test-Path $config.TemplateFile)) {
        throw "Template file not found: $($config.TemplateFile)"
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

function Read-RoomData {
    param(
        [string]$Path,
        [string]$RoomCodeColumn
    )

    $data = Import-Excel -Path $Path
    Write-Host "Loaded $($data.Count) rooms from Excel file" -ForegroundColor Cyan
    return $data
}

function Read-RoomDataFromAPI {
    param(
        [hashtable]$APIConfig,
        [string]$ScriptDir
    )

    $keyPath = $APIConfig.KeyFile
    if (-not [System.IO.Path]::IsPathRooted($keyPath)) {
        $keyPath = Join-Path $ScriptDir $keyPath
    }

    if (-not (Test-Path $keyPath)) {
        throw "API key file not found: $keyPath"
    }

    $apiKey = (Get-Content $keyPath -Raw).Trim()

    # Get endpoint URL from file or direct value
    if ($APIConfig.EndpointFile) {
        $endpointPath = $APIConfig.EndpointFile
        if (-not [System.IO.Path]::IsPathRooted($endpointPath)) {
            $endpointPath = Join-Path $ScriptDir $endpointPath
        }
        $endpoint = (Get-Content $endpointPath -Raw).Trim()
    } else {
        $endpoint = $APIConfig.Endpoint
    }

    $headers = @{
        'Authorization' = "Reference $apiKey"
        'Accept' = 'application/json'
    }

    Write-Host "Fetching data from API..." -ForegroundColor Cyan
    $response = Invoke-RestMethod -Uri $endpoint -Headers $headers

    Write-Host "Loaded $($response.Count) rooms from API" -ForegroundColor Cyan
    return $response
}

function Get-TemplatePlaceholders {
    param(
        [object]$WordApp,
        [string]$TemplatePath
    )

    $doc = $WordApp.Documents.Open($TemplatePath, $false, $true)
    $text = $doc.Content.Text
    $doc.Close($false)

    $placeholders = @()
    $matches = [regex]::Matches($text, '<<([^>]+)>>')
    foreach ($m in $matches) {
        $placeholders += $m.Groups[1].Value
    }

    return $placeholders | Select-Object -Unique
}

function Process-RoomWithOpenXML {
    param(
        [string]$TemplatePath,
        [PSCustomObject]$RoomData,
        [string]$OutputPath,
        [hashtable]$ValueMap
    )

    # Copy template to output
    Copy-Item $TemplatePath $OutputPath -Force

    # Open docx as zip and replace placeholders in XML
    $zip = [System.IO.Compression.ZipFile]::Open($OutputPath, "Update")

    try {
        $entry = $zip.GetEntry("word/document.xml")
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()

        # Find all placeholders (XML-encoded angle brackets)
        $matches = [regex]::Matches($xml, "&lt;&lt;([^&]+)&gt;&gt;")

        foreach ($m in $matches) {
            $fieldName = $m.Groups[1].Value
            $placeholder = "&lt;&lt;$fieldName&gt;&gt;"
            $value = $RoomData.$fieldName

            if ($null -eq $value) {
                $value = ""
            } else {
                $value = $value.ToString()
            }

            # Apply value map (true/false -> ja/nein)
            if ($ValueMap -and $ValueMap.Count -gt 0 -and $value.Length -gt 0) {
                $lowerValue = $value.ToLower()
                if ($ValueMap.ContainsKey($lowerValue)) {
                    $value = $ValueMap[$lowerValue]
                }
            }

            # XML-escape the value
            $value = [System.Security.SecurityElement]::Escape($value)

            if ($value.Length -gt 255) {
                $value = $value.Substring(0, 252) + "..."
            }

            $xml = $xml.Replace($placeholder, $value)
        }

        # Write back
        $entry.Delete()
        $newEntry = $zip.CreateEntry("word/document.xml")
        $stream = $newEntry.Open()
        $writer = New-Object System.IO.StreamWriter($stream)
        $writer.Write($xml)
        $writer.Close()
        $stream.Close()

        return @{ Success = $true; ReplacedCount = $matches.Count }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
    finally {
        $zip.Dispose()
    }
}

function Process-SingleRoom {
    param(
        [object]$WordApp,
        [string]$TemplatePath,
        [PSCustomObject]$RoomData,
        [string]$OutputPath,
        [hashtable]$ValueMap,
        [array]$PlaceholderNames
    )

    $doc = $null
    $success = $false
    $replacedCount = 0
    $tempFile = $null

    try {
        $outputDir = Split-Path $OutputPath
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputPath)
        $tempFile = Join-Path $outputDir "_temp_$baseName.docx"

        $result = Process-RoomWithOpenXML `
            -TemplatePath $TemplatePath `
            -RoomData $RoomData `
            -OutputPath $tempFile `
            -ValueMap $ValueMap

        if (-not $result.Success) {
            Write-Host " Error: $($result.Error)" -ForegroundColor Red
            return @{ Success = $false; ReplacedCount = 0 }
        }

        $replacedCount = $result.ReplacedCount

        $doc = $WordApp.Documents.Open($tempFile, $false, $false)
        $doc.ExportAsFixedFormat($OutputPath, 17)
        $success = $true
    }
    catch {
        Write-Host " Error: $_" -ForegroundColor Red
    }
    finally {
        if ($null -ne $doc) {
            $doc.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        if ($tempFile -and (Test-Path $tempFile)) {
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }

    return @{ Success = $success; ReplacedCount = $replacedCount }
}

function Process-MergedRooms {
    param(
        [object]$WordApp,
        [string]$TemplatePath,
        [array]$AllRoomData,
        [string]$RoomCodeField,
        [string]$OutputPath,
        [hashtable]$ValueMap,
        [array]$PlaceholderNames
    )

    $successCount = 0
    $failCount = 0
    $tempFiles = @()

    try {
        $totalRooms = $AllRoomData.Count
        $i = 0
        $tempFolder = Split-Path $OutputPath

        # Step 1: Generate individual temp files using fast OpenXML replacement
        foreach ($room in $AllRoomData) {
            $i++
            $roomCode = $room.$RoomCodeField

            if ([string]::IsNullOrWhiteSpace($roomCode)) {
                $failCount++
                continue
            }

            Write-Host "[$i/$totalRooms] $roomCode..." -NoNewline

            try {
                $safeCode = $roomCode -replace '[\\/:*?"<>|]', '_'
                $tempFile = Join-Path $tempFolder "_temp_$safeCode.docx"

                $result = Process-RoomWithOpenXML `
                    -TemplatePath $TemplatePath `
                    -RoomData $room `
                    -OutputPath $tempFile `
                    -ValueMap $ValueMap

                if ($result.Success) {
                    $tempFiles += $tempFile
                    $successCount++
                    Write-Host " OK" -ForegroundColor Green
                } else {
                    $failCount++
                    Write-Host " FAILED: $($result.Error)" -ForegroundColor Red
                }
            }
            catch {
                $failCount++
                Write-Host " FAILED: $_" -ForegroundColor Red
            }
        }

        # Step 2: Merge all temp files into one document
        if ($tempFiles.Count -gt 0) {
            Write-Host ""
            Write-Host "Merging $($tempFiles.Count) documents..." -ForegroundColor Cyan

            $mergedDoc = $WordApp.Documents.Open($tempFiles[0], $false, $false)

            for ($j = 1; $j -lt $tempFiles.Count; $j++) {
                $range = $mergedDoc.Content
                $range.Collapse(0)  # wdCollapseEnd
                $range.InsertBreak(7)  # wdPageBreak
                $range.InsertFile($tempFiles[$j])
            }

            Write-Host "Exporting merged PDF..." -ForegroundColor Cyan
            $mergedDoc.ExportAsFixedFormat($OutputPath, 17)
            Write-Host "  Saved: $OutputPath" -ForegroundColor Green

            $mergedDoc.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mergedDoc) | Out-Null
        }
    }
    finally {
        # Cleanup temp files
        foreach ($tf in $tempFiles) {
            if (Test-Path $tf) {
                Remove-Item $tf -Force -ErrorAction SilentlyContinue
            }
        }
    }

    return @{ SuccessCount = $successCount; FailCount = $failCount }
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
    $config = Load-Configuration -Path $ConfigPath -SourceOverride $Source -TemplateOverride $Template -ExcelFileOverride $ExcelFile
    $dataSource = if ($config.DataSource) { $config.DataSource } else { 'Excel' }
    Write-Host "  Data source: $dataSource"
    if ($dataSource -eq 'API') {
        Write-Host "  API endpoint: (from file)"
    } else {
        Write-Host "  Data file: $($config.Excel.DataFile)"
    }
    Write-Host "  Template: $($config.TemplateFile)"
    Write-Host "  Output: $($config.OutputFolder)"
    Write-Host ""

    $dataSource = $config.DataSource
    if ($dataSource -eq 'API') {
        $roomData = Read-RoomDataFromAPI -APIConfig $config.API -ScriptDir $scriptDir
        $roomCodeField = $config.API.RoomCodeField
    } else {
        $roomData = Read-RoomData -Path $config.Excel.DataFile -RoomCodeColumn $config.Excel.RoomCodeColumn
        $roomCodeField = $config.Excel.RoomCodeColumn
    }
    Write-Host ""

    if (-not [string]::IsNullOrEmpty($RoomCode)) {
        $roomCodes = $RoomCode -split ',' | ForEach-Object { $_.Trim() }
        $roomData = $roomData | Where-Object { $_.$roomCodeField -in $roomCodes }
        if ($roomData.Count -eq 0) {
            throw "Room(s) not found: $RoomCode"
        }
        Write-Host "Filtering to $($roomCodes.Count) room(s): $($roomCodes -join ', ')" -ForegroundColor Cyan
        Write-Host ""
    }

    $startTotal = Get-Date

    Write-Host "Starting Word..." -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # Pre-scan template for placeholders (optimization: only replace what exists)
    Write-Host "Scanning template for placeholders..." -ForegroundColor Cyan
    $templatePlaceholders = Get-TemplatePlaceholders -WordApp $word -TemplatePath $config.TemplateFile
    if ($templatePlaceholders.Count -eq 0) {
        Write-Host "  WARNING: No <<placeholder>> markers found in template!" -ForegroundColor Yellow
        Write-Host "  Output PDFs will be identical copies of the template." -ForegroundColor Yellow
    } else {
        Write-Host "  Found $($templatePlaceholders.Count) placeholders" -ForegroundColor Green
    }

    $totalRooms = @($roomData).Count
    $successCount = 0
    $failCount = 0

    Write-Host ""
    if ($Merge) {
        Write-Host "Processing $totalRooms rooms (MERGED OUTPUT)..." -ForegroundColor Yellow
    } else {
        Write-Host "Processing $totalRooms rooms..." -ForegroundColor Yellow
    }
    Write-Host "-" * 60 -ForegroundColor Yellow

    $startProcess = Get-Date

    if ($Merge) {
        # Merged mode: combine all rooms into single PDF
        $mergedOutputPath = Join-Path $config.OutputFolder "AllRooms_Merged.pdf"

        $result = Process-MergedRooms `
            -WordApp $word `
            -TemplatePath $config.TemplateFile `
            -AllRoomData $roomData `
            -RoomCodeField $roomCodeField `
            -OutputPath $mergedOutputPath `
            -ValueMap $config.ValueMap `
            -PlaceholderNames $templatePlaceholders

        $successCount = $result.SuccessCount
        $failCount = $result.FailCount
    }
    else {
        # Individual mode: separate PDF per room
        $i = 0
        foreach ($room in $roomData) {
            $i++
            $roomCode = $room.$roomCodeField

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
                -OutputPath $outputPath `
                -ValueMap $config.ValueMap `
                -PlaceholderNames $templatePlaceholders

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
    }

    $processTime = (Get-Date) - $startProcess
    $totalTime = (Get-Date) - $startTotal

    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Yellow
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  Total rooms: $totalRooms"
    Write-Host "  Successful: $successCount" -ForegroundColor Green
    Write-Host "  Failed: $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
    if ($Merge) {
        Write-Host "  Output mode: MERGED (single PDF)"
    } else {
        Write-Host "  Output mode: Individual PDFs"
    }
    Write-Host ""
    Write-Host "Performance:" -ForegroundColor Yellow
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
