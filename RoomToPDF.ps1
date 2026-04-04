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
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = "",
    [string]$RoomCode = "",
    [ValidateSet('Excel', 'API')]
    [string]$Source = "",
    [string]$Template = "",
    [string]$ExcelFile = ""
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

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

function Process-SingleRoom {
    param(
        [object]$WordApp,
        [string]$TemplatePath,
        [PSCustomObject]$RoomData,
        [string]$OutputPath,
        [hashtable]$ValueMap
    )

    $doc = $null
    $success = $false
    $replacedCount = 0

    try {
        $doc = $WordApp.Documents.Open($TemplatePath, $false, $false)

        foreach ($prop in $RoomData.PSObject.Properties) {
            $placeholder = "<<$($prop.Name)>>"
            $value = if ($null -ne $prop.Value) { $prop.Value.ToString() } else { "" }

            if ([string]::IsNullOrWhiteSpace($value)) {
                $value = ""
            }

            if ($ValueMap -and $ValueMap.Count -gt 0 -and $value.Length -gt 0) {
                $lowerValue = $value.ToLower()
                if ($ValueMap.ContainsKey($lowerValue)) {
                    $value = $ValueMap[$lowerValue]
                }
            }

            if ($value.Length -gt 255) {
                $value = $value.Substring(0, 252) + "..."
            }

            $find = $doc.Content.Find
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()

            if ($find.Execute($placeholder, $false, $false, $false, $false, $false, $true, 1, $false, $value, 2)) {
                $replacedCount++
            }
        }

        $doc.ExportAsFixedFormat($OutputPath, 17)
        $success = $true
    }
    catch {
        Write-Host " Error: $_" -ForegroundColor Red
    }
    finally {
        if ($null -ne $doc) {
            $doc.Close([ref]$false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
    }

    return @{ Success = $success; ReplacedCount = $replacedCount }
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
            -ValueMap $config.ValueMap

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
