<#
.SYNOPSIS
    Sorts image files into year\month folder structure based on the most
    accurate available date, preferring EXIF DateTimeOriginal.

.DESCRIPTION
    Date source priority (best → fallback):
      1. EXIF DateTimeOriginal  — set by the camera at the moment of capture
      2. EXIF DateTimeDigitized — set when image was digitised
      3. File LastWriteTime     — guaranteed fallback when EXIF is absent or stripped
      4. Unresolved\            — only if the file system cannot provide any date

    Requires Shared-Functions.ps1 in the same folder.

.PARAMETER SourcePath
    Folder containing the image files to sort.

.PARAMETER OutputPath
    Root folder where the year\month structure will be created.

.PARAMETER Recurse
    Scan subfolders of SourcePath as well.

.PARAMETER Move
    Move files instead of copying them.

.PARAMETER FolderFormat
    Named  →  06 - June  (default)
    Number →  06
    Full   →  2024-06

.PARAMETER WhatIf
    Dry run — shows what would happen without touching any files.

.EXAMPLE
    .\Invoke-SortPhotos.ps1 -SourcePath "C:\Photos" -OutputPath "C:\Sorted"

.EXAMPLE
    .\Invoke-SortPhotos.ps1 -SourcePath "C:\Photos" -OutputPath "C:\Sorted" -Recurse -Move -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string] $SourcePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [switch] $Recurse,
    [switch] $Move,
    [ValidateSet('Named', 'Number', 'Full')]
    [string] $FolderFormat = 'Named'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Load shared functions
# ---------------------------------------------------------------------------

. "$PSScriptRoot\Shared-Functions.ps1"

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

$SupportedExtensions = @(
    '.jpg', '.jpeg', '.png', '.tiff', '.tif',
    '.heic', '.heif', '.bmp', '.gif', '.webp',
    '.cr2', '.cr3', '.nef', '.arw', '.orf',
    '.rw2', '.dng', '.raf'
)

# EXIF tag IDs: 36867 = DateTimeOriginal, 36868 = DateTimeDigitized
$ExifDateTags = @(36867, 36868)

# ---------------------------------------------------------------------------
# Photo-specific date helper
# ---------------------------------------------------------------------------

function Get-ExifDate {
    param ([string]$FilePath)
    try {
        Add-Type -AssemblyName System.Drawing
        $img = [System.Drawing.Image]::FromFile($FilePath)
        try {
            $propIds = $img.PropertyIdList
            foreach ($tagId in $ExifDateTags) {
                if ($tagId -notin $propIds) { continue }
                $raw = [System.Text.Encoding]::ASCII.GetString(
                    $img.GetPropertyItem($tagId).Value).TrimEnd([char]0)
                if ($raw -match '^\d{4}:\d{2}:\d{2} \d{2}:\d{2}:\d{2}$' -and
                    $raw -notmatch '^0000') {
                    return [datetime]::ParseExact($raw, 'yyyy:MM:dd HH:mm:ss',
                               [System.Globalization.CultureInfo]::InvariantCulture)
                }
            }
        }
        finally { $img.Dispose() }
    }
    catch { <# System.Drawing can't open some RAW formats — fall through #> }
    return $null
}

function Get-BestDate {
    param ([System.IO.FileInfo]$File)

    $exifDate = Get-ExifDate -FilePath $File.FullName
    if ($null -ne $exifDate) {
        return [PSCustomObject]@{ Date = $exifDate; Source = 'EXIF DateTimeOriginal' }
    }

    # Guaranteed fallback from Shared-Functions.ps1
    return Get-LastWriteTimeFallback -File $File
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

Assert-SourcePath -Path $SourcePath

$unresolvedFolder = Join-Path $OutputPath 'Unresolved'
$modeLabel = if ($Move) { 'MOVE  (source files will be removed)' } else { 'COPY  (source untouched)' }

Write-Host "`nMode          : $modeLabel"    -ForegroundColor Yellow
Write-Host "Folder format : $FolderFormat"  -ForegroundColor Yellow
Write-Host "Output root   : $OutputPath`n"  -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# Step 1 — Collect image files
# ---------------------------------------------------------------------------

Write-Host "[1/3] Scanning '$SourcePath'..." -ForegroundColor Cyan

$allFiles = Get-FilteredFiles -FolderPath         $SourcePath `
                              -Recurse:$Recurse `
                              -IncludeExtensions   $SupportedExtensions `
                              -EmptyMessage        "  No supported image files found in '$SourcePath'."

if ($allFiles.Count -eq 0) { exit 0 }
Write-Host "      Found $($allFiles.Count) supported image file(s)."

# ---------------------------------------------------------------------------
# Step 2 — Resolve best date per file
# ---------------------------------------------------------------------------

Write-Host "`n[2/3] Resolving dates..." -ForegroundColor Cyan

$resolved   = [System.Collections.Generic.List[PSCustomObject]]::new()
$unresolved = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
$total = $allFiles.Count
$index = 0

foreach ($file in $allFiles) {
    $index++
    Write-Progress -Activity 'Reading dates' `
                   -Status   "$index / $total : $($file.Name)" `
                   -PercentComplete (($index / $total) * 100)

    $bestDate = Get-BestDate -File $file
    if ($null -ne $bestDate) {
        $resolved.Add([PSCustomObject]@{ File = $file; Date = $bestDate.Date; DateSource = $bestDate.Source })
    }
    else { $unresolved.Add($file) }
}
Write-Progress -Activity 'Reading dates' -Completed

$exifCount     = @($resolved | Where-Object { $_.DateSource -eq 'EXIF DateTimeOriginal' }).Count
$fallbackCount = @($resolved | Where-Object { $_.DateSource -like '*fallback*' }).Count

Write-Host "      EXIF date found       : $exifCount file(s)"
Write-Host "      Fallback date used    : $fallbackCount file(s)"
Write-Host "      No date (Unresolved)  : $($unresolved.Count) file(s)"

# ---------------------------------------------------------------------------
# Step 3 — Sort into year\month folders
# ---------------------------------------------------------------------------

Write-Host "`n[3/3] Sorting files..." -ForegroundColor Cyan

$stats = @{ Copied = 0; Moved = 0; Skipped = 0; Unresolved = 0; Errors = 0 }
$index = 0
$total = $resolved.Count

foreach ($item in $resolved) {
    $index++
    Write-Progress -Activity 'Sorting' `
                   -Status   "$index / $total : $($item.File.Name)" `
                   -PercentComplete (($index / $total) * 100)
    try {
        $destFolder = Join-Path (Join-Path $OutputPath $item.Date.Year.ToString()) `
                                (Get-MonthFolderName -Date $item.Date -Format $FolderFormat)

        $result = Invoke-FileTransfer -FilePath     $item.File.FullName `
                                      -DestFolder   $destFolder `
                                      -DeleteSource $Move.IsPresent

        switch -Wildcard ($result.Action) {
            'Moved'      { $stats.Moved++ }
            'Copied'     { $stats.Copied++ }
            'Skipped*'   { $stats.Skipped++ }
        }
        Write-Verbose "$($result.Action): '$($item.File.Name)' [$($item.DateSource)] → $($result.Path)"
    }
    catch {
        Write-Warning "Failed to process '$($item.File.FullName)': $_"
        $stats.Errors++
    }
}

foreach ($file in $unresolved) {
    try {
        Invoke-FileTransfer -FilePath $file.FullName -DestFolder $unresolvedFolder `
                            -DeleteSource $Move.IsPresent | Out-Null
        $stats.Unresolved++
    }
    catch {
        Write-Warning "Failed to copy unresolved '$($file.FullName)': $_"
        $stats.Errors++
    }
}
Write-Progress -Activity 'Sorting' -Completed

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  Photo Sort Complete"                     -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode              : $modeLabel"
Write-Host "  Images found      : $($allFiles.Count)"
Write-Host "  EXIF dates        : $exifCount"
Write-Host "  Fallback dates    : $fallbackCount"
if ($Move) { Write-Host "  Moved   : $($stats.Moved)" }
else       { Write-Host "  Copied  : $($stats.Copied)" }
if ($stats.Skipped    -gt 0) { Write-Host "  Skipped (identical)  : $($stats.Skipped)"              -ForegroundColor Yellow }
if ($stats.Unresolved -gt 0) { Write-Host "  Unresolved           : $($stats.Unresolved)  →  $unresolvedFolder" -ForegroundColor Yellow }
if ($stats.Errors     -gt 0) { Write-Host "  Errors               : $($stats.Errors)"               -ForegroundColor Red }
Write-Host "========================================`n" -ForegroundColor Green
