<#
.SYNOPSIS
    Sorts video files into year\month folder structure based on the most
    accurate available date, using Windows built-in Shell property APIs.
    No third-party tools required.

.DESCRIPTION
    Date source priority (best → fallback):
      1. Shell: System.Media.DateEncoded    — recording date embedded by camera/device
      2. Shell: System.Document.DateCreated — document creation date in container metadata
      3. Shell: System.Media.DateReleased   — release/publish date (last resort media tag)
      4. File LastWriteTime                 — GUARANTEED fallback when all metadata absent
      5. Unresolved\                        — only if the file itself is unreadable

    All metadata is read via the Windows Shell COM API (the same engine that
    powers File Explorer's Details pane). No third-party tools are required.

    Output structure:
      OutputPath\
        2022\
          01 - January\
          06 - June\
        2023\
          11 - November\
        Unresolved\   ← files where no date could be determined at all

    Supported formats:
      MP4, MOV, M4V, AVI, MKV, WMV, FLV, WEBM, MPG, MPEG,
      3GP, 3G2, MTS, M2TS, TS, VOB, ASF, DV

    Note: Windows Shell metadata support varies by format. MP4, MOV, AVI, WMV,
    and MTS tend to have the best coverage. Obscure or broadcast formats (MXF,
    RM, RMVB) may fall through to the LastWriteTime fallback. If you need
    maximum metadata coverage across all formats, FFprobe is the better choice.

.PARAMETER SourcePath
    Folder containing the video files to sort.

.PARAMETER OutputPath
    Root folder where the year\month structure will be created.

.PARAMETER Recurse
    Scan subfolders of SourcePath as well.

.PARAMETER Move
    Move files instead of copying them. Source file is removed after a
    successful transfer.

.PARAMETER FolderFormat
    Controls the month subfolder naming style.
      Named    →  06 - June          (default)
      Number   →  06
      Full     →  2024-06

.PARAMETER WhatIf
    Dry run — shows what would happen without moving or copying anything.

.EXAMPLE
    # Copy videos, named month folders
    .\Invoke-SortVideos.ps1 -SourcePath "C:\Videos" -OutputPath "C:\Sorted"

.EXAMPLE
    # Move videos, recurse subfolders, numeric month folders
    .\Invoke-SortVideos.ps1 -SourcePath "C:\Videos" -OutputPath "C:\Sorted" -Recurse -Move -FolderFormat Number

.EXAMPLE
    # Dry run first to preview results before committing
    .\Invoke-SortVideos.ps1 -SourcePath "C:\Videos" -OutputPath "C:\Sorted" -Recurse -Move -WhatIf
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
# Constants
# ---------------------------------------------------------------------------

$SupportedExtensions = @(
    '.mp4', '.mov', '.m4v', '.avi', '.mkv',
    '.wmv', '.flv', '.webm', '.mpg', '.mpeg',
    '.3gp', '.3g2', '.mts', '.m2ts', '.ts',
    '.vob', '.asf', '.dv'
)

$MonthNames = @{
    1='January'; 2='February'; 3='March';     4='April'
    5='May';     6='June';     7='July';       8='August'
    9='September'; 10='October'; 11='November'; 12='December'
}

# Windows Shell property keys for video recording dates, in priority order.
# These are the canonical property System IDs used by the Windows Shell API.
# The same values you see in File Explorer → right-click → Properties → Details.
#
#   System.Media.DateEncoded    — date the media was encoded/recorded (most accurate)
#   System.Document.DateCreated — date embedded in the file container
#   System.Media.DateReleased   — release date tag (last resort; often year only)
#
$ShellDateProperties = @(
    'System.Media.DateEncoded',
    'System.Document.DateCreated',
    'System.Media.DateReleased'
)

# ---------------------------------------------------------------------------
# Shell COM helper — initialised once and reused for performance
# ---------------------------------------------------------------------------

# The Windows Shell Folder object lets us read extended file properties
# (the same ones shown in File Explorer Details) without any extra installs.
$ShellApp = New-Object -ComObject Shell.Application

function Get-ShellDate {
    <#
    Reads video metadata date properties from a file using the Windows Shell
    COM API. Returns a [datetime] or $null if no usable date is found.

    The Shell API queries the same property handlers that File Explorer uses,
    so any format Windows can show details for will work here too.
    #>
    param ([string]$FilePath)

    try {
        $folder   = $ShellApp.NameSpace([System.IO.Path]::GetDirectoryName($FilePath))
        $fileItem = $folder.ParseName([System.IO.Path]::GetFileName($FilePath))

        if ($null -eq $fileItem) { return $null }

        foreach ($propName in $ShellDateProperties) {
            try {
                # GetDetailsOf with index -1 lets us query by property name string
                # We use the property system's canonical name for reliability
                $propValue = $folder.GetDetailsOf($fileItem,
                    # Map property name to column index dynamically
                    ($folder.Items() | Select-Object -First 1 |
                     ForEach-Object { 0..500 | Where-Object {
                         $folder.GetDetailsOf($folder.Items().Item($_), -1) -eq $propName
                     } } | Select-Object -First 1)
                )

                if ($propValue -and $propValue.Trim() -ne '') {
                    $parsed = $null
                    if ([datetime]::TryParse($propValue.Trim(),
                            [System.Globalization.CultureInfo]::CurrentCulture,
                            [System.Globalization.DateTimeStyles]::None,
                            [ref]$parsed)) {
                        # Sanity check — ignore epoch zeros and far-future dates
                        if ($parsed.Year -gt 1970 -and $parsed.Year -le ([datetime]::Now.Year + 1)) {
                            return $parsed
                        }
                    }
                }
            }
            catch { <# Property not supported by this file's handler — try next #> }
        }
    }
    catch { <# Shell couldn't open the file — fall through to LastWriteTime #> }

    return $null
}

function Get-ShellDateByIndex {
    <#
    More reliable alternative: queries Shell properties by their well-known
    column indices rather than by name string matching.
    Windows Explorer column indices for the Details pane:

      Column 208 = Media created (System.Media.DateEncoded) — most accurate for video
      Column 4   = Date modified  (fallback within Shell)

    This approach is faster and more compatible across Windows versions.
    #>
    param ([string]$FilePath)

    try {
        $folder   = $ShellApp.NameSpace([System.IO.Path]::GetDirectoryName($FilePath))
        $fileItem = $folder.ParseName([System.IO.Path]::GetFileName($FilePath))

        if ($null -eq $fileItem) { return $null }

        # Try known column indices for media date properties
        # 208 = Media created, 191 = Date encoded, 197 = Date released
        $indicesToTry = @(208, 191, 197)

        foreach ($idx in $indicesToTry) {
            try {
                $val = $folder.GetDetailsOf($fileItem, $idx)
                if ($val -and $val.Trim() -ne '') {
                    # Shell returns date strings with hidden Unicode chars — strip them
                    $cleaned = $val -replace '[^\x20-\x7E]', '' | ForEach-Object { $_.Trim() }
                    $parsed  = $null
                    if ($cleaned -and [datetime]::TryParse($cleaned,
                            [System.Globalization.CultureInfo]::CurrentCulture,
                            [System.Globalization.DateTimeStyles]::None,
                            [ref]$parsed)) {
                        if ($parsed.Year -gt 1970 -and $parsed.Year -le ([datetime]::Now.Year + 1)) {
                            return [PSCustomObject]@{ Date = $parsed; ColumnIndex = $idx }
                        }
                    }
                }
            }
            catch { <# Column not available for this file type #> }
        }
    }
    catch { <# Shell couldn't open file #> }

    return $null
}

function Get-BestDate {
    <#
    Returns the best available date for a video file and a label describing
    which source was used.

    Priority:
      1. Windows Shell media metadata  — embedded recording date (no install needed)
      2. File LastWriteTime            — GUARANTEED fallback; every readable file has one.
         CreationTime is NOT used — Windows resets it on every file copy.
    #>
    param ([System.IO.FileInfo]$File)

    # ── Priority 1: Windows Shell metadata ───────────────────────────────
    $shellResult = Get-ShellDateByIndex -FilePath $File.FullName
    if ($null -ne $shellResult) {
        $sourceLabel = switch ($shellResult.ColumnIndex) {
            208 { 'Shell: Media created (DateEncoded)' }
            191 { 'Shell: Date encoded' }
            197 { 'Shell: Date released' }
            default { "Shell: column $($shellResult.ColumnIndex)" }
        }
        return [PSCustomObject]@{ Date = $shellResult.Date; Source = $sourceLabel }
    }

    # ── Priority 2: File LastWriteTime (guaranteed fallback) ─────────────
    # Reached when Windows Shell finds no embedded media date.
    # Common cases: AVI with no metadata, stripped MP4, unsupported container.
    # No year-range gate — any valid timestamp is accepted.
    $lwt = $File.LastWriteTime
    if ($null -ne $lwt -and $lwt -ne [datetime]::MinValue) {
        return [PSCustomObject]@{ Date = $lwt; Source = 'LastWriteTime (metadata absent — fallback)' }
    }

    # ── Unresolvable ──────────────────────────────────────────────────────
    # In practice never reached on a readable Windows file.
    return $null
}

function Get-MonthFolderName {
    param ([datetime]$Date)

    switch ($FolderFormat) {
        'Number' { return '{0:D2}' -f $Date.Month }
        'Full'   { return '{0}-{1:D2}' -f $Date.Year, $Date.Month }
        default  { return '{0:D2} - {1}' -f $Date.Month, $MonthNames[$Date.Month] }
    }
}

function Copy-ToDestination {
    param (
        [string] $FilePath,
        [string] $DestFolder,
        [bool]   $DeleteSource
    )

    if (-not (Test-Path $DestFolder)) {
        if ($PSCmdlet.ShouldProcess($DestFolder, 'Create directory')) {
            New-Item -ItemType Directory -Path $DestFolder -Force | Out-Null
        }
    }

    $fileName  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $destPath  = Join-Path $DestFolder "$fileName$extension"

    # Collision handling — skip if same name + size (already processed),
    # rename with incrementing suffix otherwise
    $counter = 1
    while (Test-Path $destPath) {
        $existing = Get-Item $destPath
        $incoming = Get-Item $FilePath
        if ($existing.Length -eq $incoming.Length) {
            return [PSCustomObject]@{ Path = $destPath; Action = 'Skipped (already exists)' }
        }
        $destPath = Join-Path $DestFolder "${fileName}_$counter$extension"
        $counter++
    }

    if ($PSCmdlet.ShouldProcess($destPath, "$(if ($DeleteSource) {'Move'} else {'Copy'}) '$FilePath'")) {
        Copy-Item -LiteralPath $FilePath -Destination $destPath -Force
        if ($DeleteSource) {
            Remove-Item -LiteralPath $FilePath -Force
        }
    }

    return [PSCustomObject]@{ Path = $destPath; Action = if ($DeleteSource) { 'Moved' } else { 'Copied' } }
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

$unresolvedFolder = Join-Path $OutputPath 'Unresolved'

$modeLabel = if ($Move) { 'MOVE  (source files will be removed)' } else { 'COPY  (source untouched)' }
Write-Host "`nMode          : $modeLabel"   -ForegroundColor Yellow
Write-Host "Folder format : $FolderFormat" -ForegroundColor Yellow
Write-Host "Output root   : $OutputPath`n" -ForegroundColor Yellow
Write-Host "Metadata via  : Windows Shell API (no third-party tools needed)" -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# Step 1 — Collect video files
# ---------------------------------------------------------------------------

Write-Host "`n[1/3] Scanning '$SourcePath'..." -ForegroundColor Cyan

$getChildParams = @{ LiteralPath = $SourcePath; File = $true }
if ($Recurse) { $getChildParams['Recurse'] = $true }

# @() wraps the result in an array — prevents .Count failing when the
# folder is empty or contains no matching files (PowerShell returns $null
# from a pipeline with no results, and $null has no .Count property)
$allFiles = @(Get-ChildItem @getChildParams |
    Where-Object { $_.Extension.ToLower() -in $SupportedExtensions })

Write-Host "      Found $($allFiles.Count) supported video file(s)."

if ($allFiles.Count -eq 0) {
    Write-Host "`n  No supported video files found in '$SourcePath'." -ForegroundColor Yellow
    Write-Host "  Supported extensions: $($SupportedExtensions -join ', ')`n" -ForegroundColor Yellow
    exit 0
}

# ---------------------------------------------------------------------------
# Step 2 — Resolve best date for each file
# ---------------------------------------------------------------------------

Write-Host "`n[2/3] Resolving dates (Windows Shell metadata)..." -ForegroundColor Cyan

$resolved   = [System.Collections.Generic.List[PSCustomObject]]::new()
$unresolved = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

$total = $allFiles.Count
$index = 0

foreach ($file in $allFiles) {
    $index++
    Write-Progress -Activity 'Reading metadata' `
                   -Status   "$index / $total : $($file.Name)" `
                   -PercentComplete (($index / $total) * 100)

    $bestDate = Get-BestDate -File $file

    if ($null -ne $bestDate) {
        $resolved.Add([PSCustomObject]@{
            File       = $file
            Date       = $bestDate.Date
            DateSource = $bestDate.Source
        })
    }
    else {
        $unresolved.Add($file)
    }
}

Write-Progress -Activity 'Reading metadata' -Completed

$shellCount    = @($resolved | Where-Object { $_.DateSource -like 'Shell:*' }).Count
$fallbackCount = @($resolved | Where-Object { $_.DateSource -like '*fallback*' }).Count

Write-Host "      Shell metadata found   : $shellCount file(s)"
Write-Host "      Fallback date used     : $fallbackCount file(s)"
Write-Host "      No date (Unresolved)   : $($unresolved.Count) file(s)"

# ---------------------------------------------------------------------------
# Step 3 — Copy / Move files into year\month structure
# ---------------------------------------------------------------------------

Write-Host "`n[3/3] Sorting files..." -ForegroundColor Cyan

$stats = @{
    Copied     = 0
    Moved      = 0
    Skipped    = 0
    Unresolved = 0
    Errors     = 0
}

$index = 0
$total = $resolved.Count

foreach ($item in $resolved) {
    $index++
    Write-Progress -Activity 'Sorting' `
                   -Status   "$index / $total : $($item.File.Name)" `
                   -PercentComplete (($index / $total) * 100)

    try {
        $yearFolder  = Join-Path $OutputPath $item.Date.Year.ToString()
        $monthFolder = Join-Path $yearFolder (Get-MonthFolderName -Date $item.Date)

        $result = Copy-ToDestination -FilePath     $item.File.FullName `
                                     -DestFolder   $monthFolder `
                                     -DeleteSource $Move.IsPresent

        switch ($result.Action) {
            'Moved'                    { $stats.Moved++ }
            'Copied'                   { $stats.Copied++ }
            'Skipped (already exists)' { $stats.Skipped++ }
        }

        Write-Verbose "$($result.Action): '$($item.File.Name)' [$($item.DateSource)] → $($result.Path)"
    }
    catch {
        Write-Warning "Failed to process '$($item.File.FullName)': $_"
        $stats.Errors++
    }
}

# Unresolved files → Unresolved\ folder
foreach ($file in $unresolved) {
    try {
        Copy-ToDestination -FilePath     $file.FullName `
                           -DestFolder   $unresolvedFolder `
                           -DeleteSource $Move.IsPresent | Out-Null
        $stats.Unresolved++
    }
    catch {
        Write-Warning "Failed to copy unresolved file '$($file.FullName)': $_"
        $stats.Errors++
    }
}

Write-Progress -Activity 'Sorting' -Completed

# ---------------------------------------------------------------------------
# Cleanup COM object
# ---------------------------------------------------------------------------

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ShellApp) | Out-Null

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  Video Sort Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode              : $modeLabel"
Write-Host "  Videos found      : $($allFiles.Count)"
Write-Host "  Shell metadata    : $shellCount"
Write-Host "  Fallback dates    : $fallbackCount"
if ($Move) {
    Write-Host "  Moved             : $($stats.Moved)"
} else {
    Write-Host "  Copied            : $($stats.Copied)"
}
if ($stats.Skipped -gt 0) {
    Write-Host "  Skipped (dupes)   : $($stats.Skipped)"  -ForegroundColor Yellow
}
if ($stats.Unresolved -gt 0) {
    Write-Host "  Unresolved        : $($stats.Unresolved)  →  $unresolvedFolder" -ForegroundColor Yellow
}
if ($stats.Errors -gt 0) {
    Write-Host "  Errors            : $($stats.Errors)"   -ForegroundColor Red
}
Write-Host "========================================`n" -ForegroundColor Green
