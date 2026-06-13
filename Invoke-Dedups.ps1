<#
.SYNOPSIS
    Deduplicates files using a tiered approach:
    1. Group by file size + extension (fast metadata filter)
    2. Partial hash of first/last 64KB (cheap secondary filter)
    3. Full SHA256 hash (confirmed duplicate detection)

.DESCRIPTION
    Scans a source folder and sorts files into:
      originals\  : unique files or the first-seen copy of a duplicate
      copies\     : confirmed duplicate files

    Files are NEVER deleted. They are COPIED into the output folders.
    Your source files are left completely untouched.

.PARAMETER SourcePath
    Path to the folder containing files to deduplicate.
    Use -Recurse to include subfolders.

.PARAMETER OutputPath
    Path to the output folder. 'originals' and 'copies' subfolders
    will be created here automatically.

.PARAMETER Recurse
    If specified, scans subfolders of SourcePath as well.

.PARAMETER WhatIf
    Dry run : reports what would happen without copying any files.

.EXAMPLE
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted"

.EXAMPLE
    .\Invoke-Dedup.ps1 -SourcePath "C:\MyFiles" -OutputPath "C:\Sorted" -Recurse -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string] $SourcePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [switch] $Recurse
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-PartialHash {
    <#
    Reads the first and last 64 KB of a file and returns their combined MD5.
    Falls back to full content for files smaller than 128 KB.
    #>
    param ([string]$FilePath)

    $chunkSize = 64KB
    $bytes = [System.Collections.Generic.List[byte]]::new()

    $stream = [System.IO.File]::OpenRead($FilePath)
    try {
        $fileLength = $stream.Length

        if ($fileLength -le ($chunkSize * 2)) {
            # Small file — read everything
            $buf = [byte[]]::new($fileLength)
            $null = $stream.Read($buf, 0, $fileLength)
            $bytes.AddRange($buf)
        }
        else {
            # Read first 64 KB
            $buf = [byte[]]::new($chunkSize)
            $null = $stream.Read($buf, 0, $chunkSize)
            $bytes.AddRange($buf)

            # Read last 64 KB
            $null = $stream.Seek(-$chunkSize, [System.IO.SeekOrigin]::End)
            $buf = [byte[]]::new($chunkSize)
            $null = $stream.Read($buf, 0, $chunkSize)
            $bytes.AddRange($buf)
        }
    }
    finally {
        $stream.Dispose()
    }

    $md5 = [System.Security.Cryptography.MD5]::Create()
    try {
        $hash = $md5.ComputeHash($bytes.ToArray())
        return [System.BitConverter]::ToString($hash) -replace '-', ''
    }
    finally {
        $md5.Dispose()
    }
}

function Get-FullHash {
    param ([string]$FilePath)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [System.IO.File]::OpenRead($FilePath)
        try {
            $hash = $sha.ComputeHash($stream)
            return [System.BitConverter]::ToString($hash) -replace '-', ''
        }
        finally { $stream.Dispose() }
    }
    finally { $sha.Dispose() }
}

function Copy-ToOutput {
    param (
        [string]$FilePath,
        [string]$DestFolder
    )

    $fileName  = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $destPath  = Join-Path $DestFolder "$fileName$extension"

    # Avoid filename collisions in the destination folder
    $counter = 1
    while (Test-Path $destPath) {
        $destPath = Join-Path $DestFolder "${fileName}_$counter$extension"
        $counter++
    }

    if ($PSCmdlet.ShouldProcess($destPath, "Copy '$FilePath'")) {
        Copy-Item -LiteralPath $FilePath -Destination $destPath -Force
    }

    return $destPath
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

$originalsFolder = Join-Path $OutputPath 'originals'
$copiesFolder    = Join-Path $OutputPath 'copies'

if (-not (Test-Path $originalsFolder)) { New-Item -ItemType Directory -Path $originalsFolder | Out-Null }
if (-not (Test-Path $copiesFolder))    { New-Item -ItemType Directory -Path $copiesFolder    | Out-Null }

# ---------------------------------------------------------------------------
# Step 1 — Collect files
# ---------------------------------------------------------------------------

Write-Host "`n[1/4] Scanning '$SourcePath'..." -ForegroundColor Cyan

$getChildParams = @{ LiteralPath = $SourcePath; File = $true }
if ($Recurse) { $getChildParams['Recurse'] = $true }

$allFiles = Get-ChildItem @getChildParams

Write-Host "      Found $($allFiles.Count) file(s)."

# ---------------------------------------------------------------------------
# Step 2 — Group by size + extension (metadata pre-filter)
# ---------------------------------------------------------------------------

Write-Host "`n[2/4] Grouping by size + extension..." -ForegroundColor Cyan

$groups = $allFiles | Group-Object -Property {
    "$($_.Length)|$($_.Extension.ToLower())"
}

$singletons     = $groups | Where-Object { $_.Count -eq 1 }
$candidateGroups = $groups | Where-Object { $_.Count -gt 1 }

$singletonCount  = ($singletons  | Measure-Object -Property Count -Sum).Sum
$candidateCount  = ($candidateGroups | Measure-Object -Property Count -Sum).Sum

Write-Host "      Singletons (unique size+ext): $singletonCount file(s) → originals/"
Write-Host "      Candidates for hashing:       $candidateCount file(s)"

# Copy singletons straight to originals
foreach ($group in $singletons) {
    $file = $group.Group[0]
    Copy-ToOutput -FilePath $file.FullName -DestFolder $originalsFolder | Out-Null
}

# ---------------------------------------------------------------------------
# Step 3 & 4 — Partial hash, then full SHA256 on matches
# ---------------------------------------------------------------------------

Write-Host "`n[3/4] Running partial hash on candidates..." -ForegroundColor Cyan

# Track confirmed duplicates: SHA256 → first-seen file path
$seenHashes = @{}

$stats = @{ Originals = $singletonCount; Copies = 0; Errors = 0 }

$candidateFiles = $candidateGroups | ForEach-Object { $_.Group }
$total          = @($candidateFiles).Count
$index          = 0

# Group candidates by partial hash first
$partialHashGroups = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[string]]]::new()

foreach ($file in $candidateFiles) {
    $index++
    Write-Progress -Activity "Partial hashing" -Status "$index / $total : $($file.Name)" `
                   -PercentComplete (($index / $total) * 100)

    try {
        $ph = Get-PartialHash -FilePath $file.FullName

        if (-not $partialHashGroups.ContainsKey($ph)) {
            $partialHashGroups[$ph] = [System.Collections.Generic.List[string]]::new()
        }
        $partialHashGroups[$ph].Add($file.FullName)
    }
    catch {
        Write-Warning "Could not hash '$($file.FullName)': $_"
        $stats.Errors++
    }
}

Write-Progress -Activity "Partial hashing" -Completed

Write-Host "`n[4/4] Running full SHA256 on partial-hash matches..." -ForegroundColor Cyan

$phGroups      = $partialHashGroups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
$phSingletons  = $partialHashGroups.GetEnumerator() | Where-Object { $_.Value.Count -eq 1 }

# Partial-hash singletons are unique — send to originals
foreach ($entry in $phSingletons) {
    Copy-ToOutput -FilePath $entry.Value[0] -DestFolder $originalsFolder | Out-Null
    $stats.Originals++
}

# Full SHA256 on remaining groups
$index = 0
$phGroupList = @($phGroups)
$total = ($phGroupList | ForEach-Object { $_.Value.Count } | Measure-Object -Sum).Sum

foreach ($entry in $phGroupList) {
    foreach ($filePath in $entry.Value) {
        $index++
        Write-Progress -Activity "Full SHA256" -Status "$index / $total : $(Split-Path $filePath -Leaf)" `
                       -PercentComplete (($index / $total) * 100)

        try {
            $fullHash = Get-FullHash -FilePath $filePath

            if ($seenHashes.ContainsKey($fullHash)) {
                # Confirmed duplicate
                Copy-ToOutput -FilePath $filePath -DestFolder $copiesFolder | Out-Null
                $stats.Copies++
            }
            else {
                $seenHashes[$fullHash] = $filePath
                Copy-ToOutput -FilePath $filePath -DestFolder $originalsFolder | Out-Null
                $stats.Originals++
            }
        }
        catch {
            Write-Warning "Could not hash '$filePath': $_"
            $stats.Errors++
        }
    }
}

Write-Progress -Activity "Full SHA256" -Completed

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  Deduplication Complete" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Total files scanned : $($allFiles.Count)"
Write-Host "  Originals            : $($stats.Originals)  →  $originalsFolder"
Write-Host "  Copies (duplicates)  : $($stats.Copies)  →  $copiesFolder"
if ($stats.Errors -gt 0) {
    Write-Host "  Errors               : $($stats.Errors)" -ForegroundColor Yellow
}
Write-Host "========================================`n" -ForegroundColor Green
