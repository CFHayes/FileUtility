<#
.SYNOPSIS
    Sorts document files into flat file-type folders.
    Only processes extensions explicitly listed in $CategoryMap —
    the inclusion list makes a separate exclusion list unnecessary.

.DESCRIPTION
    Files are organised by type category only — no date subfolders:

      OutputPath\
        PDF\
        Word\
        Excel\
        PowerPoint\
        OneNote\
        Access\
        Visio\
        Project\
        Text\
        CSV\
        eBooks\
        Markup\
        Unclassified\   ← any file whose extension is not in $CategoryMap

    File type groupings (mixed granularity):

      PDF          : .pdf
      Word         : .doc .docx .docm .dot .dotx .dotm .odt .rtf .wps
      Excel        : .xls .xlsx .xlsm .xlsb .xlt .xltx .xltm .ods .numbers
      PowerPoint   : .ppt .pptx .pptm .pot .potx .potm .odp .key
      OneNote      : .one .onetoc2
      Access       : .mdb .accdb .accde .accdt
      Visio        : .vsd .vsdx .vsdm .vss .vst
      Project      : .mpp .mpt
      Text         : .txt .log .nfo .asc .me
      CSV          : .csv .tsv .tab
      eBooks       : .epub .mobi .azw .azw3 .fb2 .lit .pdb .djvu
      Markup       : .html .htm .xml .json .yaml .yml .md .markdown .rst .tex

    Only files whose extension appears in $CategoryMap are collected.
    All other extensions — photos, videos, archives, system files, etc. —
    are naturally excluded because they are simply not in the inclusion list.

    Requires Shared-Functions.ps1 in the same folder.

.PARAMETER SourcePath
    Folder containing the document files to sort.

.PARAMETER OutputPath
    Root folder where the type\year\month structure will be created.

.PARAMETER Recurse
    Scan subfolders of SourcePath as well.

.PARAMETER Move
    Move files instead of copying them. Source file is removed after
    a successful transfer.

.PARAMETER WhatIf
    Dry run — shows what would happen without touching any files.

.EXAMPLE
    .\Invoke-SortDocuments.ps1 -SourcePath "C:\Documents" -OutputPath "C:\Sorted"

.EXAMPLE
    .\Invoke-SortDocuments.ps1 -SourcePath "C:\Documents" -OutputPath "C:\Sorted" -Recurse -Move

.EXAMPLE
    .\Invoke-SortDocuments.ps1 -SourcePath "C:\Documents" -OutputPath "C:\Sorted" -Recurse -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string] $SourcePath,
    [Parameter(Mandatory)][string] $OutputPath,
    [switch] $Recurse,
    [switch] $Move
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Load shared functions
# ---------------------------------------------------------------------------

. "$PSScriptRoot\Shared-Functions.ps1"

# ---------------------------------------------------------------------------
# File type category definitions
# Mixed granularity: specific for Office apps, broad for everything else
# ---------------------------------------------------------------------------

$CategoryMap = [ordered]@{

    # ── Specific Office categories ────────────────────────────────────────
    'PDF'         = @('.pdf')
    'Word'        = @('.doc','.docx','.docm','.dot','.dotx','.dotm','.odt','.rtf','.wps')
    'Excel'       = @('.xls','.xlsx','.xlsm','.xlsb','.xlt','.xltx','.xltm','.ods','.numbers')
    'PowerPoint'  = @('.ppt','.pptx','.pptm','.pot','.potx','.potm','.odp','.key')
    'OneNote'     = @('.one','.onetoc2')
    'Access'      = @('.mdb','.accdb','.accde','.accdt')
    'Visio'       = @('.vsd','.vsdx','.vsdm','.vss','.vst')
    'Project'     = @('.mpp','.mpt')

    # ── Broad common categories ───────────────────────────────────────────
    'Text'        = @('.txt','.log','.nfo','.asc','.me')
    'CSV'         = @('.csv','.tsv','.tab')
    'eBooks'      = @('.epub','.mobi','.azw','.azw3','.fb2','.lit','.pdb','.djvu')
    'Markup'      = @('.html','.htm','.xml','.json','.yaml','.yml','.md','.markdown','.rst','.tex')
}

# ---------------------------------------------------------------------------
# Build flat extension → category lookup for fast resolution
# The inclusion list ($SupportedExtensions) is the only filter needed —
# files not in $CategoryMap are never collected by Get-FilteredFiles.
# ---------------------------------------------------------------------------

$ExtCategoryLookup = @{}
foreach ($category in $CategoryMap.Keys) {
    foreach ($ext in $CategoryMap[$category]) {
        $ExtCategoryLookup[$ext] = $category
    }
}

# All supported extensions = everything in CategoryMap
$SupportedExtensions = @($ExtCategoryLookup.Keys)


# ---------------------------------------------------------------------------
# Category resolution
# ---------------------------------------------------------------------------

function Get-FileCategory {
    <#
    Returns the category folder name for a given file extension.
    Returns $null if the extension is not in the supported list.
    #>
    param ([string]$Extension)
    $ext = $Extension.ToLower()
    if ($ExtCategoryLookup.ContainsKey($ext)) {
        return $ExtCategoryLookup[$ext]
    }
    return $null
}

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

Assert-SourcePath -Path $SourcePath

$unclassifiedFolder = Join-Path $OutputPath 'Unclassified'
$modeLabel = if ($Move) { 'MOVE  (source files will be removed)' } else { 'COPY  (source untouched)' }

Write-Host "`nMode          : $modeLabel"    -ForegroundColor Yellow
Write-Host "Output root   : $OutputPath"    -ForegroundColor Yellow
Write-Host "Structure     : Type\ (flat — no date subfolders)" -ForegroundColor Yellow


# ---------------------------------------------------------------------------
# Step 1 — Collect files
# Includes: supported document extensions
# Excludes: all extensions owned by other suite scripts
# ---------------------------------------------------------------------------

Write-Host "[1/3] Scanning '$SourcePath'..." -ForegroundColor Cyan

$allFiles = Get-FilteredFiles `
    -FolderPath        $SourcePath `
    -Recurse:$Recurse `
    -IncludeExtensions $SupportedExtensions `
    -EmptyMessage      "  No supported document files found in '$SourcePath'."

if ($allFiles.Count -eq 0) { exit 0 }

# Per-category counts for scan summary
Write-Host "      Found $($allFiles.Count) document file(s):"
foreach ($cat in $CategoryMap.Keys) {
    $catExts  = $CategoryMap[$cat]
    $catCount = @($allFiles | Where-Object { $_.Extension.ToLower() -in $catExts }).Count
    if ($catCount -gt 0) {
        Write-Host ("        {0,-14}: {1}" -f $cat, $catCount)
    }
}

# ---------------------------------------------------------------------------
# Step 2 — Resolve best date and category per file
# ---------------------------------------------------------------------------

Write-Host "`n[2/3] Resolving categories..." -ForegroundColor Cyan

$resolved     = [System.Collections.Generic.List[PSCustomObject]]::new()
$unclassified = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

$total = $allFiles.Count
$index = 0

foreach ($file in $allFiles) {
    $index++
    Write-Progress -Activity 'Reading metadata' `
                   -Status   "$index / $total : $($file.Name)" `
                   -PercentComplete (($index / $total) * 100)

    try {
        $category = Get-FileCategory -Extension $file.Extension

        if ($null -ne $category) {
            $resolved.Add([PSCustomObject]@{
                File     = $file
                Category = $category
            })
        }
        else {
            # Extension not in CategoryMap — send to Unclassified\
            $unclassified.Add($file)
        }
    }
    catch {
        Write-Warning "Could not process '$($file.FullName)': $_"
        $unclassified.Add($file)
    }
}
Write-Progress -Activity 'Reading metadata' -Completed

Write-Host "      Categorised   : $($resolved.Count) file(s)"
Write-Host "      Unclassified  : $($unclassified.Count) file(s)"


# ---------------------------------------------------------------------------
# Step 3 — Copy / Move into Type\Year\Month structure
# ---------------------------------------------------------------------------

Write-Host "`n[3/3] Copying files to type folders..." -ForegroundColor Cyan

$stats = @{
    Copied        = 0
    Moved         = 0
    Skipped       = 0
    Unclassified  = 0
    Errors        = 0
}

$index = 0
$total = $resolved.Count

foreach ($item in $resolved) {
    $index++
    Write-Progress -Activity 'Sorting' `
                   -Status   "$index / $total : $($item.File.Name)" `
                   -PercentComplete (($index / $total) * 100)

    try {
        # Structure: OutputPath\Category\ (flat — no date subfolders)
        $destFolder = Join-Path $OutputPath $item.Category

        $result = Invoke-FileTransfer `
            -FilePath     $item.File.FullName `
            -DestFolder   $destFolder `
            -DeleteSource $Move.IsPresent

        switch -Wildcard ($result.Action) {
            'Moved'    { $stats.Moved++ }
            'Copied'   { $stats.Copied++ }
            'Skipped*' { $stats.Skipped++ }
        }

        Write-Verbose "$($result.Action): '$($item.File.Name)' [$($item.Category)] → $($result.Path)"
    }
    catch {
        Write-Warning "Failed to process '$($item.File.FullName)': $_"
        $stats.Errors++
    }
}

# Unclassified files → Unclassified\ folder for manual review
foreach ($file in $unclassified) {
    try {
        Invoke-FileTransfer `
            -FilePath     $file.FullName `
            -DestFolder   $unclassifiedFolder `
            -DeleteSource $Move.IsPresent | Out-Null
        $stats.Unclassified++
    }
    catch {
        Write-Warning "Failed to copy unclassified '$($file.FullName)': $_"
        $stats.Errors++
    }
}

Write-Progress -Activity 'Sorting' -Completed

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  Document Sort Complete"                  -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Mode              : $modeLabel"
Write-Host "  Files found       : $($allFiles.Count)"
if ($Move) { Write-Host "  Moved             : $($stats.Moved)" }
else       { Write-Host "  Copied            : $($stats.Copied)" }
if ($stats.Skipped       -gt 0) { Write-Host "  Skipped (identical)  : $($stats.Skipped)"                      -ForegroundColor Yellow }
if ($stats.Unclassified  -gt 0) { Write-Host "  Unclassified         : $($stats.Unclassified)  →  $unclassifiedFolder" -ForegroundColor Yellow }
if ($stats.Errors        -gt 0) { Write-Host "  Errors               : $($stats.Errors)"                       -ForegroundColor Red }
Write-Host ""

# Per-category breakdown
Write-Host "  Files sorted by category:" -ForegroundColor Cyan
foreach ($cat in $CategoryMap.Keys) {
    $catExts   = $CategoryMap[$cat]
    $catSorted = @($resolved | Where-Object { $_.Category -eq $cat }).Count
    if ($catSorted -gt 0) {
        Write-Host ("    {0,-14}: {1}" -f $cat, $catSorted)
    }
}
Write-Host "========================================`n" -ForegroundColor Green
