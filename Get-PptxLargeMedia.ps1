<#
.SYNOPSIS
Find large images/media in a PowerPoint (.pptx) and show where they are used.

.DESCRIPTION
Unpacks a .pptx into a temporary folder, inspects ppt/media for embedded media, 
and reports the largest items either by "top N" or by a minimum size threshold.
For each media file, the script maps its usage to:
  - Slides (with per-slide picture shape names), and
  - Other PowerPoint parts: slide masters, slide layouts, notes slides, and charts.

Optionally exports a CSV report and detects duplicate media by file hash.
By default the temporary folder is removed.

.PARAMETER Path
Path to a .pptx  or .potx file. Literal path is supported (e.g., names with [ and ]).

.PARAMETER Top
Return the N largest media files (sorted by size). Mutually exclusive with -MinKB.
Default: 10.

.PARAMETER MinKB
Return every media file whose size is >= MinKB (kilobytes). Mutually exclusive with -Top.

.PARAMETER Kind
Filter for media kinds. One of: All, Images, Video, Audio. Default: All.

.PARAMETER DetectDuplicates
When set, compute SHA256 hashes to identify duplicate embedded media and annotate results.

.PARAMETER ExportCsv
If provided, export the final results to the specified CSV path (UTF-8).

.PARAMETER KeepTemp
When set, keep the temporary extraction folder for your inspection. Default is to delete it.

.PARAMETER PassThru
When set, also return the result objects to the pipeline (useful in scripts).

.EXAMPLE
PS> .\Get-PptxLargeMedia.ps1 -Path .\Deck.pptx -Top 15 -Kind Images -Verbose

.EXAMPLE
PS> .\Get-PptxLargeMedia.ps1 -Path .\Deck.pptx -MinKB 500 -DetectDuplicates -ExportCsv .\large-media.csv

.NOTES
- Works on Windows PowerShell 5.1+ and PowerShell 7+.
- Slide usage is resolved via slide relationship files; picture shape names come from slide XML.
- OtherRefs lists references from masters/layouts/notes/charts that aren’t tied to a specific slide.
#>

[CmdletBinding(DefaultParameterSetName = 'Top', PositionalBinding = $false)]
param(
    [Parameter(Mandatory, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
            if (-not (Test-Path -LiteralPath $_ -PathType Leaf)) {
                throw "File not found: $_"
            }
            $allowedExt = @('.pptx', '.potx')
            $ext = [System.IO.Path]::GetExtension($_).ToLowerInvariant()
            if ($allowedExt -notcontains $ext) {
                throw "File must have one of: $($allowedExt -join ', ')."
            }
            $true
        })]
    [string]$Path,


    [Parameter(ParameterSetName = 'Top')]
    [ValidateRange(1, 100000)]
    [int]$Top = 10,

    [Parameter(ParameterSetName = 'MinKB', Mandatory)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$MinKB,

    [ValidateSet('All', 'Images', 'Video', 'Audio')]
    [string]$Kind = 'All',

    [switch]$DetectDuplicates,

    [string]$ExportCsv,

    [switch]$KeepTemp,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Helper: classify media by extension -------------------------------------
function Get-MediaKind {
    param([string]$Extension)

    $ext = ''
    if ($Extension) {
        $ext = $Extension.ToLowerInvariant() 
    }

    $imageExt = '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.emf', '.wmf', '.svg'
    $videoExt = '.mp4', '.mov', '.wmv', '.avi', '.m4v'
    $audioExt = '.mp3', '.m4a', '.wav', '.wma'

    if ($imageExt -contains $ext) {
        return 'Images' 
    }
    if ($videoExt -contains $ext) {
        return 'Video' 
    }
    if ($audioExt -contains $ext) {
        return 'Audio' 
    }
    return 'Other'
}

# --- Helper: names of picture shapes embedding a media file on a slide -------
function Get-PicNamesOnSlideForFile {
    param(
        [Parameter(Mandatory)][string]$SlidesDir,
        [Parameter(Mandatory)][int]$SlideNum,
        [Parameter(Mandatory)][string]$MediaFileName
    )

    $relsPath = Join-Path $SlidesDir ("_rels\slide{0}.xml.rels" -f $SlideNum)
    $slidePath = Join-Path $SlidesDir ("slide{0}.xml" -f $SlideNum)
    if (-not (Test-Path -LiteralPath $relsPath -PathType Leaf)) {
        return @() 
    }
    if (-not (Test-Path -LiteralPath $slidePath -PathType Leaf)) {
        return @() 
    }

    try {
        [xml]$rels = Get-Content -LiteralPath $relsPath -Encoding UTF8
        [xml]$slide = Get-Content -LiteralPath $slidePath -Encoding UTF8

        # rIds whose Target == ../media/<MediaFileName> (or media/<MediaFileName>)
        $rIds = @()
        foreach ($r in @($rels.Relationships.Relationship)) {
            $target = ''
            if ($r.Target) {
                $target = $r.Target.ToString() 
            }
            if ($target -match '(^\.\./)?media/(.+)$' -and $Matches[2] -eq $MediaFileName) {
                $rIds += $r.Id
            }
        }
        if (-not $rIds) {
            return @() 
        }

        $names = @()
        foreach ($rid in $rIds) {
            # Find <p:pic> referencing this rId; read its cNvPr name/descr
            $xpath = "//*[local-name()='pic'][.//*[local-name()='blip' and @*[local-name()='embed' and .='$rid']]]/*[local-name()='nvPicPr']/*[local-name()='cNvPr']"
            $nodes = Select-Xml -Xml $slide -XPath $xpath
            foreach ($n in @($nodes)) {
                $node = $n.Node
                $name = $node.GetAttribute('name')
                $descr = $node.GetAttribute('descr')
                if ($name) {
                    $names += $name; continue 
                }
                if ($descr) {
                    $names += $descr; continue 
                }
            }
        }
        return ($names | Sort-Object -Unique)
    } catch {
        Write-Verbose "Failed to resolve shape names for slide $SlideNum / $($MediaFileName): $($_.Exception.Message)"
        return @()
    }
}

# --- Prepare temp and unzip ---------------------------------------------------
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue 
} catch {
}

$ResolvedPath = (Resolve-Path -LiteralPath $Path).Path

$guid = [Guid]::NewGuid().ToString('N')
$tempRoot = Join-Path -Path $env:TEMP -ChildPath ("PPTX_Unpack_{0}" -f $guid)
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

Write-Verbose "Extracting '$ResolvedPath' to '$tempRoot'..."
[System.IO.Compression.ZipFile]::ExtractToDirectory($ResolvedPath, $tempRoot)

try {
    # --- Discover media files -------------------------------------------------
    $pptDir = Join-Path $tempRoot 'ppt'
    $mediaDir = Join-Path $pptDir 'media'

    if (-not (Test-Path -LiteralPath $mediaDir -PathType Container)) {
        Write-Warning "No 'ppt/media' folder found in the presentation (no embedded media detected)."
        $records = @()
        $selected = @()
        return
    }

    $mediaFiles = Get-ChildItem -LiteralPath $mediaDir -File -ErrorAction Stop
    if (-not $mediaFiles) {
        Write-Warning "No media files found under '$mediaDir'."
    }

    # --- Build usage maps -----------------------------------------------------
    $usageByFile = @{}  # filename -> List[int] of slide numbers
    $otherRefsByFile = @{}  # filename -> List[string] of contexts (Master:n | Layout:n | Notes:n | Chart:n)

    function Add-OtherRef {
        param([string]$FileName, [string]$Context)
        if (-not $otherRefsByFile.ContainsKey($FileName)) {
            $otherRefsByFile[$FileName] = New-Object System.Collections.Generic.List[string]
        }
        $null = $otherRefsByFile[$FileName].Add($Context)
    }

    # -- Slides: map media to slide numbers + gather shape names later --------
    $slidesDir = Join-Path $pptDir 'slides'
    $relsDir = Join-Path $slidesDir '_rels'

    if (Test-Path -LiteralPath $relsDir -PathType Container) {
        $relFiles = Get-ChildItem -LiteralPath $relsDir -Filter 'slide*.xml.rels' -File
        foreach ($rel in $relFiles) {
            if ($rel.BaseName -notmatch 'slide(\d+)\.xml') {
                continue 
            }
            $slideNum = [int]$Matches[1]

            try {
                [xml]$relsXml = Get-Content -LiteralPath $rel.FullName -Encoding UTF8
                $relationships = @($relsXml.Relationships.Relationship)
                foreach ($r in $relationships) {
                    $target = ''
                    if ($r.Target) {
                        $target = $r.Target.ToString() 
                    }
                    if ($target -match '(^\.\./)?media/(.+)$') {
                        $fileName = $Matches[2]
                        if (-not $usageByFile.ContainsKey($fileName)) {
                            $usageByFile[$fileName] = New-Object System.Collections.Generic.List[int]
                        }
                        $null = $usageByFile[$fileName].Add($slideNum)
                    }
                }
            } catch {
                Write-Verbose "Failed to parse rels file '$($rel.FullName)': $($_.Exception.Message)"
            }
        }
    } else {
        Write-Verbose "No relationships dir found at '$relsDir'; slide usage mapping may be incomplete."
    }

    # -- Masters ---------------------------------------------------------------
    $mastersRels = Join-Path $pptDir 'slideMasters\_rels'
    if (Test-Path -LiteralPath $mastersRels -PathType Container) {
        foreach ($rel in (Get-ChildItem -LiteralPath $mastersRels -Filter 'slideMaster*.xml.rels' -File)) {
            if ($rel.BaseName -notmatch 'slideMaster(\d+)\.xml') {
                continue 
            }
            $n = [int]$Matches[1]
            try {
                [xml]$relsXml = Get-Content -LiteralPath $rel.FullName -Encoding UTF8
                foreach ($r in @($relsXml.Relationships.Relationship)) {
                    $target = ''
                    if ($r.Target) {
                        $target = $r.Target.ToString() 
                    }
                    if ($target -match '(^\.\./)?media/(.+)$') {
                        Add-OtherRef -FileName $Matches[2] -Context ("Master:{0}" -f $n)
                    }
                }
            } catch {
                Write-Verbose "Failed to parse master rels '$($rel.FullName)': $($_.Exception.Message)"
            }
        }
    }

    # -- Layouts ---------------------------------------------------------------
    $layoutsRels = Join-Path $pptDir 'slideLayouts\_rels'
    if (Test-Path -LiteralPath $layoutsRels -PathType Container) {
        foreach ($rel in (Get-ChildItem -LiteralPath $layoutsRels -Filter 'slideLayout*.xml.rels' -File)) {
            if ($rel.BaseName -notmatch 'slideLayout(\d+)\.xml') {
                continue 
            }
            $n = [int]$Matches[1]
            try {
                [xml]$relsXml = Get-Content -LiteralPath $rel.FullName -Encoding UTF8
                foreach ($r in @($relsXml.Relationships.Relationship)) {
                    $target = ''
                    if ($r.Target) {
                        $target = $r.Target.ToString() 
                    }
                    if ($target -match '(^\.\./)?media/(.+)$') {
                        Add-OtherRef -FileName $Matches[2] -Context ("Layout:{0}" -f $n)
                    }
                }
            } catch {
                Write-Verbose "Failed to parse layout rels '$($rel.FullName)': $($_.Exception.Message)"
            }
        }
    }

    # -- Notes slides ----------------------------------------------------------
    $notesRels = Join-Path $pptDir 'notesSlides\_rels'
    if (Test-Path -LiteralPath $notesRels -PathType Container) {
        foreach ($rel in (Get-ChildItem -LiteralPath $notesRels -Filter 'notesSlide*.xml.rels' -File)) {
            if ($rel.BaseName -notmatch 'notesSlide(\d+)\.xml') {
                continue 
            }
            $n = [int]$Matches[1]
            try {
                [xml]$relsXml = Get-Content -LiteralPath $rel.FullName -Encoding UTF8
                foreach ($r in @($relsXml.Relationships.Relationship)) {
                    $target = ''
                    if ($r.Target) {
                        $target = $r.Target.ToString() 
                    }
                    if ($target -match '(^\.\./)?media/(.+)$') {
                        Add-OtherRef -FileName $Matches[2] -Context ("Notes:{0}" -f $n)
                    }
                }
            } catch {
                Write-Verbose "Failed to parse notes rels '$($rel.FullName)': $($_.Exception.Message)"
            }
        }
    }

    # -- Charts ----------------------------------------------------------------
    $chartsRels = Join-Path $pptDir 'charts\_rels'
    if (Test-Path -LiteralPath $chartsRels -PathType Container) {
        foreach ($rel in (Get-ChildItem -LiteralPath $chartsRels -Filter 'chart*.xml.rels' -File)) {
            if ($rel.BaseName -notmatch 'chart(\d+)\.xml') {
                continue 
            }
            $n = [int]$Matches[1]
            try {
                [xml]$relsXml = Get-Content -LiteralPath $rel.FullName -Encoding UTF8
                foreach ($r in @($relsXml.Relationships.Relationship)) {
                    $target = ''
                    if ($r.Target) {
                        $target = $r.Target.ToString() 
                    }
                    if ($target -match '(^\.\./)?media/(.+)$') {
                        Add-OtherRef -FileName $Matches[2] -Context ("Chart:{0}" -f $n)
                    }
                }
            } catch {
                Write-Verbose "Failed to parse chart rels '$($rel.FullName)': $($_.Exception.Message)"
            }
        }
    }

    # --- Assemble records -----------------------------------------------------
    $records = foreach ($f in $mediaFiles) {
        $mediaKind = Get-MediaKind -Extension $f.Extension
        if ($Kind -ne 'All' -and $mediaKind -ne $Kind) {
            continue 
        }

        # Slides list (numeric)
        $slideNums = @()
        if ($usageByFile.ContainsKey($f.Name)) {
            $slideNums = @(
                $usageByFile[$f.Name].ToArray() | Sort-Object -Unique
            )
        }

        # Other refs (contexts)
        $otherRefs = @()
        if ($otherRefsByFile.ContainsKey($f.Name)) {
            $otherRefs = @(
                $otherRefsByFile[$f.Name].ToArray() | Sort-Object -Unique
            )
        }

        # Per-slide shape hints (e.g., "9: Picture 21 | 10: Picture 3")
        $shapeHints = if ($slideNums.Count -gt 0) {
            ($slideNums | ForEach-Object {
                $names = @( Get-PicNamesOnSlideForFile -SlidesDir $slidesDir -SlideNum $_ -MediaFileName $f.Name )
                if ($names.Count -gt 0) {
                    "{0}: {1}" -f $_, ($names -join ', ') 
                } else {
                    "{0}" -f $_ 
                }
            }) -join ' | '
        } else {
            '' 
        }

        [PSCustomObject]@{
            FileName      = $f.Name
            Kind          = $mediaKind
            Extension     = $f.Extension.ToLowerInvariant()
            SizeKB        = [math]::Round($f.Length / 1KB, 2)
            SizeBytes     = [int64]$f.Length
            Slides        = if ($slideNums.Count -gt 0) {
                ($slideNums -join ',') 
            } else {
                '' 
            }
            OtherRefs     = if ($otherRefs.Count -gt 0) {
                ($otherRefs -join ' | ') 
            } else {
                '' 
            }
            ShapeHints    = $shapeHints
            Orphaned      = ($slideNums.Count -eq 0 -and $otherRefs.Count -eq 0)
            FullPath      = $f.FullName
            _Hash         = $null  # optionally populated below
            _DuplicateKey = $null  # optionally populated below
        }
    }

    # Ensure array so .Count is safe
    $records = @($records)

    if ($records.Count -eq 0) {
        Write-Verbose "No media matched the filter Kind='$Kind'."
    }

    # --- Optional duplicate detection ----------------------------------------
    if ($DetectDuplicates -and @($records).Count -gt 0) {
        Write-Verbose "Computing hashes for duplicate detection..."
        foreach ($r in $records) {
            try {
                $r._Hash = (Get-FileHash -LiteralPath $r.FullPath -Algorithm SHA256).Hash
            } catch {
                Write-Verbose "Hashing failed for '$($r.FullPath)': $($_.Exception.Message)"
            }
            # init visible fields
            Add-Member -InputObject $r -NotePropertyName DupGroup -NotePropertyValue $null -Force
            Add-Member -InputObject $r -NotePropertyName DupCount -NotePropertyValue 1 -Force
        }

        $groups = $records | Where-Object { $_._Hash } | Group-Object _Hash
        foreach ($g in $groups) {
            if ($g.Count -gt 1) {
                $groupId = $g.Name.Substring(0, 8)   # short hash
                $count = $g.Count
                foreach ($item in $g.Group) {
                    $item.DupGroup = $groupId
                    $item.DupCount = $count
                }
            }
        }
    }


    # --- Size selection (Top vs MinKB) ---------------------------------------
    $sorted = $records | Sort-Object -Property SizeBytes -Descending

    switch ($PSCmdlet.ParameterSetName) {
        'MinKB' {
            $selected = $sorted | Where-Object { $_.SizeKB -ge $MinKB } 
        }
        'Top' {
            $selected = $sorted | Select-Object -First $Top 
        }
        default {
            $selected = $sorted | Select-Object -First $Top 
        }
    }

    # --- Export CSV if requested ---------------------------------------------
    if ($ExportCsv) {
        $exportCols = 'FileName', 'Kind', 'Extension', 'SizeKB', 'Slides', 'OtherRefs', 'ShapeHints', 'Orphaned', 'FullPath'
        if ($DetectDuplicates) {
            $exportCols += '_Hash', '_DuplicateKey' 
        }
        $selected | Select-Object $exportCols | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $ExportCsv
        Write-Verbose "Exported CSV to '$ExportCsv'."
    }
    
    # --- Emit results ---------------------------------------------------------
    $displayCols = @('SizeKB', 'Kind', 'FileName', 'Slides', 'OtherRefs', 'ShapeHints', 'Orphaned')
    if ($DetectDuplicates) {
        $displayCols += 'DupGroup', 'DupCount' 
    }

    $selected | Sort-Object SizeBytes -Descending | Format-Table -AutoSize $displayCols


    if ($PassThru) {
        $selected | Sort-Object SizeBytes -Descending
    }

} finally {
    if ($KeepTemp) {
        Write-Verbose "Keeping temp folder: $tempRoot"
    } else {
        try {
            if (Test-Path -LiteralPath $tempRoot) {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction Stop
            }
        } catch {
            Write-Warning "Failed to remove temp folder '$tempRoot': $($_.Exception.Message)"
        }
    }
}
