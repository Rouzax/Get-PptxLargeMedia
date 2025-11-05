# Get-PptxLargeMedia

PowerShell script to **find large images/media in a PowerPoint (.pptx)**, show **where they’re used**, and help you **trim deck size**. It unpacks the .pptx, analyzes embedded media, maps usage to slides (with **shape names**), also checks **masters/layouts/notes/charts**, and optionally detects **exact duplicates**.

---

## Features

* 🔍 **Find largest media** by **Top N** or by **minimum size (KB)**
* 🧭 **Where-used mapping**

  * `Slides`: slide numbers
  * `ShapeHints`: per-slide picture **shape names** (from Selection Pane)
  * `OtherRefs`: references from **Masters, Layouts, Notes, Charts**
* 🗂️ **CSV export** for reporting
* ♻️ **Duplicate detection** (SHA-256 exact matches)
* 🧹 **Automatic temp cleanup** (opt-out with `-KeepTemp`)
* 🧪 **Kind filter**: Images / Video / Audio / All

---

## Requirements

* Windows PowerShell **5.1+** or PowerShell **7+**
* PowerPoint file with **.pptx** extension (embedded media lives under `ppt/media`)

---

## Installation

1. Save the script as `Get-PptxLargeMedia.ps1`.
2. (Optional) Unblock the file if downloaded:

   ```powershell
   Unblock-File .\Get-PptxLargeMedia.ps1
   ```
3. Ensure your execution policy allows running local scripts:

   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   ```

---

## Usage

### Basic

```powershell
# Top 10 largest media (default)
.\Get-PptxLargeMedia.ps1 -Path '.\Deck.pptx' -Top 10
```

```powershell
# Everything ≥ 500 KB
.\Get-PptxLargeMedia.ps1 -Path '.\Deck.pptx' -MinKB 500
```

### Filter by kind

```powershell
# Only images
.\Get-PptxLargeMedia.ps1 -Path '.\Deck.pptx' -MinKB 1 -Kind Images
```

### Detect exact duplicates

```powershell
# Show duplicates across entire deck (example pipe)
.\Get-PptxLargeMedia.ps1 -Path '.\Deck.pptx' -MinKB 1 -DetectDuplicates -PassThru |
  Where-Object _Hash |
  Group-Object _Hash | Where-Object Count -gt 1 |
  ForEach-Object { $_.Group } |
  Sort-Object SizeBytes -Descending |
  Format-Table SizeKB, Kind, FileName, Slides, OtherRefs, ShapeHints -AutoSize
```

### Export to CSV

```powershell
.\Get-PptxLargeMedia.ps1 -Path '.\Deck.pptx' -MinKB 1 -DetectDuplicates -ExportCsv .\media-report.csv
```

---

## Parameters

* `-Path <string>`: **Required.** Path to `.pptx` (supports brackets and special chars).
* `-Top <int>`: Return the N largest media (default **10**). Mutually exclusive with `-MinKB`.
* `-MinKB <int>`: Return all media with size ≥ KB threshold. Mutually exclusive with `-Top`.
* `-Kind <All|Images|Video|Audio>`: Filter by media kind (default **All**).
* `-DetectDuplicates`: Compute SHA-256; add duplicate info fields.
* `-ExportCsv <path>`: Write results to CSV (UTF-8).
* `-KeepTemp`: Keep the temp extraction folder.
* `-PassThru`: Output result objects to the pipeline (useful for further processing).

---

## Output Columns

* **SizeKB**: File size in KB (rounded).
* **Kind**: Images / Video / Audio / Other.
* **FileName**: Name as embedded in `ppt/media`.
* **Slides**: Slide numbers where the media is embedded.
* **OtherRefs**: Usage from parts not tied to a slide (e.g., `Layout:5`, `Master:1`, `Notes:3`, `Chart:4`).
* **ShapeHints**: Per-slide **picture shape names**, e.g., `9: Picture 21 | 10: Picture 3`.
* **Orphaned**: `True` if not referenced by slides **or** other parts (likely removable).
* *(When `-DetectDuplicates` is used)*

  * **DupGroup**: Short hash (first 8 chars) for duplicate grouping (blank if unique).
  * **DupCount**: Number of files in the duplicate group.

**Example**

```
SizeKB   Kind   FileName     Slides OtherRefs       ShapeHints                              Orphaned DupGroup DupCount
------   ----   --------     ------ ---------       ----------                              -------- -------- --------
17079.86 Images image59.png  9,10                   9: Picture 21 | 10: Picture 3           False                 1
 3665.57 Images image3.png          Layout:1                                                  False                 1
 1894.57 Images image39.emf  1                       1: think-cell data - do not delete      False                 1
```

---

## How it works

1. **Unpacks** the `.pptx` (ZIP) to a unique temp folder.
2. Scans `ppt/media` for embedded binaries and collects sizes/types.
3. **Maps usage** by reading relationship files:

   * Slides: `ppt/slides/_rels/slide*.xml.rels`
   * Masters: `ppt/slideMasters/_rels/slideMaster*.xml.rels`
   * Layouts: `ppt/slideLayouts/_rels/slideLayout*.xml.rels`
   * Notes: `ppt/notesSlides/_rels/notesSlide*.xml.rels`
   * Charts: `ppt/charts/_rels/chart*.xml.rels`
4. For each slide, parses `slide*.xml` to **resolve picture shape names** for each referenced media.
5. **Cleans up** temp folder (unless `-KeepTemp`).

---

## Tips to Reduce Deck Size

* Replace massive images/captures with appropriately sized versions for your slide resolution
* **Compress Pictures** (Picture Format → Compress Pictures)
* **Delete cropped areas** and **allow compression** (File → Options → Advanced → *Image Size and Quality*)
* **Compress Media** (File → Info)
* **Save As** a new file to purge stale parts

---

## Troubleshooting

* **“File not found” with brackets**: The script uses `-LiteralPath`, so paths like `[name].pptx` are supported. If you still see it, verify the file exists and the extension is `.pptx`.
* **No output / Orphaned media**: Some media may be in masters/layouts/notes/charts—check **OtherRefs**. Truly orphaned items are likely removable (work on a copy if removing by hand).
* **Duplicates don’t appear**: Use `-MinKB 1 -DetectDuplicates` to include everything (instead of `-Top`), and consider piping (see example).

---

## Known Limitations

* **Exact duplicates only** (SHA-256). Near-duplicates (re-encoded/cropped) won’t match.
* Only **embedded** media (`ppt/media`) are analyzed. **Linked** external media aren’t included.
* Shape name extraction covers **pictures** (not every possible OOXML element).