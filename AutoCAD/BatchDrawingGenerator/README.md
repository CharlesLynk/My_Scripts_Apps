# AutoCAD batch drawing generator

This Windows PowerShell/Core Console tool creates one DWG per CSV row. It uses the title-block inserts already present in the selected DWT and attaches the selected static border XREF during a one-time preflight. The XREF name is taken from the selected DWG filename, without its extension; it is not a CSV field or a separate prompt. Generation runs in one Core Console session: it updates the mapped attributes and saves a staged DWG for each row. Existing output files are never overwritten.

## Requirements

- Windows PowerShell and an installed AutoCAD version that includes `accoreconsole.exe`.
- Core Console is launched with the `accoreconsole` profile (`/p accoreconsole`) for both preflight and drawing jobs.
- A DWT containing the title-block inserts and attribute tags referenced by the CSV. Target blocks must be top-level inserts in paper space; nested blocks are not searched.
- The border XREF DWG. It is attached in paper space at `(0,0,0)` with unit scale and overlay mode. The filename without `.dwg` is used as its XREF block name.
- A CSV with a `DrawingFileName` column. Every other column must use `BLOCKNAME.ATTRIBUTENAME`. Replace the illustrative names in `sample_drawings.csv` with the names in your drawing.
- The AutoLISP helper must be in an AutoCAD trusted location when `SECURELOAD` is enabled. Add the `BatchDrawingGenerator` folder to AutoCAD's **Options > Files > Trusted Locations** (the `TRUSTEDPATHS` setting), or ask IT to approve it. The script does not bypass AutoCAD's executable-code security.

Attribute values and file paths are passed as text. The generator does not depend on a separately installed Python runtime.

## Run a small test batch

Edit `sample_drawings.csv` so each qualified header matches an actual paper-space block insert and attribute tag. Run without path arguments to browse for the DWT, CSV, border DWG, and output folder. If `accoreconsole.exe` is not found in common AutoCAD install locations, the script prompts you to locate it:

```powershell
.\Generate-Drawings.ps1 -MaxDrawings 1
```

You can also pass paths to skip the file/folder dialogs:

```powershell
.\Generate-Drawings.ps1 `
  -AccoreConsolePath "C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe" `
  -TemplatePath "C:\CAD\TemplateName.dwt" `
  -CsvPath ".\sample_drawings.csv" `
  -OutputDirectory "C:\CAD\BatchTest" `
  -BorderXrefPath "C:\Users\YourName\Autodesk\Forma\Project\Border.dwg" `
  -MaxDrawings 1
```

The preflight runs once before output generation. It attaches or updates the border XREF, checks that all XREFs resolve, and validates every mapped attribute on every matching top-level paper-space insert. It reports all successful and missing block/attribute mappings before aborting on mapping failures. The prepared drawing is placed temporarily in the output folder so relative XREF paths are checked from the same folder as the generated drawings. A failure stops the batch before drawing output begins. The main generation session opens this prepared drawing once, then saves each CSV row to a unique staged DWG. Completed staged drawings are moved to their final names after Core Console exits; if a row fails, earlier completed drawings are still retained, while the failed and unprocessed staged files are removed.

The script processes CSV rows sequentially and stops on the first Core Console error or missing success marker. A normal run uses two Core Console launches total: one for preflight and one for the entire generation sequence. It prints concise progress messages and a readable failure summary; detailed AutoCAD output remains in the saved logs. Console output is captured in memory and discarded for successful runs; the output folder's `logs` subfolder is written only for preflight or generation failures by default. Pass `-KeepLogs` to also retain logs for successful preflight and generation runs. If AutoCAD logs `File load canceled` while loading `CreateFromTemplate.lsp`, add the generator directory to Trusted Locations and retry. Use `-WhatIf` to preview target drawing paths without creating the destination folder or launching Core Console. After checking a small batch, increase `-MaxDrawings` or omit it to process all rows.

## CSV behavior

`DrawingFileName` is the output DWG filename; `.dwg` is added when omitted. Every other CSV header is parsed as exactly one `BLOCKNAME.ATTRIBUTENAME` pair. The value in that column is written to the named attribute for each matching paper-space insert. If a block or attribute is missing from any matching insert, preflight fails with an error identifying the block and attribute. When the same named block is inserted multiple times in paper space, the CSV value is applied to all of those inserts.

The target block names and attribute names come only from CSV headers; the script has no default block name. The border XREF file is selected at run time, but the XREF block name is derived from that file's name. Files already present in the output directory cause the run to stop rather than overwrite them.
