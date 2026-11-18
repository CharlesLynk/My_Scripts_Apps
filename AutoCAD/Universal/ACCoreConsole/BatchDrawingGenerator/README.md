# AutoCAD batch drawing generator

This Windows PowerShell/Core Console tool creates one DWG per CSV row. It uses the title-block inserts already present in the selected DWT. A border XREF can optionally be attached or refreshed during preflight; if no path is supplied, existing template XREFs are retained and validated, and generation can proceed with no XREFs. When a border XREF is selected, its name is taken from the DWG filename, without its extension; it is not a CSV field. Rows can route drawings into relative subfolders under the output directory. Preflight and generation are grouped by destination folder so relative XREFs are checked from the same directory as the resulting drawings. Existing output files are never overwritten.

## Requirements

- Windows PowerShell and an installed AutoCAD version that includes `accoreconsole.exe`.
- Core Console is launched with the `accoreconsole` profile (`/p accoreconsole`) for both preflight and drawing jobs.
- A DWT containing the title-block inserts and attribute tags referenced by the CSV. Target blocks must be top-level inserts in paper space; nested blocks are not searched.
- Optional border XREF DWG. When supplied, it is attached in paper space at `(0,0,0)` with unit scale and overlay mode. The filename without `.dwg` is used as its XREF block name. Without one, any XREFs already in the DWT are preserved and checked for resolution.
- A CSV with a `DrawingRelativePath` column. Every other column must use `BLOCKNAME.ATTRIBUTENAME`. Replace the illustrative names in `sample_drawings.csv` with the names in your drawing.
- The AutoLISP helper must be in an AutoCAD trusted location when `SECURELOAD` is enabled. Add the `BatchDrawingGenerator` folder to AutoCAD's **Options > Files > Trusted Locations** (the `TRUSTEDPATHS` setting), or ask IT to approve it. The script does not bypass AutoCAD's executable-code security.

Attribute values and file paths are passed as text. The generator does not depend on a separately installed Python runtime.

## Run a small test batch

Edit `sample_drawings.csv` so each qualified header matches an actual paper-space block insert and attribute tag. Run without path arguments to browse for the DWT, CSV, and output folder. For the optional border XREF prompt, enter a DWG path, type `B` to browse, or press Enter to leave it blank. If `accoreconsole.exe` is not found in common AutoCAD install locations, the script prompts you to locate it:

```powershell
.\Generate-Drawings.ps1 -MaxDrawings 1
```

Relative input and output paths, such as `.\sample_drawings.csv` or `.\run1`, are resolved from the current PowerShell location.

You can also pass paths to skip the file/folder dialogs:

```powershell
.\Generate-Drawings.ps1 `
  -AccoreConsolePath "C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe" `
  -TemplatePath "C:\CAD\TemplateName.dwt" `
  -CsvPath ".\sample_drawings.csv" `
  -OutputDirectory ".\" `
  -BorderXrefPath "C:\Users\YourName\Autodesk\Forma\Project\Border.dwg" `
  -MaxDrawings 1
```

Preflight runs once for each distinct destination folder before output generation. If a border XREF path is supplied, preflight attaches or updates it; otherwise, it leaves the template's XREFs unchanged. Every mapped attribute is validated on each matching top-level paper-space insert. An unresolved relative XREF path produces one consolidated warning per distinct reference and does not stop generation; it may resolve after the drawing is placed in its final location. An unresolved absolute XREF path remains an error. No XREF is required when the path is left blank. Preflight reports successful and missing block/attribute mappings before aborting on mapping failures. Each prepared drawing is placed temporarily in its destination folder so XREF paths are checked from the same folder as the generated drawings. A preflight failure stops the batch before drawing output begins. For each folder, Core Console opens the prepared drawing and saves that folder's CSV rows to unique staged DWGs. Completed staged drawings are moved to their final names after Core Console exits; if a row fails, earlier completed drawings are still retained, while the failed and unprocessed staged files are removed.

The script processes CSV rows sequentially and stops on the first Core Console error or missing success marker. A normal run uses two Core Console launches per distinct destination folder: one for preflight and one for generation. It prints concise progress messages, a consolidated warning for unresolved relative XREFs, and a readable failure summary; detailed AutoCAD output remains in the saved logs. Console output is captured in memory and discarded for successful runs apart from the consolidated XREF warning; the output folder's `logs` subfolder is written only for preflight or generation failures by default. Pass `-KeepLogs` to also retain logs for successful preflight and generation runs. If AutoCAD logs `File load canceled` while loading `CreateFromTemplate.lsp`, add the generator directory to Trusted Locations and retry. Use `-WhatIf` to preview target drawing paths without creating the destination folder or launching Core Console. After checking a small batch, increase `-MaxDrawings` or omit it to process all rows.

## Import generated drawings into Plant 3D

Plant 3D Project Manager properties are handled separately from the AutoLISP drawing generation. After generating DWGs, place the `sample_databaseFields.csv` metadata manifest in the output directory beside its drawing subfolders, update its `DrawingRelativePath` values and properties, then open the collaborative project in the full Plant 3D application. Use the same relative `DrawingRelativePath` paths in both CSVs so that the generated folder structure and Plant project folders match. Build and load the add-in in [Plant3DMetadataAddin](./Plant3DMetadataAddin/README.md), then run `P3DIMPORTDRAWINGPROPERTIES`. It validates the CSV and source files, prompts for a project part, previews the batch, and requires typing `ADD` before making project changes. Folder paths in the CSV locate source DWGs below the CSV directory and must correspond to folders already created through Plant 3D Project Manager under the selected PID DWG folder. The importer will list missing folders and stop before adding drawings; create those folders manually and rerun the import. Start with one drawing in a test project. Check in successfully added drawings through Project Manager to publish them to the collaborative project.

## CSV behavior

`DrawingRelativePath` is a path relative to the selected output directory; `.dwg` is added to its final filename when omitted. For example, `Piping\Area-1\P-001` or `Piping/Area-1/P-001.dwg` creates the drawing under those subfolders. A filename without a folder path is placed directly in the output directory. Use the same path relative to the metadata CSV's directory, which should be the common/base PID DWG folder, when importing the drawings into Plant 3D. Paths must not be rooted or contain `.` or `..` segments; the output folder named `logs` is reserved for generator logs. The older `DrawingFileName` header remains accepted for existing CSVs, but new CSVs should use `DrawingRelativePath`; do not include both headers. Every other CSV header is parsed as exactly one `BLOCKNAME.ATTRIBUTENAME` pair. The value in that column is written to the named attribute for each matching paper-space insert. If a block or attribute is missing from any matching insert, preflight fails with an error identifying the block and attribute. When the same named block is inserted multiple times in paper space, the CSV value is applied to all of those inserts.

Every data record must have exactly the same number of fields as the header. The CSV is parsed with quote-aware comma handling, so commas inside quoted values are allowed; malformed quoting or a missing/extra field aborts the run with a record/line-specific error before drawing generation starts.

The target block names and attribute names come only from CSV headers; the script has no default block name. If a border XREF file is selected at run time, the XREF block name is derived from that file's name. Files already present in the output directory cause the run to stop rather than overwrite them.
