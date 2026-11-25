# Bench tools

Small console projects used to measure OCR accuracy and speed against hand-read answers. They are not part of the app or its tests (the app project excludes `tools\**`). Each builds on its own (`dotnet run -c Release` in its folder). They link the app's source files directly, using absolute paths in their `.csproj` files, so edit those paths if the repository moves. Data paths inside `Program.cs` point at `%TEMP%` and need the same care.

Test data is not committed. Recreate it under `%TEMP%`:

| Path | What |
|---|---|
| `%TEMP%\bachdata\bach.pdf` | the 380-sheet DWU Bachman set (client drawings; get it from the project share) |
| `%TEMP%\bachdata\mini.pdf` | a small set with a few OCR sheets |
| `%TEMP%\gt\truth.txt` | hand-read answers, one line per sheet: `page\|sheet id\|discipline\|title` (47 Bachman sheets: pages up to 337 are vector-lettered, 338 and up are scans). A copy is in `tools\bench\data\` (ignored by git) and in the Copilot session `files\bachman-truth.txt` |

| Tool | Use |
|---|---|
| `eval` | `dotnet run -- <export_log.txt> <truth.txt> [runIndex]` scores every `OCR ready` run in the app log against the truth (sheet-number, discipline and title errors); the run index lists the wrong entries. Reads the post-analysis `OCR page N ...` lines. |
| `vecbench` | reads the vector-lettered sheets with the real recogniser and writes every candidate to a TSV. Env: `EV_DPI`, `EV_TARGET` (letter height), `EV_OVER` (oversample), `EV_MAX`, `EV_TAG`. `large` mode tests sheet IDs with the large model. |
| `replay` | runs the current `IndexAnalyzer` over a `vecbench` TSV and prints errors against the truth (`<tsv> v <truth>`). Needs the `dictionary` folder next to its exe, including the downloaded `english-enable.txt`. |
| `scanbench` | renders scan pages with PDFium at a DPI and prints the three readings beside the truth (`<dpi> <target> [idtest pages]`). |
| `gt` | renders montages of title blocks for reading answers by eye. |
| `imginfo` | prints page sizes and picture resolution of a page. |
| `textbench` | times PdfPig against PDFium text lookup on a sample of pages (2026-11-24: PDFium 3x faster per page but disagreed on 11 of 43 pages). |

Rules of thumb that came out of using them: do not run an experiment unless it answers a specific question; measure before and after a change on the truth set; scans and vector sheets behave differently, so score them separately.
