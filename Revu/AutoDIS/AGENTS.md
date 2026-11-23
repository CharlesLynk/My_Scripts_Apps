# AGENTS.md — AutoDIS handoff

Read this first. It records what the project is, what has been changed, what the user is testing now, and what to do next.

## Purpose

AutoDIS (Mbroh Drawing Index Creator) is a WinUI 3 desktop app on .NET 10 (x64 only). It reads a set of drawing-sheet PDFs, OCRs each sheet's title block with Tesseract, extracts the sheet number, title, revision and related fields, and writes a drawing index PDF. Output is written atomically, so a locked target file is kept as a `.partial` file instead of being lost.

Packaging: MSIX (`Package.appxmanifest`), Package Name `d2bbdbd4-3016-443a-9f1f-e2344ac6a034`. The name must stay stable so upgrades install over the old package.

## Layout

- `MainWindow.xaml.cs` — almost all logic (~4,850 lines). Main pipeline: click handler about line 2993, setup about 3501, sheet loop about 3544–3641, geometry about 3702, export about 3768. `ExportResultsToPdf` is about 4326–4797.
- `Diagnostics.cs` — timer and log system. Shared with the tests through a linked source.
- `PdfContentRedactor.cs`, `AppSupport.cs` — redaction and shared helpers (also linked into tests).
- `Assets\`, `tessdata\` (copied to output by the csproj), `Properties\AssemblyInfo.cs`.
- `scripts\Build-And-Install.ps1` — build, sign, trust check, uninstall, install.
- `tests\AutoDIS.Tests\` — xUnit, 11 tests (`DiagnosticsTests`, `FileHygieneTests`, `RedactionTests`). All pass.
- `reference\` — trimmed to the x64 baseline package (1.3.4.0). The `.cer` is required by the script.
- `Directory.Packages.props` — central package versions. Fixed: PdfSharp 6.0.0 and Tesseract 4.1.1. Deferred: PdfPig 0.1.13 → 0.1.16. Current: WindowsAppSDK 2.7.0.

## Build and test

```
dotnet build MbrohDrawingIndexCreator.csproj -c Release -p:Platform=x64
cd tests\AutoDIS.Tests; dotnet test
```

Build output has about 258 nullable warnings (mostly CS8600/CS8602). These are existing, not regressions. Do not treat them as failures.

Install and test on the machine (run from the AutoDIS folder):

```
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\Build-And-Install.ps1 -Diagnostics
```

Flags: `-Diagnostics` enables timers, `-NoDiagnostics` disables them, `-SelfSign` signs as the current user. Each rebuild must be uninstalled and reinstalled. The package identity is fixed, so an upgrade over the same version is not possible. The script handles this.

## Signing and trust (current state)

- The Mbroh key is `\\FS01\Public\IT\Internal Software\AutoDIS\MbrohDrawingIndexCreator.Package_TemporaryKey.pfx`. It has no password. The `P:` mapping is not visible to PowerShell, so the script uses the UNC path. Never copy the key into the repo. `.pfx`, `.p12`, `.pvk` and `.snk` are ignored.
- Installs fail with 0x800B0109 until the root is trusted. An administrator must run `Import-Certificate -FilePath "<reference .cer>" -CertStoreLocation Cert:\LocalMachine\TrustedPeople`. This was still pending when the handoff was written.
- When the key is missing or expired, the script asks whether to sign as the current user (`CN=<username>`, stored in `Cert:\CurrentUser\My`, public cert at `%LOCALAPPDATA%\AutoDIS\personal-signing.cer`). The manifest publisher is changed temporarily and restored in `finally`. The user has approved this fallback.

## Optimizations already made

1. Per-page word geometry cache (`WordGeometry`, `GetPageWordGeometry`, `BuildPageWordGeometry`). Removes repeated word extraction per page.
2. Memoized `MeasureWrappedHeight` (`WrappedHeightCache`).
3. Single-pass folder mode.
4. Word cache, atomic partial-file output, and the switch to MSIX packaging.

The user reported a large speed gain on a large test set.

OCR is loaded eagerly at the start of every run. This is a deliberate user decision. Lazy OCR was rejected because it would stall the run mid-way.

## Diagnostics

- Enabled by `AUTODIS_DIAGNOSTICS=1` or by the marker file `%TEMP%\AutoDIS\diagnostics.on`. The marker file is the reliable switch for packaged apps. Timer sites are tagged `// DIAG`.
- Logs: `%TEMP%\AutoDIS\diagnostics\yyyyMMdd_HHmmss_fff_<run>.log`. The newest 20 are kept.
- Summary lines contain `x<count>`. Exclude them when adding totals, or the time is counted twice.
- The "sheet loop" and "total" timers overlap their sub-labels. Do not sum them.
- Writes after `EndRun` are dropped. The final commit step is not timed.

Last measured run (332 sheets, 31.5 s total, before the timers were expanded):
- `open sheet`: 20.3 s, about 65% of the run. It is I/O-bound on cold network or OneDrive reads.
- `export pdf`: 6.0 s.
- About 4.1 s before the first sheet was untimed. The current build times it as `ocr engine init`.

The user is running the new timer build now. Read the new log before changing any code.

## Cleanup done (this round)

- Removed the unused `Diagnostics.Note` method.
- Moved superseded reference files to `%TEMP%\AutoDIS-retired-reference-2026-11-23\` (about 88 MB). Moved, not deleted, because the `Revu\` folder is untracked by git and cannot be restored from history. This folder holds the top-level `Add-AppDevPackage.ps1`, `Install.bat`, `Install.ps1`, `index.html`, the `.appinstaller`, and the ARM64, x86 and win32 dependency folders.
- Kept: the x64 runtime dependencies, the `.cer`, and the old 1.3.4.0 baseline package. The baseline is the timing comparison point.
- `.gitignore` now covers per-platform output folders (anchored to the root), `Generated Files/`, `*.user`, `*.suo`, `*.binlog`, coverage files, `*.appxsym`, `*.appxbundle`, key formats (`*.p12`, `*.pvk`, `*.snk`), and `*.partial`.
  - Note: the `Dependencies\` redistributables are still ignored by the root repo's `*.appx` rule and by the existing AutoDIS `*.msix` rule. The `.cer` stays tracked. Un-ignoring the redistributables is a decision for the user.

## Known issues (not fixed)

- `TemplateHelpButton_Click` (about line 460) expects `Assets\AutoDISTemplateInstructions.docx`, which is missing. The app shows a message when it's absent.
- `Properties\AssemblyInfo.cs` reports 1.0.0.0. The MSIX version is 1.4.0.0. `GenerateAssemblyInfo` is false, so the value is not used for the package.
- The browse-button crash fix (MSIX packaging) is not yet confirmed on an installed package.
- The root repo's `.gitignore` already has uncommitted changes from outside this work. Do not touch them.
- Nothing has been committed. The `Revu\` folder is untracked.

## Next candidate optimizations (need approval before implementing)

1. Parallel sheet opening. The biggest lever, since `open sheet` dominates. Must keep the order of output and handle OCR engine thread safety. Decide after reading the timer log.
2. Remove the PNG round-trip in the OCR path, if the timers confirm it's costly.
3. Cache parsed template data instead of re-parsing per run.
4. Upgrade PdfPig 0.1.13 → 0.1.16 (deferred for risk review).
5. Reduce the nullable warning backlog in `MainWindow.xaml.cs`. This is type-safety work, not performance.

PdfSharp 6.0.0 and Tesseract 4.1.1 are fixed and must not be upgraded without a new decision from the user.

## Working rules for this project

- OCR stays eager. Do not make it lazy.
- Diagnostics stay toggleable and tagged `// DIAG`.
- Make surgical changes. Do not commit unless the user asks.
- Comment only code that needs clarification.
- Verify with a Release build and `dotnet test`. Delete `TestResults` after running tests.
- Use Windows paths. The shell is PowerShell 5.1, so use `;` and not `&&`.
