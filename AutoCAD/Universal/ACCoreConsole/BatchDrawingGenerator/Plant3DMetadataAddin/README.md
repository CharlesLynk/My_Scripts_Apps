# Plant 3D Metadata Add-in

## Purpose

This add-in imports drawing metadata into the active Plant 3D project. It targets AutoCAD Plant 3D 2026 and uses the managed assemblies installed under `C:\Program Files\Autodesk\AutoCAD 2026`.

## Project layout

```text
Plant3DMetadataAddin/
├── AGENTS.md
├── README.md
├── .gitignore
├── Plant3DMetadataAddin.csproj
├── DrawingPropertyImportCommand.cs
├── DrawingPropertyUpdateCommand.cs
├── ProjectApiInfoCommand.cs
├── reference/
│   └── decompiled/   user-provided API references; plain-text, not source
├── bin/   generated build output; ignored
└── obj/   generated intermediate files; ignored
```

The [Project class decompilation](./reference/decompiled/PlantProjectManager.Project.decompiled.cs.txt) and [folder-properties dialog decompilation](./reference/decompiled/PlantProjectManagerUI.FolderPropertiesDialog.decompiled.cs.txt) are local reverse-engineering references supplied for this investigation. Their `.cs.txt` names keep them out of C# compilation, and `.gitignore` keeps them local to this workspace. They are not authoritative API contracts; do not edit or redistribute/commit them without confirming authorization.

Keep project-specific API probe output and build artifacts out of source control. The API information command prints to AutoCAD's command line; save output outside `bin/` if it is needed for a specific investigation. `P3DProjectProperties.log` is local runtime diagnostics and is ignored.

## Commands

| Command | Purpose | Project/drawing changes |
|---|---|---|
| `P3DPROJECTAPIINFO` | Reports the active project, property metadata, project items, project-part directory mappings, folder-path resolution, and loaded Project Manager API signatures. | No |
| `P3DIMPORTDRAWINGPROPERTIES` | Validates a metadata CSV, registers drawings, writes properties, and normalizes eligible XREF paths. | Yes |
| `P3DUPDATEDRAWINGPROPERTIES` | Updates properties of drawings already registered in the active PnId part, from a metadata CSV. Unregistered rows are skipped and reported. | Yes (property values only; no drawings are copied, registered, or created) |

Both commands require the full AutoCAD Plant 3D application with an active drawing. The import command uses the currently open, signed-in Plant 3D project and does not store or request Autodesk credentials.

## Build and load

Close Plant 3D before rebuilding if the add-in is loaded; AutoCAD locks the DLL. Build from this directory:

```powershell
dotnet build .\Plant3DMetadataAddin.csproj -c Debug
```

Load the result with `NETLOAD`:

```text
bin\Debug\net10.0-windows\Plant3DProjectApiProbe.dll
```

Do not use an alternate build output path. `bin/` and `obj/` are generated and ignored.

## Metadata CSV

The CSV can be stored anywhere. The required path header is `DrawingRelativePath`; the legacy `DrawingFileName` header remains accepted for existing files, but both headers may not appear in the same CSV.

- Source DWGs are resolved relative to the active Plant 3D project's local folder, using the path in `DrawingRelativePath`. A source DWG must already exist under that folder; the import copies it only if its project-folder destination differs.
- The importer always targets the active project's PnId part. Destination folders are relative to that part; a bare filename targets its root.
- Before import, the command checks all rows for missing sources, registered drawings, and destination-file collisions, then reports all conflicts together. Missing destination folders are listed for creation; no files or folders are changed during preflight.
- One `ADD` authorizes creating missing folders through the Project Manager workflow and importing the drawings. Root folders use a null parent identifier; nested folders use their parent GUID. For collaboration projects, each created folder is added through Plant's collaboration API and the folders-only refresh runs after a successful add.
- After the user confirms `ADD`, a source outside its intended project destination is copied into that destination folder before registration. A source already at the destination is registered in place. If the destination already contains an unregistered DWG with identical contents, that file is registered in place; a different file is never overwritten.
- Each other header must match a drawing-property name or display name available in the selected part. Matching ignores case, spaces, and underscores.
- `DWGTitle`, `DWGNumber`, `DrawingArea`, and `FacilityName` are examples of supported property headers.

The sample is [`../sample_databaseFields.csv`](../sample_databaseFields.csv).

## Import workflow

1. Sign in to Plant 3D and open the intended project.
2. Run `P3DIMPORTDRAWINGPROPERTIES`.
3. Select the CSV and review preflight results, including every folder to create, planned file copies, and any aggregated conflicts.
4. Type `ADD` once to authorize folder creation and the import.
5. Review import and XREF warnings, then check in changed drawings through Project Manager.

The folder workflow mirrors Plant's Project Manager dialog: local cache directories are provisioned, folder metadata is initialized with Plant's `FileHelper` (including the active project's DWT template paths), folders are registered with `Project.CreateFolder`, and the project write time is updated. Collaboration projects call `AddProjectItemWithCIP` followed by the folders-only `DownloadHeadRevisionWithCIP` refresh after a successful add; the importer does not run a collaboration refresh before folder creation. If `AddProjectItemWithCIP` fails, the importer follows the dialog's local cleanup sequence (remove the pending project row with tombstones disabled, then delete only a newly created local directory), updates the project write time, and reports cleanup failures. Because the collaboration service may have applied the remote folder operation before returning an error, this cleanup does not guarantee removal of the remote folder; verify and reconcile Project Manager and Forma before retrying. Folder and collaboration failures stop before drawing copy and registration. The sync-aware drawing API also stops on conflicts or out-of-sync results. A copied DWG may remain, and a drawing may already have been registered, if a later import-time failure occurs.

## Update workflow

`P3DUPDATEDRAWINGPROPERTIES` changes properties of drawings that are already registered in the active project. It reads the same CSV format and property-matching rules as the import, but it never copies, registers, or creates folders, and it does not process XREFs. Source DWGs are not needed.

1. Open the project in Plant 3D and run `P3DUPDATEDRAWINGPROPERTIES`.
2. Select the CSV. `DrawingRelativePath` locates each drawing's project folder and file name.
3. Review the preview: drawings to update, rows that already match, and rows skipped because the drawing is not registered in the PnId part (or its DWG is missing on disk).
4. Type `UPDATE` to apply the changes.
5. Check in updated drawings through Project Manager.

Each changed value is written with `SetDrawingPropertyValue`, and each drawing is then accepted with `AcceptDrawingProperties`. The command stops at the first Plant error and reports the drawing where it stopped; drawings processed before it keep their changes.

## Collaboration folder investigation handoff

- The user confirmed that `1234X` was absent from Forma before running the command, then appeared there after `AddProjectItemWithCIP` failed with `Folder 'full name: PID DWG\1234X' already exist`. An earlier cleanup attempt removed the pending PnId row and new cache directory but did not remove the Forma-side folder. A later preservation attempt reintroduced a desynced project, so the user requested restoration of the removal behavior. Do not interpret the service error as proof the folder pre-existed or assume local cleanup removes a server-side partial result.
- Current code follows the manual dialog sequence: no pre-create collaboration refresh; create local directory and PnId row; set folder and project-template paths through `FileHelper`; add with `AddProjectItemWithCIP`; download folders-only state only after a successful add; update the project write time after success or local cleanup. On add failure it uses the dialog's tombstone-disabled removal and deletes only a directory created by this import. This restored behavior has not yet been runtime-tested after the latest code changes.
- The runtime `DownloadHeadRevisionWithCIP` overload has seven parameters, including a final `CancellationToken`; reflection code supplies `CancellationToken.None`.
- The decompiled dialog does not establish why `AddProjectItemWithCIP` creates/retains the Forma-side folder but returns “already exist.” Next investigate the dialog caller's initialization of `FolderLocation`, `Template`, and parent ID, plus the collaboration add implementation and its side effects. Keep tests in a disposable collaborative project; inspect Forma, Project Manager, local cache, and PnId state after each attempt before retrying.
- The last build succeeded with 0 errors and 4 warnings. Plant runtime behavior remains unverified; there is no automated Plant test harness.

## XREF handling

Drawing registration uses the sync-aware `AddPnPDrawingFile` API with its built-in XREF validation disabled, because Plant skips registration entirely when that validation fails. The importer runs `AllXrefsAreLegal` separately and reports flagged drawings as warnings only; an empty add result stops the import. After registration, the importer rewrites absolute XREF paths to relative paths only when:

- the drawing is inside the active project's local folder;
- the XREF target exists; and
- the target is also inside that project folder.

Missing targets and absolute references outside the project folder are left unchanged and reported. The XREF pass edits the registered DWGs in the local project folder; check in those changes through Project Manager. A successful import does not guarantee every XREF resolves, so review warnings and verify references in Plant 3D.

## Diagnostics and manual validation

`P3DPROJECTAPIINFO` reports runtime API information in the command line and does not modify the project. Validate folder creation and synchronization in a disposable collaborative project with backed-up drawings before using it in production. There is no automated test harness for Plant 3D runtime behavior.
