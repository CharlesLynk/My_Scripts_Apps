# Plant 3D Metadata Add-in Contributor Guide

Read this file and `README.md` before changing this add-in.

## Scope and runtime

- This is a .NET 10 Windows add-in for AutoCAD Plant 3D 2026.
- Plant APIs require the full AutoCAD Plant 3D application, an active drawing, and the intended signed-in project. Accoreconsole is not a supported runtime for these commands.
- `P3DPROJECTAPIINFO` is read-only diagnostics. `P3DIMPORTDRAWINGPROPERTIES` changes the active project and its drawings.
- The import command is PnId-only; access the `PnId` project part directly instead of prompting for a part.

## Import invariants

- Keep strict CSV field-count and property-mapping validation before the user confirms `ADD`.
- `DrawingRelativePath` is canonical. Continue accepting `DrawingFileName` only as a legacy alias; reject a CSV containing both.
- Resolve source DWGs relative to the CSV directory. Destination folder paths are relative to the PnId project part.
- After `ADD`, register each drawing from its intended project-folder location. Copy an external source into the existing destination folder without overwriting; allow an already-present destination DWG when its contents match the source and it is not registered.
- Aggregate missing-source, missing/unregistered destination-folder, registered-drawing, and different-destination-file conflicts before copying DWGs or changing the project.
- Show all missing folder paths before confirmation. One `ADD` authorizes the Project Manager folder workflow and import.
- Match `frmProjectFolderProperties`: validate the folder name and path, create the local cache directory, initialize `PnPProjectFolder` path and template fields with `FileHelper` relative to the project XML directory, call `Project.CreateFolder(null, folder)` at the implicit PnId root or pass the parent folder GUID for child folders, update the project write time, and verify the resolved local path.
- For vault/collaboration projects, match the UI's order: call `Project.CreateFolder`, then `DocumentServerUtils.AddProjectItemWithCIP`; call `DownloadHeadRevisionWithCIP` with folders-only options only after the add succeeds. Do not run that refresh before folder creation or call `DocumentService.CreateFolder` as a replacement.
- If `AddProjectItemWithCIP` fails, mirror the UI cleanup: temporarily disable PnP tombstones, remove the pending project item, restore the tombstone setting, delete only a local directory created by this import, and update the project write time. Report each cleanup failure explicitly. Local cleanup does not guarantee rollback of a folder the collaboration service may have already created remotely; report the partial state and stop before drawing import.
- Treat folder API or collaboration errors as explicit partial failures; stop before copying or registering drawings and report folders/directories that may remain. Do not claim rollback unless all relevant cleanup succeeded.
- The PnId part root is implicit and may not be returned by `FindProjectFolderByPath`. `P3DPROJECTAPIINFO` reports project-part mappings and folder path metadata without changing the project.
- Keep the sync-aware `AddPnPDrawingFile` overload and inspect its sync-conflict and out-of-sync results. Call it with `bValidateXref` false: with validation on, Plant returns an empty result without registering the drawing when its XREFs are illegal (decompiled `AddPnPDrawingFile`). Run `AllXrefsAreLegal` separately for a warning-only report, and treat an empty add result as an error.
- After registration, only rewrite an absolute XREF path when both the drawing and existing XREF target are inside the active project's local folder. Leave missing or outside-project targets unchanged and report them.
- The XREF pass edits registered DWGs; users must review warnings and check in changes through Project Manager.
- Stop on the first import error and report it. Do not silently roll back, retry, or claim that a failed drawing was not registered.

## Collaboration-folder investigation handoff

- Read the local [Project implementation decompilation](./reference/decompiled/PlantProjectManager.Project.decompiled.cs.txt) and [folder dialog decompilation](./reference/decompiled/PlantProjectManagerUI.FolderPropertiesDialog.decompiled.cs.txt) before changing folder creation. They are plain-text, user-provided references, not source files or authoritative API contracts; do not edit or redistribute them.
- The user confirmed `1234X` was absent from Forma before the command, then appeared after `AddProjectItemWithCIP` failed with “already exist.” A prior cleanup run removed the local row/cache but did not remove the cloud-side folder. A later attempt to preserve the local row caused a desynced project; the user explicitly requested the removal behavior. Do not assume the folder pre-existed, do not preserve a failed pending row, and do not claim local cleanup rolls back the cloud operation.
- The dialog sequence is local directory creation, PnP folder construction (including template paths), `Project.CreateFolder`, `AddProjectItemWithCIP`, then folders-only download only on successful add. The dialog also calls `SetWriteTimeToCurrentUtc` after successful handling or failed-add cleanup. The importer now mirrors this order and uses the UI cleanup on add failure.
- The available dialog decompilation does not include its caller's initialization of `FolderLocation`, `Template`, or parent ID. If the add still fails, inspect/decompile `PrjMgrBaseUsrCtrl.newFolder_MenuItem_Click` and the runtime `AddProjectItemWithCIP` side effects. Keep testing in a disposable collaborative project and inspect Forma, Project Manager, local cache, and PnId state after each attempt.
- The latest build succeeded with 0 errors and 4 warnings; the latest folder workflow has not yet been runtime-tested. There is no automated Plant runtime test harness.

## Build and validation

- Build with `dotnet build .\Plant3DMetadataAddin.csproj -c Debug`.
- The expected output is `bin\Debug\net10.0-windows\Plant3DProjectApiProbe.dll`.
- Plant 3D locks loaded add-in DLLs. Ask the user to close Plant 3D and wait for confirmation before rebuilding; do not build to an alternate output path.
- No automated Plant runtime test is available. Use a disposable project and backed-up drawings for manual validation.
- `bin/` and `obj/` are generated artifacts. Do not edit or commit their contents, including captured API output.
