using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Plant3DMetadataAddin
{
    public partial class DrawingPropertyImportCommand
    {
        [CommandMethod("P3DUPDATEDRAWINGPROPERTIES")]
        public void UpdateDrawingProperties()
        {
            var editor = Application.DocumentManager.MdiActiveDocument?.Editor;
            if (editor == null)
            {
                Application.ShowAlertDialog(
                    "Open a drawing in Plant 3D before running P3DUPDATEDRAWINGPROPERTIES.");
                return;
            }

            editor.WriteMessage(
                "\nDrawing-property updater loaded from: {0}",
                typeof(DrawingPropertyImportCommand).Assembly.Location);

            var plantProject = PlantApplication.CurrentProject;
            if (plantProject == null)
            {
                editor.WriteMessage("\nNo active Plant 3D project is open.");
                return;
            }

            var csvPrompt = new PromptStringOptions(
                "\nEnter the full path to the metadata CSV (DWG paths are relative to the active project folder): ")
            {
                AllowSpaces = true
            };
            var csvResult = editor.GetString(csvPrompt);
            if (csvResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nUpdate canceled.");
                return;
            }

            var csvPath = NormalizeEnteredPath(csvResult.StringResult);
            if (!File.Exists(csvPath))
            {
                editor.WriteMessage("\nCSV file does not exist: {0}", csvPath);
                return;
            }

            Project projectPart;
            try
            {
                projectPart = plantProject.ProjectParts[ProjectPartKey];
            }
            catch (System.Exception exception)
            {
                editor.WriteMessage(
                    "\nCould not open project part '{0}': {1}",
                    ProjectPartKey,
                    exception.GetBaseException().Message);
                return;
            }

            if (projectPart == null)
            {
                editor.WriteMessage(
                    "\nProject part '{0}' is not available.",
                    ProjectPartKey);
                return;
            }

            try
            {
                var mappings = ReadPropertyMappings(projectPart);
                var records = ReadDrawingRecords(csvPath, mappings, plantProject.ProjectFolderPath);

                var unregisteredPaths = new List<string>();
                var unchangedPaths = new List<string>();
                var updates = PlanPropertyUpdates(
                    projectPart,
                    records,
                    unregisteredPaths,
                    unchangedPaths);

                editor.WriteMessage(
                    "\nPreflight passed: {0} CSV row(s), {1} property column(s), project part '{2}'.",
                    records.Count,
                    mappings.Count,
                    ProjectPartKey);
                editor.WriteMessage(
                    "\nDrawings to update: {0}. Already matching (no change): {1}. Not registered in this project (will be skipped): {2}.",
                    updates.Count,
                    unchangedPaths.Count,
                    unregisteredPaths.Count);

                foreach (var update in updates.Take(10))
                {
                    editor.WriteMessage(
                        "\n  {0}: {1}",
                        update.Record.RelativePath,
                        string.Join(
                            ", ",
                            update.Changes.Select(change =>
                                change.Mapping.DisplayName + " = '" + change.Value + "'")));
                }
                if (updates.Count > 10)
                {
                    editor.WriteMessage("\n  ... and {0} more drawing(s).", updates.Count - 10);
                }

                if (unregisteredPaths.Count > 0)
                {
                    editor.WriteMessage(
                        "\nSkipped, not registered in project part '{0}' (use P3DIMPORTDRAWINGPROPERTIES to add them):",
                        ProjectPartKey);
                    foreach (var path in unregisteredPaths)
                    {
                        editor.WriteMessage("\n  {0}", path);
                    }
                }

                if (updates.Count == 0)
                {
                    editor.WriteMessage("\nNo project changes are needed.");
                    return;
                }

                var confirmation = editor.GetString(
                    new PromptStringOptions(
                        "\nType UPDATE to set these Plant 3D drawing properties: ")
                    {
                        AllowSpaces = false
                    });
                if (confirmation.Status != PromptStatus.OK ||
                    !string.Equals(
                        confirmation.StringResult,
                        "UPDATE",
                        StringComparison.OrdinalIgnoreCase))
                {
                    editor.WriteMessage("\nUpdate canceled; no project changes were made.");
                    return;
                }

                ApplyPropertyUpdates(projectPart, updates, editor);
            }
            catch (System.Exception exception)
            {
                editor.WriteMessage(
                    "\nUpdate stopped: {0}",
                    exception.GetBaseException().Message);
            }
        }

        private static List<DrawingUpdate> PlanPropertyUpdates(
            Project projectPart,
            List<DrawingRecord> records,
            List<string> unregisteredPaths,
            List<string> unchangedPaths)
        {
            var updates = new List<DrawingUpdate>();
            foreach (var record in records)
            {
                var drawingPath = FindRegisteredDrawingPath(projectPart, record);
                if (drawingPath == null)
                {
                    unregisteredPaths.Add(record.RelativePath);
                    continue;
                }

                var changes = record.PropertyValues
                    .Where(value => !DrawingPropertyHasValue(
                        projectPart,
                        drawingPath,
                        value.Mapping,
                        value.Value))
                    .ToList();
                if (changes.Count == 0)
                {
                    unchangedPaths.Add(record.RelativePath);
                    continue;
                }

                updates.Add(new DrawingUpdate(record, drawingPath, changes));
            }

            return updates;
        }

        // Looks up a drawing that is already registered in the project. Unlike
        // GetProjectDrawingPath, this never falls back to the source path and never
        // copies or registers files.
        private static string FindRegisteredDrawingPath(
            Project projectPart,
            DrawingRecord record)
        {
            var folder = FindProjectFolder(projectPart, record.RelativeFolderPath);
            if (!string.IsNullOrEmpty(record.RelativeFolderPath) && folder == null)
            {
                return null;
            }

            var projectFolder = folder == null
                ? null
                : projectPart.GetFolder(GetProjectFolderGuid(folder));
            if (folder != null && projectFolder == null)
            {
                return null;
            }

            var projectItems = projectFolder == null ? projectPart.ProjectItems : projectFolder.Children;
            foreach (var item in projectItems)
            {
                if (string.Equals(
                        ReadItemProperty(item, "ItemType"),
                        "DwgFile",
                        StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(
                        ReadItemProperty(item, "Name"),
                        record.FileName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    var absoluteFileName = ReadItemProperty(item, "AbsoluteFileName");
                    if (!string.IsNullOrWhiteSpace(absoluteFileName) &&
                        File.Exists(absoluteFileName))
                    {
                        return absoluteFileName;
                    }
                }
            }

            return null;
        }

        private static void ApplyPropertyUpdates(
            Project projectPart,
            List<DrawingUpdate> updates,
            Editor editor)
        {
            var updatedCount = 0;
            foreach (var update in updates)
            {
                try
                {
                    foreach (var change in update.Changes)
                    {
                        if (!projectPart.SetDrawingPropertyValue(
                                update.DrawingPath,
                                change.Mapping.Category,
                                change.Mapping.Name,
                                change.Value) &&
                            !DrawingPropertyHasValue(
                                projectPart,
                                update.DrawingPath,
                                change.Mapping,
                                change.Value))
                        {
                            throw new InvalidOperationException(
                                "Plant 3D rejected property " +
                                change.Mapping.Category + " / " +
                                change.Mapping.DisplayName +
                                " with value '" + change.Value +
                                "' for drawing '" + update.DrawingPath + "'.");
                        }
                    }

                    if (!projectPart.AcceptDrawingProperties(update.DrawingPath))
                    {
                        throw new InvalidOperationException(
                            "Plant 3D did not accept the drawing property changes.");
                    }

                    updatedCount++;
                    editor.WriteMessage(
                        "\nUpdated {0} ({1}/{2}).",
                        update.Record.FileName,
                        updatedCount,
                        updates.Count);
                }
                catch (System.Exception exception)
                {
                    editor.WriteMessage(
                        "\nUpdate stopped at '{0}' after {1} completed drawing(s): {2}",
                        update.Record.RelativePath,
                        updatedCount,
                        exception.GetBaseException().Message);
                    editor.WriteMessage(
                        "\nProperties already set on earlier drawings were accepted. Properties on '{0}' may be partially set; review it in Plant 3D before retrying.",
                        update.Record.RelativePath);
                    return;
                }
            }

            editor.WriteMessage(
                "\nUpdate finished. {0} drawing(s) were updated and their properties accepted.",
                updatedCount);
        }

        private sealed class DrawingUpdate
        {
            public DrawingUpdate(
                DrawingRecord record,
                string drawingPath,
                List<PropertyValue> changes)
            {
                Record = record;
                DrawingPath = drawingPath;
                Changes = changes;
            }

            public DrawingRecord Record { get; }
            public string DrawingPath { get; }
            public List<PropertyValue> Changes { get; }
        }
    }
}
