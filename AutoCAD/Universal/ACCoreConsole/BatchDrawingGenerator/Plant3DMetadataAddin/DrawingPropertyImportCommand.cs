using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.DataObjects.Synchronization;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;

namespace Plant3DMetadataAddin
{
    public class DrawingPropertyImportCommand
    {
        private const string DrawingRelativePathHeader = "DrawingRelativePath";
        private const string LegacyDrawingFileNameHeader = "DrawingFileName";
        private const string ProjectPartKey = "PnId";

        [CommandMethod("P3DIMPORTDRAWINGPROPERTIES")]
        public void ImportDrawingProperties()
        {
            var editor = Application.DocumentManager.MdiActiveDocument?.Editor;
            if (editor == null)
            {
                Application.ShowAlertDialog(
                    "Open a drawing in Plant 3D before running P3DIMPORTDRAWINGPROPERTIES.");
                return;
            }

            editor.WriteMessage(
                "\nDrawing-property importer loaded from: {0}",
                typeof(DrawingPropertyImportCommand).Assembly.Location);

            var plantProject = PlantApplication.CurrentProject;
            if (plantProject == null)
            {
                editor.WriteMessage("\nNo active Plant 3D project is open.");
                return;
            }

            var csvPrompt = new PromptStringOptions(
                "\nEnter the full path to the metadata CSV (DWG paths are relative to its folder): ")
            {
                AllowSpaces = true
            };
            var csvResult = editor.GetString(csvPrompt);
            if (csvResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nImport canceled.");
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

            var projectRootPath = plantProject.ProjectFolderPath;
            if (string.IsNullOrWhiteSpace(projectRootPath) ||
                !Directory.Exists(projectRootPath))
            {
                editor.WriteMessage(
                    "\nThe active Plant 3D project's local folder is unavailable; no drawings were added.");
                return;
            }

            try
            {
                var mappings = ReadPropertyMappings(projectPart);
                var records = ReadDrawingRecords(csvPath, mappings);
                var folderPathsToCreate = ValidateRecords(
                    projectPart,
                    records,
                    csvPath,
                    projectRootPath);
                editor.WriteMessage(
                    "\nPreflight passed: {0} drawing(s), {1} property column(s), project part '{2}'.",
                    records.Count,
                    mappings.Count,
                    ProjectPartKey);
                editor.WriteMessage(
                    "\nFile placement: {0} source DWG(s) will be copied to their project folders; existing identical destination files will be registered in place.",
                    records.Count(record => record.CopySourceToDestination));
                editor.WriteMessage(
                    "\nSource DWGs are resolved relative to the CSV folder: {0}",
                    Path.GetDirectoryName(csvPath));
                editor.WriteMessage(
                    "\nPlant 3D will register each DWG from its intended project-folder location. Review any XREF validation warnings after import.");
                if (folderPathsToCreate.Count > 0)
                {
                    editor.WriteMessage(
                        "\nMissing project folder(s) to create and register through the Project Manager workflow ({0}):",
                        folderPathsToCreate.Count);
                    foreach (var folderPath in folderPathsToCreate)
                    {
                        editor.WriteMessage("\n  {0}", folderPath);
                    }
                    editor.WriteMessage(
                        "\nOne ADD command authorizes creating local cache directories, registering the folders with Plant, and synchronizing them before drawing import.");
                }
                foreach (var record in records.Take(10))
                {
                    editor.WriteMessage(
                        "\n  {0} -> {1}",
                        record.RelativePath,
                        string.IsNullOrEmpty(record.RelativeFolderPath)
                            ? "<project root>"
                            : record.RelativeFolderPath);
                }
                if (records.Count > 10)
                {
                    editor.WriteMessage("\n  ... and {0} more drawing(s).", records.Count - 10);
                }

                var confirmation = editor.GetString(
                    new PromptStringOptions(
                        "\nType ADD to create missing folders, place drawings, register them, and set their Plant 3D properties: ")
                    {
                        AllowSpaces = false
                    });
                if (confirmation.Status != PromptStatus.OK ||
                    !string.Equals(
                        confirmation.StringResult,
                        "ADD",
                        StringComparison.OrdinalIgnoreCase))
                {
                    editor.WriteMessage("\nImport canceled; no project changes were made.");
                    return;
                }

                var createdFolderPaths = new List<string>();
                try
                {
                    CreateMissingProjectFolders(
                        plantProject,
                        projectPart,
                        projectRootPath,
                        folderPathsToCreate,
                        createdFolderPaths,
                        editor);
                }
                catch (System.Exception exception)
                {
                    editor.WriteMessage(
                        "\nProject folder setup stopped before drawing copy or registration: {0}",
                        exception.Message);
                    editor.WriteMessage(
                        "\nNo drawings were copied or registered. Folder rows or local cache directories created before the failure may remain; review the listed paths and Project Manager before retrying.");
                    foreach (var createdFolderPath in createdFolderPaths)
                    {
                        editor.WriteMessage("\n  {0}", createdFolderPath);
                    }
                    return;
                }

                ImportRecords(
                    projectPart,
                    records,
                    projectRootPath,
                    editor);
            }
            catch (System.Exception exception)
            {
                editor.WriteMessage(
                    "\nImport preflight failed; no drawings were added: {0}",
                    exception.GetBaseException().Message);
            }
        }

        private static List<PropertyMapping> ReadPropertyMappings(Project projectPart)
        {
            var metadata = new List<DrawingProperty>();
            foreach (var category in projectPart.GetCategoryNames(CategoryType.eDrawing))
            {
                metadata.AddRange(projectPart.GetDrawingPropertyMetadata(category));
            }

            if (metadata.Count == 0)
            {
                throw new InvalidOperationException(
                    "The selected project part has no drawing-property metadata.");
            }

            return metadata
                .Select(property => new PropertyMapping(
                    property.Category,
                    property.Name,
                    property.DisplayName))
                .ToList();
        }

        private static List<DrawingRecord> ReadDrawingRecords(
            string csvPath,
            List<PropertyMapping> metadata)
        {
            string[] headers;
            var rows = new List<string[]>();
            using (var parser = new TextFieldParser(csvPath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                try
                {
                    headers = parser.ReadFields();
                    if (headers == null || headers.Length == 0)
                    {
                        throw new InvalidOperationException("CSV has no header row.");
                    }

                    var seenHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var header in headers)
                    {
                        if (string.IsNullOrWhiteSpace(header))
                        {
                            throw new InvalidOperationException("CSV headers cannot be empty.");
                        }
                        if (!seenHeaders.Add(header))
                        {
                            throw new InvalidOperationException(
                                "CSV contains a duplicate header: " + header);
                        }
                    }

                    while (!parser.EndOfData)
                    {
                        var fields = parser.ReadFields();
                        if (fields.Length != headers.Length)
                        {
                            throw new InvalidOperationException(
                                "CSV record " + parser.LineNumber +
                                " has " + fields.Length +
                                " field(s); expected " + headers.Length + ".");
                        }
                        rows.Add(fields);
                    }
                }
                catch (MalformedLineException exception)
                {
                    throw new InvalidOperationException(
                        "CSV is malformed near line " + parser.ErrorLineNumber +
                        ": " + exception.Message,
                        exception);
                }
            }

            var drawingPathIndex = Array.FindIndex(
                headers,
                header => string.Equals(
                    header,
                    DrawingRelativePathHeader,
                    StringComparison.OrdinalIgnoreCase));
            var legacyDrawingPathIndex = Array.FindIndex(
                headers,
                header => string.Equals(
                    header,
                    LegacyDrawingFileNameHeader,
                    StringComparison.OrdinalIgnoreCase));
            if (drawingPathIndex >= 0 && legacyDrawingPathIndex >= 0)
            {
                throw new InvalidOperationException(
                    "CSV cannot contain both DrawingRelativePath and the legacy DrawingFileName column.");
            }
            if (drawingPathIndex < 0)
            {
                drawingPathIndex = legacyDrawingPathIndex;
            }
            if (drawingPathIndex < 0)
            {
                throw new InvalidOperationException(
                    "CSV must contain a DrawingRelativePath column.");
            }

            var mappingsByColumn = new List<PropertyMapping>();
            for (var index = 0; index < headers.Length; index++)
            {
                if (index == drawingPathIndex)
                {
                    continue;
                }

                var mapping = MatchProperty(headers[index], metadata);
                if (mapping == null)
                {
                    throw new InvalidOperationException(
                        "CSV column '" + headers[index] +
                        "' does not match a drawing property in the selected project part.");
                }
                mappingsByColumn.Add(mapping);
            }

            if (mappingsByColumn.Count == 0)
            {
                throw new InvalidOperationException(
                    "CSV must contain at least one drawing-property column.");
            }

            var seenProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var mapping in mappingsByColumn)
            {
                if (!seenProperties.Add(mapping.Category + "\0" + mapping.Name))
                {
                    throw new InvalidOperationException(
                        "CSV maps more than one column to " +
                        mapping.Category + " / " + mapping.DisplayName + ".");
                }
            }

            var records = new List<DrawingRecord>();
            var sourceDirectory = Path.GetDirectoryName(Path.GetFullPath(csvPath));
            foreach (var fields in rows)
            {
                var drawingPath = fields[drawingPathIndex].Trim();
                if (string.IsNullOrWhiteSpace(drawingPath))
                {
                    throw new InvalidOperationException(
                        "Every CSV record must have a DrawingRelativePath.");
                }

                var pathSegments = NormalizeDrawingPath(drawingPath);
                var fileName = pathSegments[pathSegments.Length - 1];
                if (string.IsNullOrEmpty(Path.GetExtension(fileName)))
                {
                    fileName += ".dwg";
                    pathSegments[pathSegments.Length - 1] = fileName;
                }
                else if (!string.Equals(
                             Path.GetExtension(fileName),
                             ".dwg",
                             StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "DrawingRelativePath must end in .dwg or have no extension: " +
                        drawingPath);
                }

                var relativePath = Path.Combine(pathSegments);
                var relativeFolderPath = pathSegments.Length > 1
                    ? Path.Combine(pathSegments.Take(pathSegments.Length - 1).ToArray())
                    : string.Empty;
                var propertyValues = new List<PropertyValue>();
                var propertyColumnIndex = 0;
                for (var index = 0; index < headers.Length; index++)
                {
                    if (index == drawingPathIndex)
                    {
                        continue;
                    }
                    propertyValues.Add(new PropertyValue(
                        mappingsByColumn[propertyColumnIndex++],
                        fields[index]));
                }

                records.Add(new DrawingRecord(
                    fileName,
                    relativePath,
                    relativeFolderPath,
                    Path.Combine(sourceDirectory, relativePath),
                    propertyValues));
            }

            if (records.Count == 0)
            {
                throw new InvalidOperationException("CSV contains no drawing rows.");
            }

            return records;
        }

        private static string[] NormalizeDrawingPath(string drawingPath)
        {
            var normalizedPath = drawingPath.Replace('/', Path.DirectorySeparatorChar)
                .Replace('\\', Path.DirectorySeparatorChar);
            if (Path.IsPathRooted(normalizedPath) ||
                (normalizedPath.Length >= 2 &&
                 char.IsLetter(normalizedPath[0]) &&
                 normalizedPath[1] == ':'))
            {
                throw new InvalidOperationException(
                    "DrawingRelativePath must be relative to the CSV directory: " +
                    drawingPath);
            }

            var segments = normalizedPath.Split(
                new[] { Path.DirectorySeparatorChar },
                StringSplitOptions.None);
            foreach (var segment in segments)
            {
                if (string.IsNullOrWhiteSpace(segment) ||
                    segment == "." ||
                    segment == ".." ||
                    segment.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                    segment.EndsWith(".", StringComparison.Ordinal) ||
                    segment.EndsWith(" ", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        "DrawingRelativePath contains an invalid path segment: " +
                        drawingPath);
                }
            }

            return segments;
        }

        private static PropertyMapping MatchProperty(
            string header,
            List<PropertyMapping> metadata)
        {
            var normalizedHeader = NormalizePropertyName(header);
            var matches = metadata
                .Where(property =>
                    NormalizePropertyName(property.Name) == normalizedHeader ||
                    NormalizePropertyName(property.DisplayName) == normalizedHeader)
                .GroupBy(property => property.Category + "\0" + property.Name)
                .Select(group => group.First())
                .ToList();

            if (matches.Count > 1)
            {
                throw new InvalidOperationException(
                    "CSV column '" + header +
                    "' matches drawing properties in more than one category.");
            }

            return matches.SingleOrDefault();
        }

        private static string NormalizePropertyName(string value)
        {
            return new string(
                value.Where(char.IsLetterOrDigit)
                    .Select(char.ToUpperInvariant)
                    .ToArray());
        }

        private static List<string> ValidateRecords(
            Project projectPart,
            List<DrawingRecord> records,
            string csvPath,
            string projectRootPath)
        {
            var preflightIssues = new List<string>();
            var missingFolders = FindMissingProjectFolders(
                projectPart,
                records);
            var missingFolderSet = new HashSet<string>(
                missingFolders,
                StringComparer.OrdinalIgnoreCase);

            foreach (var duplicatePath in records
                         .GroupBy(
                             record => record.RelativePath,
                             StringComparer.OrdinalIgnoreCase)
                         .Where(group => group.Count() > 1))
            {
                preflightIssues.Add(
                    "CSV repeats DrawingRelativePath '" + duplicatePath.Key +
                    "' (" + duplicatePath.Count() + " rows).");
            }

            foreach (var record in records)
            {
                if (!File.Exists(record.SourceFilePath))
                {
                    preflightIssues.Add(
                        "Source DWG listed in the CSV was not found: " +
                        record.RelativePath + " (" + record.SourceFilePath + ").");
                }

                var existingFolder = FindProjectFolder(
                    projectPart,
                    record.RelativeFolderPath);
                if (!string.IsNullOrEmpty(record.RelativeFolderPath) &&
                    existingFolder == null)
                {
                    if (!missingFolderSet.Contains(record.RelativeFolderPath))
                    {
                        preflightIssues.Add(
                            "Could not resolve destination folder in the selected Plant project part: " +
                            record.RelativeFolderPath);
                        continue;
                    }
                }

                var existingItems = string.IsNullOrEmpty(record.RelativeFolderPath)
                    ? projectPart.ProjectItems
                    : existingFolder?.Children;
                if (existingItems != null && existingItems.Any(item =>
                        string.Equals(
                            ReadItemProperty(item, "ItemType"),
                            "DwgFile",
                            StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(
                            ReadItemProperty(item, "Name"),
                            record.FileName,
                            StringComparison.OrdinalIgnoreCase)))
                {
                    preflightIssues.Add(
                        "Drawing is already registered in the selected project folder: " +
                        record.RelativePath);
                }

                var destinationDirectory = GetDestinationFolderPath(
                    projectPart,
                    projectRootPath,
                    record.RelativeFolderPath);
                if (!Directory.Exists(destinationDirectory) &&
                    !missingFolderSet.Contains(record.RelativeFolderPath))
                {
                    preflightIssues.Add(
                        "Project destination folder is not available locally for '" +
                        record.RelativePath + "': " + destinationDirectory);
                    continue;
                }

                var destinationFilePath = Path.GetFullPath(
                    Path.Combine(destinationDirectory, record.FileName));
                if (!IsPathWithinRoot(destinationFilePath, projectRootPath))
                {
                    preflightIssues.Add(
                        "Resolved destination is outside the active project's local folder: " +
                        destinationFilePath);
                    continue;
                }

                record.DestinationFilePath = destinationFilePath;
                if (!File.Exists(record.SourceFilePath))
                {
                    continue;
                }

                if (PathsAreEqual(record.SourceFilePath, destinationFilePath))
                {
                    record.FilePath = destinationFilePath;
                    continue;
                }

                if (Directory.Exists(destinationFilePath))
                {
                    preflightIssues.Add(
                        "A directory already occupies the destination DWG path: " +
                        destinationFilePath);
                    continue;
                }

                if (File.Exists(destinationFilePath))
                {
                    try
                    {
                        if (!FilesHaveSameContents(
                                record.SourceFilePath,
                                destinationFilePath))
                        {
                            preflightIssues.Add(
                                "A different DWG already occupies the destination and will not be overwritten: " +
                                destinationFilePath);
                            continue;
                        }
                    }
                    catch (IOException exception)
                    {
                        preflightIssues.Add(
                            "Could not compare the source and destination DWGs for '" +
                            record.RelativePath + "': " +
                            exception.GetBaseException().Message);
                        continue;
                    }
                    catch (UnauthorizedAccessException exception)
                    {
                        preflightIssues.Add(
                            "Could not compare the source and destination DWGs for '" +
                            record.RelativePath + "': " +
                            exception.GetBaseException().Message);
                        continue;
                    }

                    record.FilePath = destinationFilePath;
                    continue;
                }

                record.FilePath = destinationFilePath;
                record.CopySourceToDestination = true;
            }

            if (string.IsNullOrWhiteSpace(Path.GetDirectoryName(csvPath)))
            {
                preflightIssues.Add(
                    "Could not determine the CSV directory for source drawings.");
            }

            if (preflightIssues.Count > 0)
            {
                throw new InvalidOperationException(
                    "CSV preflight found " + preflightIssues.Count +
                    " issue(s). No DWG files were copied and no drawings were added:\n  " +
                    string.Join("\n  ", preflightIssues));
            }

            return missingFolders;
        }

        private static List<string> FindMissingProjectFolders(
            Project projectPart,
            List<DrawingRecord> records)
        {
            var missingFolderPaths = new List<string>();
            var seenFolderPaths = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (var folderPath in records
                         .Select(record => record.RelativeFolderPath)
                         .Where(path => !string.IsNullOrEmpty(path))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                string parentGuid = null;
                var currentRelativePath = string.Empty;
                var hasMissingAncestor = false;
                foreach (var segment in folderPath.Split(Path.DirectorySeparatorChar))
                {
                    currentRelativePath = Path.Combine(currentRelativePath, segment);
                    if (!hasMissingAncestor)
                    {
                        var projectFolder = projectPart.GetChildFolder(
                            parentGuid,
                            segment);
                        if (projectFolder != null)
                        {
                            parentGuid = GetProjectFolderGuid(projectFolder);
                            continue;
                        }

                        hasMissingAncestor = true;
                    }

                    if (seenFolderPaths.Add(currentRelativePath))
                    {
                        missingFolderPaths.Add(currentRelativePath);
                    }
                }
            }

            return missingFolderPaths;
        }

        private static string GetDestinationFolderPath(
            Project projectPart,
            string projectRootPath,
            string relativeFolderPath)
        {
            string destinationFolderPath;
            if (string.IsNullOrEmpty(relativeFolderPath))
            {
                destinationFolderPath = GetProjectPartRootFolderPath(
                    projectPart,
                    projectRootPath);
            }
            else
            {
                destinationFolderPath = GetProjectPartRootFolderPath(
                    projectPart,
                    projectRootPath);
                string parentGuid = null;
                var hasMissingAncestor = false;
                foreach (var segment in relativeFolderPath.Split(
                             Path.DirectorySeparatorChar))
                {
                    var projectFolder = hasMissingAncestor
                        ? null
                        : projectPart.GetChildFolder(parentGuid, segment);
                    if (projectFolder == null)
                    {
                        hasMissingAncestor = true;
                        destinationFolderPath = Path.Combine(
                            destinationFolderPath,
                            segment);
                    }
                    else
                    {
                        destinationFolderPath = projectPart.GetLocalResolvedFolderPath(
                            projectFolder);
                        parentGuid = GetProjectFolderGuid(projectFolder);
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(destinationFolderPath))
            {
                throw new InvalidOperationException(
                    "Plant 3D did not provide a local path for the destination folder.");
            }

            destinationFolderPath = Path.GetFullPath(destinationFolderPath);
            if (!IsPathWithinRoot(destinationFolderPath, projectRootPath))
            {
                throw new InvalidOperationException(
                    "Resolved destination folder is outside the active project's local folder: " +
                    destinationFolderPath);
            }

            return destinationFolderPath;
        }

        private static void CreateMissingProjectFolders(
            PlantProject plantProject,
            Project projectPart,
            string projectRootPath,
            List<string> folderPathsToCreate,
            List<string> createdFolderPaths,
            Editor editor)
        {
            if (folderPathsToCreate.Count == 0)
            {
                return;
            }

            var documentService = plantProject.DocumentService;
            var documentServerUtils = FindDocumentServerUtilsType();
            var isVaultBased = InvokeBooleanUtilityMethod(
                documentServerUtils,
                "IsVaultBasedProject",
                plantProject);
            if (isVaultBased &&
                !InvokeBooleanUtilityMethod(
                    documentServerUtils,
                    "LogInRequired",
                    documentService))
            {
                throw new InvalidOperationException(
                    "Plant's document service is not ready for the collaboration folder operation. Sign in to the project and retry.");
            }

            var projectXmlDirectory = Path.GetDirectoryName(
                projectPart.ProjectXmlFile);
            if (string.IsNullOrWhiteSpace(projectXmlDirectory))
            {
                throw new InvalidOperationException(
                    "Could not resolve the project XML directory needed to initialize Plant folder metadata.");
            }

            foreach (var folderPath in folderPathsToCreate)
            {
                string parentGuid = null;
                var currentLocalPath = GetProjectPartRootFolderPath(
                    projectPart,
                    projectRootPath);
                var currentRelativePath = string.Empty;
                foreach (var segment in folderPath.Split(
                             Path.DirectorySeparatorChar))
                {
                    currentRelativePath = Path.Combine(
                        currentRelativePath,
                        segment);
                    var existingFolder = projectPart.GetChildFolder(
                        parentGuid,
                        segment);
                    if (existingFolder != null)
                    {
                        currentLocalPath = projectPart.GetLocalResolvedFolderPath(
                            existingFolder);
                        if (string.IsNullOrWhiteSpace(currentLocalPath) ||
                            !IsPathWithinRoot(currentLocalPath, projectRootPath) ||
                            !Directory.Exists(currentLocalPath))
                        {
                            throw new InvalidOperationException(
                                "Existing Plant folder '" + currentRelativePath +
                                "' is not available at a valid local path.");
                        }

                        parentGuid = GetProjectFolderGuid(existingFolder);
                        continue;
                    }

                    var absoluteFolderPath = Path.GetFullPath(
                        Path.Combine(currentLocalPath, segment));
                    if (!IsPathWithinRoot(
                            absoluteFolderPath,
                            projectRootPath))
                    {
                        throw new InvalidOperationException(
                            "Folder path resolves outside the active project's local folder: " +
                            absoluteFolderPath);
                    }

                    if (FileHelper.IsValidPathname(segment) !=
                        FileHelper.ErrorStatus.eOk)
                    {
                        throw new InvalidOperationException(
                            "Plant Project Manager rejected folder name '" +
                            segment + "'.");
                    }

                    if (projectPart.FindProjectFolderByPath(
                            absoluteFolderPath) != null)
                    {
                        throw new InvalidOperationException(
                            "Plant Project Manager already has a folder at the destination path, but it is not registered under the expected parent: " +
                            absoluteFolderPath);
                    }

                    var localDirectoryCreated = !Directory.Exists(
                        absoluteFolderPath);
                    if (localDirectoryCreated)
                    {
                        Directory.CreateDirectory(absoluteFolderPath);
                        if (!createdFolderPaths.Contains(
                                currentRelativePath,
                                StringComparer.OrdinalIgnoreCase))
                        {
                            createdFolderPaths.Add(currentRelativePath);
                        }
                    }

                    var fileHelper = new FileHelper(
                        projectXmlDirectory,
                        absoluteFolderPath);
                    var templatePath = projectPart.DwtTemplateFile ??
                        string.Empty;
                    var templateFileHelper = new FileHelper(
                        projectXmlDirectory,
                        templatePath);
                    var folder = new PnPProjectFolder
                    {
                        Name = segment,
                        AbsoluteFolderPath = fileHelper.FileName,
                        RelativeFolderPath = fileHelper.RelativeFileName,
                        UniversalFileName = fileHelper.UNCFileName,
                        TemplateFile = templatePath,
                        TemplateRelativePath =
                            templateFileHelper.RelativeFileName,
                        TemplateUniversalPath =
                            templateFileHelper.UNCFileName,
                        PromptForTemplate = "0"
                    };
                    var folderGuid = projectPart.CreateFolder(
                        parentGuid,
                        folder);
                    if (string.IsNullOrWhiteSpace(folderGuid))
                    {
                        throw new InvalidOperationException(
                            "Plant's project part did not create folder '" +
                            currentRelativePath + "'.");
                    }

                    var registeredFolder = projectPart.GetFolder(folderGuid);
                    if (registeredFolder == null)
                    {
                        throw new InvalidOperationException(
                            "Plant returned folder GUID '" + folderGuid +
                            "' but the folder could not be resolved for '" +
                            currentRelativePath + "'.");
                    }
                    parentGuid = GetProjectFolderGuid(registeredFolder);
                    if (!createdFolderPaths.Contains(
                            currentRelativePath,
                            StringComparer.OrdinalIgnoreCase))
                    {
                        createdFolderPaths.Add(currentRelativePath);
                    }

                    var resolvedFolderPath = projectPart.GetLocalResolvedFolderPath(
                        registeredFolder);
                    if (string.IsNullOrWhiteSpace(resolvedFolderPath) ||
                        !PathsAreEqual(resolvedFolderPath, absoluteFolderPath) ||
                        !Directory.Exists(resolvedFolderPath))
                    {
                        throw new InvalidOperationException(
                            "Plant registered folder '" + currentRelativePath +
                            "' but its local path did not resolve as expected. Resolved: " +
                            (resolvedFolderPath ?? "<null>") +
                            "; expected: " + absoluteFolderPath);
                    }

                    if (isVaultBased)
                    {
                        try
                        {
                            SynchronizeNewProjectFolder(
                                documentServerUtils,
                                documentService,
                                projectPart,
                                folderGuid,
                                currentRelativePath);
                        }
                        catch (CollaborationFolderAddException exception)
                        {
                            var cleanupIssues = RemoveUnpublishedProjectFolder(
                                projectPart,
                                folderGuid,
                                absoluteFolderPath,
                                localDirectoryCreated);
                            try
                            {
                                projectPart.SetWriteTimeToCurrentUtc();
                            }
                            catch (System.Exception timestampException)
                            {
                                cleanupIssues.Add(
                                    "could not update the project write time: " +
                                    timestampException.GetBaseException().Message);
                            }

                            if (cleanupIssues.Count == 0)
                            {
                                createdFolderPaths.RemoveAll(path =>
                                    string.Equals(
                                        path,
                                        currentRelativePath,
                                        StringComparison.OrdinalIgnoreCase));
                            }

                            var cleanupSummary = cleanupIssues.Count == 0
                                ? "The pending Plant folder entry was removed without creating a tombstone" +
                                  (localDirectoryCreated
                                      ? " and the newly created local directory was deleted."
                                      : ".")
                                : "Cleanup was incomplete: " +
                                  string.Join("; ", cleanupIssues);
                            throw new InvalidOperationException(
                                "Plant collaboration rejected folder '" +
                                currentRelativePath + "': " +
                                exception.Message +
                                " " + cleanupSummary,
                                exception);
                        }
                    }

                    projectPart.SetWriteTimeToCurrentUtc();
                    currentLocalPath = resolvedFolderPath;
                    editor.WriteMessage(
                        "\nCreated and registered Plant project folder: {0}",
                        currentRelativePath);
                }
            }

        }

        private static void SynchronizeNewProjectFolder(
            Type documentServerUtils,
            DocumentServer documentService,
            Project projectPart,
            string folderGuid,
            string relativeFolderPath)
        {
            var projectItem = projectPart.GetProjectItem(folderGuid);
            if (projectItem == null)
            {
                throw new CollaborationFolderAddException(
                    "Plant registered folder '" + relativeFolderPath +
                    "' but could not resolve its project item for collaboration sync.",
                    null);
            }

            var addArguments = new object[]
            {
                documentService,
                projectItem,
                string.Empty,
                false,
                null,
                null,
                null
            };
            try
            {
                var addMethod = FindDocumentServerUtilityMethod(
                    documentServerUtils,
                    "AddProjectItemWithCIP",
                    parameters =>
                        parameters.Length == 7 &&
                        parameters[0].ParameterType.IsInstanceOfType(documentService) &&
                        parameters[1].ParameterType.IsInstanceOfType(projectItem) &&
                        parameters[2].ParameterType == typeof(string) &&
                        parameters[3].ParameterType == typeof(bool) &&
                        IsSyncConflictsOutput(parameters[6]));
                InvokeDocumentServerUtility(addMethod, addArguments);
            }
            catch (System.Exception exception)
            {
                throw new CollaborationFolderAddException(
                    "AddProjectItemWithCIP failed for folder '" +
                    relativeFolderPath + "': " + exception.Message,
                    exception);
            }
            if (HasSyncResults(addArguments[6]))
            {
                throw new CollaborationFolderAddException(
                    "Plant reported collaboration sync conflicts while adding folder '" +
                    relativeFolderPath + "'.",
                    null);
            }

            RefreshCollaborativeProjectFolders(
                documentServerUtils,
                documentService,
                projectPart,
                "after adding '" + relativeFolderPath + "'");
        }

        private static void RefreshCollaborativeProjectFolders(
            Type documentServerUtils,
            DocumentServer documentService,
            Project projectPart,
            string operationDescription)
        {
            var downloadMethod = FindDocumentServerUtilityMethod(
                documentServerUtils,
                "DownloadHeadRevisionWithCIP",
                parameters =>
                    parameters.Length == 7 &&
                    parameters[0].ParameterType.IsInstanceOfType(documentService) &&
                    parameters[1].ParameterType.IsInstanceOfType(projectPart) &&
                    parameters[2].ParameterType == typeof(bool) &&
                    parameters[3].ParameterType == typeof(bool) &&
                    IsSyncConflictsOutput(parameters[5]) &&
                    parameters[6].ParameterType ==
                        typeof(System.Threading.CancellationToken));
            var downloadArguments = new object[]
            {
                documentService,
                projectPart,
                true,
                false,
                null,
                null,
                System.Threading.CancellationToken.None
            };
            InvokeDocumentServerUtility(downloadMethod, downloadArguments);
            if (HasSyncResults(downloadArguments[5]))
            {
                throw new InvalidOperationException(
                    "Plant reported collaboration sync conflicts while refreshing folders " +
                    operationDescription + ".");
            }
        }

        private static List<string> RemoveUnpublishedProjectFolder(
            Project projectPart,
            string folderGuid,
            string localFolderPath,
            bool deleteLocalDirectory)
        {
            var cleanupIssues = new List<string>();
            PnPDatabase database = null;
            var originalCreatingTombstones = false;
            var tombstoneSettingChanged = false;
            try
            {
                var dataLinksManager = projectPart.DataLinksManager;
                database = dataLinksManager?.GetPnPDatabase();
                if (database != null)
                {
                    originalCreatingTombstones = database.CreatingTombstones;
                    database.CreatingTombstones = false;
                    tombstoneSettingChanged = true;
                }
                else
                {
                    cleanupIssues.Add(
                        "the PnP database was unavailable to disable tombstone creation");
                }

                projectPart.RemoveProjectItem(folderGuid);
            }
            catch (System.Exception exception)
            {
                cleanupIssues.Add(
                    "could not remove pending project item: " +
                    exception.GetBaseException().Message);
            }
            finally
            {
                if (database != null && tombstoneSettingChanged)
                {
                    try
                    {
                        database.CreatingTombstones = originalCreatingTombstones;
                    }
                    catch (System.Exception exception)
                    {
                        cleanupIssues.Add(
                            "could not restore the tombstone setting: " +
                            exception.GetBaseException().Message);
                    }
                }
            }

            try
            {
                if (projectPart.GetProjectItem(folderGuid) != null)
                {
                    cleanupIssues.Add(
                        "the project item is still present after RemoveProjectItem");
                }
            }
            catch (System.Exception exception)
            {
                cleanupIssues.Add(
                    "could not verify project-item removal: " +
                    exception.GetBaseException().Message);
            }

            if (deleteLocalDirectory && Directory.Exists(localFolderPath))
            {
                try
                {
                    Directory.Delete(localFolderPath, true);
                }
                catch (IOException exception)
                {
                    cleanupIssues.Add(
                        "could not delete the new local cache directory: " +
                        exception.GetBaseException().Message);
                }
                catch (UnauthorizedAccessException exception)
                {
                    cleanupIssues.Add(
                        "could not delete the new local cache directory: " +
                        exception.GetBaseException().Message);
                }
            }

            return cleanupIssues;
        }

        private static Type FindDocumentServerUtilsType()
        {
            var utilityTypes = new List<Type>();
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                var knownType = assembly.GetType(
                    "Autodesk.ProcessPower.ProjectManagerUI.DocumentServerUtils",
                    false);
                if (knownType != null)
                {
                    utilityTypes.Add(knownType);
                }
            }

            if (utilityTypes.Count == 0)
            {
                foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                {
                    Type[] types;
                    try
                    {
                        types = assembly.GetTypes();
                    }
                    catch (ReflectionTypeLoadException exception)
                    {
                        types = exception.Types
                            .Where(type => type != null)
                            .ToArray();
                    }

                    utilityTypes.AddRange(types.Where(type =>
                        string.Equals(
                            type.Name,
                            "DocumentServerUtils",
                            StringComparison.Ordinal)));
                }
            }

            var distinctTypes = utilityTypes
                .Distinct()
                .ToList();
            if (distinctTypes.Count != 1)
            {
                throw new InvalidOperationException(
                    "Expected one loaded Plant DocumentServerUtils type for collaboration folder synchronization; found " +
                    distinctTypes.Count + ".");
            }

            return distinctTypes[0];
        }

        private static bool InvokeBooleanUtilityMethod(
            Type utilityType,
            string methodName,
            object argument)
        {
            var method = FindDocumentServerUtilityMethod(
                utilityType,
                methodName,
                parameters =>
                    parameters.Length == 1 &&
                    parameters[0].ParameterType.IsInstanceOfType(argument));
            if (method.ReturnType != typeof(bool))
            {
                throw new InvalidOperationException(
                    "Plant utility method '" + methodName +
                    "' does not return the expected Boolean result.");
            }

            var result = InvokeDocumentServerUtility(
                method,
                new[] { argument });
            return result is bool booleanResult
                ? booleanResult
                : throw new InvalidOperationException(
                    "Plant utility method '" + methodName +
                    "' returned an unexpected result.");
        }

        private static MethodInfo FindDocumentServerUtilityMethod(
            Type utilityType,
            string methodName,
            Func<ParameterInfo[], bool> parameterMatch)
        {
            var namedMethods = utilityType.GetMethods(
                    BindingFlags.Public |
                    BindingFlags.NonPublic |
                    BindingFlags.Static)
                .Where(method =>
                    string.Equals(
                        method.Name,
                        methodName,
                        StringComparison.Ordinal))
                .ToList();
            var methods = namedMethods
                .Where(method => parameterMatch(method.GetParameters()))
                .ToList();
            if (methods.Count != 1)
            {
                var availableSignatures = namedMethods.Count == 0
                    ? " No overloads with that name are loaded."
                    : " Loaded overload(s): " +
                      string.Join(
                          "; ",
                          namedMethods.Select(method =>
                              method.ToString() + " [" +
                              string.Join(
                                  ", ",
                                  method.GetParameters().Select(parameter =>
                                      parameter.Name + ": " +
                                      parameter.ParameterType.FullName +
                                      (parameter.IsOut ? " out" : string.Empty) +
                                      (parameter.IsIn ? " in" : string.Empty))) +
                              "]"));
                throw new InvalidOperationException(
                    "Expected one compatible Plant utility method '" +
                    methodName + "'; found " + methods.Count + "." +
                    availableSignatures);
            }

            return methods[0];
        }

        private static bool IsSyncConflictsOutput(ParameterInfo parameter)
        {
            return parameter.IsOut &&
                   parameter.ParameterType.IsByRef &&
                   string.Equals(
                       parameter.ParameterType.GetElementType()?.Name,
                       "PnPSyncConflicts",
                       StringComparison.Ordinal);
        }

        private static object InvokeDocumentServerUtility(
            MethodInfo method,
            object[] arguments)
        {
            try
            {
                return method.Invoke(null, arguments);
            }
            catch (TargetInvocationException exception)
                when (exception.InnerException != null)
            {
                throw new InvalidOperationException(
                    "Plant collaboration API method '" + method +
                    "' failed: " +
                    DescribeExceptionChain(exception.InnerException),
                    exception.InnerException);
            }
        }

        private static string DescribeExceptionChain(System.Exception exception)
        {
            var messages = new List<string>();
            for (var current = exception;
                 current != null;
                 current = current.InnerException)
            {
                var statusCode = current is HttpRequestException httpException
                    ? httpException.StatusCode
                    : null;
                var statusText = statusCode.HasValue
                    ? " [HTTP " + statusCode.Value + "]"
                    : string.Empty;
                messages.Add(
                    current.GetType().FullName + ": " +
                    current.Message + statusText);
            }

            return string.Join(" --> ", messages);
        }

        private static string GetProjectPartRootFolderPath(
            Project projectPart,
            string projectRootPath)
        {
            var fullProjectRootPath = Path.GetFullPath(projectRootPath);
            foreach (var item in projectPart.ProjectItems)
            {
                if (item is PnPProjectFolder projectFolder)
                {
                    var resolvedFolderPath = projectPart.GetLocalResolvedFolderPath(
                        projectFolder);
                    if (!string.IsNullOrWhiteSpace(resolvedFolderPath) &&
                        IsPathWithinRoot(resolvedFolderPath, fullProjectRootPath))
                    {
                        var partRootPath = Path.GetDirectoryName(
                            Path.GetFullPath(resolvedFolderPath));
                        if (!string.IsNullOrWhiteSpace(partRootPath) &&
                            Directory.Exists(partRootPath))
                        {
                            return partRootPath;
                        }
                    }
                }

                if (!string.Equals(
                        ReadItemProperty(item, "ItemType"),
                        "DwgFile",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var relativeFileName = ReadItemProperty(item, "RelativeFileName");
                if (string.IsNullOrWhiteSpace(relativeFileName) ||
                    Path.IsPathRooted(relativeFileName))
                {
                    continue;
                }

                var relativeSegments = relativeFileName
                    .Replace('/', Path.DirectorySeparatorChar)
                    .Replace('\\', Path.DirectorySeparatorChar)
                    .Split(
                        new[] { Path.DirectorySeparatorChar },
                        StringSplitOptions.RemoveEmptyEntries);
                var candidatePartRootPath = relativeSegments.Length > 1
                    ? Path.Combine(fullProjectRootPath, relativeSegments[0])
                    : fullProjectRootPath;
                if (Directory.Exists(candidatePartRootPath) &&
                    IsPathWithinRoot(candidatePartRootPath, fullProjectRootPath))
                {
                    return Path.GetFullPath(candidatePartRootPath);
                }
            }

            throw new InvalidOperationException(
                "Could not resolve the selected Plant project part's root folder from its project items.");
        }

        private static bool PathsAreEqual(string firstPath, string secondPath)
        {
            return string.Equals(
                Path.GetFullPath(firstPath),
                Path.GetFullPath(secondPath),
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool FilesHaveSameContents(
            string firstPath,
            string secondPath)
        {
            var firstFile = new FileInfo(firstPath);
            var secondFile = new FileInfo(secondPath);
            if (firstFile.Length != secondFile.Length)
            {
                return false;
            }

            using (var firstStream = File.OpenRead(firstPath))
            using (var secondStream = File.OpenRead(secondPath))
            using (var sha256 = SHA256.Create())
            {
                return sha256.ComputeHash(firstStream)
                    .SequenceEqual(sha256.ComputeHash(secondStream));
            }
        }

        private static bool CopySourceToDestination(DrawingRecord record)
        {
            if (!record.CopySourceToDestination)
            {
                return false;
            }

            if (File.Exists(record.DestinationFilePath))
            {
                if (!FilesHaveSameContents(
                        record.SourceFilePath,
                        record.DestinationFilePath))
                {
                    throw new InvalidOperationException(
                        "A different DWG appeared at the destination after preflight; it was not overwritten: " +
                        record.DestinationFilePath);
                }

                record.FilePath = record.DestinationFilePath;
                record.CopySourceToDestination = false;
                return false;
            }

            var stagingFilePath = Path.Combine(
                Path.GetDirectoryName(record.DestinationFilePath),
                "." + Path.GetFileName(record.DestinationFilePath) + "." +
                Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.Copy(record.SourceFilePath, stagingFilePath, false);
                File.Move(stagingFilePath, record.DestinationFilePath);
            }
            finally
            {
                if (File.Exists(stagingFilePath))
                {
                    File.Delete(stagingFilePath);
                }
            }

            record.FilePath = record.DestinationFilePath;
            record.CopySourceToDestination = false;
            return true;
        }

        private static void ImportRecords(
            Project projectPart,
            List<DrawingRecord> records,
            string projectRootPath,
            Editor editor)
        {
            var folderGuids = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            foreach (var folderPath in records
                         .Select(record => record.RelativeFolderPath)
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(folderPath))
                {
                    folderGuids.Add(folderPath, string.Empty);
                    continue;
                }

                var folder = FindProjectFolder(
                    projectPart,
                    folderPath);
                if (folder == null)
                {
                    throw new InvalidOperationException(
                        "Destination folder '" + folderPath +
                        "' no longer exists. No drawings were added; create it in Plant 3D Project Manager and rerun.");
                }
                folderGuids.Add(folderPath, GetProjectFolderGuid(folder));
            }

            var createdCount = 0;
            var registeredDrawingPaths = new List<string>();
            var xrefValidationWarnings = new List<string>();
            foreach (var record in records)
            {
                var copiedToProject = false;
                var addAttempted = false;
                try
                {
                    copiedToProject = CopySourceToDestination(record);
                    if (copiedToProject)
                    {
                        editor.WriteMessage(
                            "\nCopied source DWG to project folder: {0}",
                            record.DestinationFilePath);
                    }

                    var fileAlreadyExists = false;
                    var illegalXrefs = false;
                    var syncConflicts = default(PnPSyncConflicts);
                    PnPRowIdArray drawingIdsOutOfSync;
                    var title = GetTitle(record);

                    // AddPnPDrawingFile returns an empty result without registering the drawing
                    // when its own XREF validation fails, so validate separately and report only.
                    var xrefsAreLegal = projectPart.AllXrefsAreLegal(record.FilePath, false);

                    addAttempted = true;
                    var addResult = projectPart.AddPnPDrawingFile(
                        record.FilePath,
                        folderGuids[record.RelativeFolderPath],
                        title,
                        false,
                        ref fileAlreadyExists,
                        ref illegalXrefs,
                        ref syncConflicts,
                        out drawingIdsOutOfSync);

                    if (fileAlreadyExists)
                    {
                        throw new InvalidOperationException(
                            "Plant 3D reports that this drawing already exists.");
                    }
                    if (HasSyncResults(syncConflicts) ||
                        HasSyncResults(drawingIdsOutOfSync))
                    {
                        throw new InvalidOperationException(
                            "Plant 3D reported synchronization conflicts or out-of-sync drawings for this add.");
                    }
                    if (string.IsNullOrEmpty(addResult))
                    {
                        throw new InvalidOperationException(
                            "Plant 3D did not register the drawing (AddPnPDrawingFile returned no drawing identifier).");
                    }
                    illegalXrefs = illegalXrefs || !xrefsAreLegal;

                    var projectDrawingPath = GetProjectDrawingPath(
                        projectPart,
                        folderGuids[record.RelativeFolderPath],
                        record.FileName,
                        addResult,
                        record.FilePath);
                    registeredDrawingPaths.Add(projectDrawingPath);
                    if (illegalXrefs)
                    {
                        xrefValidationWarnings.Add(record.RelativePath);
                    }

                    foreach (var propertyValue in record.PropertyValues)
                    {
                        if (IsDrawingTitleProperty(propertyValue.Mapping))
                        {
                            editor.WriteMessage(
                                "\nSetting CSV DWG Title for {0} to '{1}'.",
                                record.FileName,
                                propertyValue.Value);
                        }

                        if (DrawingPropertyHasValue(
                                projectPart,
                                projectDrawingPath,
                                propertyValue.Mapping,
                                propertyValue.Value))
                        {
                            continue;
                        }

                        if (!projectPart.SetDrawingPropertyValue(
                                projectDrawingPath,
                                propertyValue.Mapping.Category,
                                propertyValue.Mapping.Name,
                                propertyValue.Value) &&
                            !DrawingPropertyHasValue(
                                projectPart,
                                projectDrawingPath,
                                propertyValue.Mapping,
                                propertyValue.Value))
                        {
                            throw new InvalidOperationException(
                                "Plant 3D rejected property " +
                                propertyValue.Mapping.Category + " / " +
                                propertyValue.Mapping.DisplayName +
                                " with value '" + propertyValue.Value +
                                "' for drawing '" + projectDrawingPath + "'.");
                        }
                    }

                    if (!projectPart.AcceptDrawingProperties(projectDrawingPath))
                    {
                        throw new InvalidOperationException(
                            "Plant 3D did not accept the drawing property changes.");
                    }

                    createdCount++;
                    editor.WriteMessage(
                        "\nAdded and updated {0} ({1}/{2}).",
                        record.FileName,
                        createdCount,
                        records.Count);
                }
                catch (System.Exception exception)
                {
                    editor.WriteMessage(
                        "\nImport stopped at '{0}' after {1} completed drawing(s): {2}",
                        record.FileName,
                        createdCount,
                        exception.GetBaseException().Message);
                    if (copiedToProject)
                    {
                        editor.WriteMessage(
                            "\nThe copied DWG remains at '{0}' although import did not complete.",
                            record.DestinationFilePath);
                    }
                    FinishXrefProcessing(
                        registeredDrawingPaths,
                        projectRootPath,
                        xrefValidationWarnings,
                        editor);
                    if (addAttempted)
                    {
                        editor.WriteMessage(
                            "\nThe drawing may already be registered because the failure occurred during or after AddPnPDrawingFile.");
                    }
                    return;
                }
            }

            FinishXrefProcessing(
                registeredDrawingPaths,
                projectRootPath,
                xrefValidationWarnings,
                editor);
            editor.WriteMessage(
                "\nImport finished. {0} drawing(s) were added and their properties accepted.",
                createdCount);
            editor.WriteMessage(
                "\nCheck in the drawings through Plant 3D Project Manager to publish them to the collaborative project.");
        }

        private static void FinishXrefProcessing(
            List<string> drawingPaths,
            string projectRootPath,
            List<string> xrefValidationWarnings,
            Editor editor)
        {
            if (drawingPaths.Count == 0)
            {
                return;
            }

            var conversionWarnings = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            var convertedDrawingCount = 0;
            var convertedXrefCount = 0;
            foreach (var drawingPath in drawingPaths.Distinct(
                         StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    if (!File.Exists(drawingPath))
                    {
                        conversionWarnings.Add(
                            "Project drawing is not available locally: " +
                            drawingPath);
                        continue;
                    }
                    if (!IsPathWithinRoot(drawingPath, projectRootPath))
                    {
                        conversionWarnings.Add(
                            "Project drawing is outside the active project's local folder: " +
                            drawingPath);
                        continue;
                    }

                    var changedXrefCount = ConvertDrawingXrefsToRelative(
                        drawingPath,
                        projectRootPath,
                        conversionWarnings);
                    if (changedXrefCount > 0)
                    {
                        convertedDrawingCount++;
                        convertedXrefCount += changedXrefCount;
                    }
                }
                catch (System.Exception exception)
                {
                    conversionWarnings.Add(
                        "Could not convert XREF paths in '" + drawingPath +
                        "': " + exception.GetBaseException().Message);
                }
            }

            editor.WriteMessage(
                "\nRelative XREF conversion updated {0} reference path(s) in {1} drawing(s).",
                convertedXrefCount,
                convertedDrawingCount);

            if (xrefValidationWarnings.Count > 0)
            {
                editor.WriteMessage(
                    "\nPlant reported XREF validation issues while registering {0} drawing(s); these did not block import.",
                    xrefValidationWarnings.Count);
                foreach (var drawingPath in xrefValidationWarnings
                             .Distinct(StringComparer.OrdinalIgnoreCase)
                             .Take(10))
                {
                    editor.WriteMessage("\n  {0}", drawingPath);
                }
                if (xrefValidationWarnings
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count() > 10)
                {
                    editor.WriteMessage("\n  Additional flagged drawings omitted.");
                }
            }

            if (conversionWarnings.Count > 0)
            {
                editor.WriteMessage(
                    "\nSome XREF paths were left unchanged:");
                foreach (var warning in conversionWarnings.Take(20))
                {
                    editor.WriteMessage("\n  {0}", warning);
                }
                if (conversionWarnings.Count > 20)
                {
                    editor.WriteMessage("\n  Additional conversion warnings omitted.");
                }
            }
        }

        private static int ConvertDrawingXrefsToRelative(
            string drawingPath,
            string projectRootPath,
            HashSet<string> warnings)
        {
            var xrefIds = new ObjectIdCollection();
            var changedCount = 0;
            using (var database = new Database(false, true))
            {
                database.ReadDwgFile(
                    drawingPath,
                    FileOpenMode.OpenForReadAndWriteNoShare,
                    false,
                    string.Empty);

                using (var transaction =
                       database.TransactionManager.StartTransaction())
                {
                    var blockTable = (BlockTable)transaction.GetObject(
                        database.BlockTableId,
                        OpenMode.ForRead);
                    foreach (ObjectId blockRecordId in blockTable)
                    {
                        var blockRecord = (BlockTableRecord)transaction.GetObject(
                            blockRecordId,
                            OpenMode.ForRead);
                        if (!blockRecord.IsFromExternalReference)
                        {
                            continue;
                        }

                        var oldPath = blockRecord.PathName;
                        if (!Path.IsPathRooted(oldPath))
                        {
                            continue;
                        }

                        var xrefPath = Path.GetFullPath(oldPath);
                        if (!IsPathWithinRoot(xrefPath, projectRootPath))
                        {
                            warnings.Add(
                                "Absolute XREF is outside the active project's local folder and was not changed: " +
                                drawingPath + " -> " + oldPath);
                            continue;
                        }
                        if (!File.Exists(xrefPath))
                        {
                            warnings.Add(
                                "Absolute XREF target is missing and was not changed: " +
                                drawingPath + " -> " + oldPath);
                            continue;
                        }

                        var relativePath = Path.GetRelativePath(
                            Path.GetDirectoryName(drawingPath),
                            xrefPath);
                        if (Path.IsPathRooted(relativePath))
                        {
                            warnings.Add(
                                "Could not form a relative XREF path: " +
                                drawingPath + " -> " + oldPath);
                            continue;
                        }

                        blockRecord.UpgradeOpen();
                        blockRecord.PathName = relativePath;
                        xrefIds.Add(blockRecordId);
                        changedCount++;
                    }

                    transaction.Commit();
                }

                if (changedCount > 0)
                {
                    database.SaveAs(
                        drawingPath,
                        true,
                        database.OriginalFileVersion,
                        database.SecurityParameters);
                    try
                    {
                        database.ReloadXrefs(xrefIds);
                    }
                    catch (System.Exception exception)
                    {
                        warnings.Add(
                            "Relative XREF paths were saved, but references could not be reloaded in '" +
                            drawingPath + "': " +
                            exception.GetBaseException().Message);
                    }
                }
            }

            return changedCount;
        }

        private static bool IsPathWithinRoot(string path, string rootPath)
        {
            var fullPath = Path.GetFullPath(path);
            var fullRoot = Path.GetFullPath(rootPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return string.Equals(
                       fullPath,
                       fullRoot,
                       StringComparison.OrdinalIgnoreCase) ||
                   fullPath.StartsWith(
                       fullRoot + Path.DirectorySeparatorChar,
                       StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasSyncResults(object result)
        {
            if (result == null)
            {
                return false;
            }

            if (result is IEnumerable enumerable)
            {
                var enumerator = enumerable.GetEnumerator();
                try
                {
                    if (enumerator.MoveNext())
                    {
                        return true;
                    }
                }
                finally
                {
                    (enumerator as IDisposable)?.Dispose();
                }
            }

            var resultType = result.GetType();
            var hasConflictsProperty = resultType.GetProperty("HasConflicts");
            if (hasConflictsProperty != null &&
                hasConflictsProperty.PropertyType == typeof(bool) &&
                (bool)hasConflictsProperty.GetValue(result))
            {
                return true;
            }

            foreach (var propertyName in new[] { "Count", "Length" })
            {
                var countProperty = resultType.GetProperty(propertyName);
                if (countProperty != null)
                {
                    var count = Convert.ToInt32(countProperty.GetValue(result));
                    if (count > 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool DrawingPropertyHasValue(
            Project projectPart,
            string drawingPath,
            PropertyMapping mapping,
            string expectedValue)
        {
            return projectPart.GetDrawingPropertyValue(
                       drawingPath,
                       mapping.Category,
                       mapping.Name,
                       out var currentValue) &&
                   string.Equals(
                       currentValue,
                       expectedValue,
                       StringComparison.Ordinal);
        }

        private static bool IsDrawingTitleProperty(PropertyMapping mapping)
        {
            return string.Equals(
                       mapping.Name,
                       "DWG_Title",
                       StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(
                       NormalizePropertyName(mapping.DisplayName),
                       "DWGTITLE",
                       StringComparison.OrdinalIgnoreCase);
        }

        private static string GetProjectFolderGuid(PnPProjectFolder folder)
        {
            var folderGuid = folder.FolderGuid;
            if (string.IsNullOrWhiteSpace(folderGuid))
            {
                folderGuid = folder.Guid;
            }
            if (string.IsNullOrWhiteSpace(folderGuid))
            {
                throw new InvalidOperationException(
                    "Plant 3D folder '" + folder.Name + "' has no folder GUID.");
            }
            return folderGuid;
        }

        private static PnPProjectFolder FindProjectFolder(
            Project projectPart,
            string relativeFolderPath)
        {
            if (string.IsNullOrWhiteSpace(relativeFolderPath))
            {
                return null;
            }

            string parentGuid = null;
            PnPProjectFolder folder = null;
            foreach (var segment in relativeFolderPath.Split(Path.DirectorySeparatorChar))
            {
                folder = projectPart.GetChildFolder(parentGuid, segment);
                if (folder == null)
                {
                    return null;
                }
                parentGuid = GetProjectFolderGuid(folder);
            }

            return folder;
        }

        private static string GetTitle(DrawingRecord record)
        {
            var title = record.PropertyValues.FirstOrDefault(value =>
                IsDrawingTitleProperty(value.Mapping));
            return title == null || string.IsNullOrWhiteSpace(title.Value)
                ? Path.GetFileNameWithoutExtension(record.FileName)
                : title.Value;
        }

        private static string GetProjectDrawingPath(
            Project projectPart,
            string folderGuid,
            string fileName,
            string addResult,
            string sourcePath)
        {
            var projectFolder = string.IsNullOrWhiteSpace(folderGuid)
                ? null
                : projectPart.GetFolder(folderGuid);
            if (!string.IsNullOrWhiteSpace(folderGuid) && projectFolder == null)
            {
                throw new InvalidOperationException(
                    "Plant 3D could not resolve the destination project folder.");
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
                        fileName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    var absoluteFileName = ReadItemProperty(item, "AbsoluteFileName");
                    if (!string.IsNullOrWhiteSpace(absoluteFileName))
                    {
                        return absoluteFileName;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(addResult) && File.Exists(addResult))
            {
                return addResult;
            }

            return sourcePath;
        }

        private static string ReadItemProperty(object item, string propertyName)
        {
            var property = item.GetType().GetProperty(propertyName);
            return property == null
                ? null
                : Convert.ToString(property.GetValue(item));
        }

        private static string NormalizeEnteredPath(string path)
        {
            return path.Trim().Trim('"');
        }

        private sealed class CollaborationFolderAddException : InvalidOperationException
        {
            public CollaborationFolderAddException(
                string message,
                System.Exception innerException)
                : base(message, innerException)
            {
            }
        }

        private sealed class PropertyMapping
        {
            public PropertyMapping(string category, string name, string displayName)
            {
                Category = category;
                Name = name;
                DisplayName = displayName;
            }

            public string Category { get; }
            public string Name { get; }
            public string DisplayName { get; }
        }

        private sealed class PropertyValue
        {
            public PropertyValue(PropertyMapping mapping, string value)
            {
                Mapping = mapping;
                Value = value;
            }

            public PropertyMapping Mapping { get; }
            public string Value { get; }
        }

        private sealed class DrawingRecord
        {
            public DrawingRecord(
                string fileName,
                string relativePath,
                string relativeFolderPath,
                string filePath,
                List<PropertyValue> propertyValues)
            {
                FileName = fileName;
                RelativePath = relativePath;
                RelativeFolderPath = relativeFolderPath;
                FilePath = filePath;
                SourceFilePath = filePath;
                PropertyValues = propertyValues;
            }

            public string FileName { get; }
            public string RelativePath { get; }
            public string RelativeFolderPath { get; }
            public string FilePath { get; set; }
            public string SourceFilePath { get; }
            public string DestinationFilePath { get; set; }
            public bool CopySourceToDestination { get; set; }
            public List<PropertyValue> PropertyValues { get; }
        }
    }
}
