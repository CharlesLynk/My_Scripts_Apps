using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Plant3DMetadataAddin
{
    public class ProjectApiInfoCommand
    {
        [CommandMethod("P3DPROJECTAPIINFO")]
        public void ShowProjectApiInfo()
        {
            var editor = Application.DocumentManager.MdiActiveDocument?.Editor;
            if (editor == null)
            {
                Application.ShowAlertDialog(
                    "Open a drawing in Plant 3D before running P3DPROJECTAPIINFO.");
                return;
            }

            var projectManagerAssembly = AppDomain.CurrentDomain
                .GetAssemblies()
                .FirstOrDefault(assembly =>
                    assembly.GetName().Name.IndexOf(
                        "PnPProjectManagerMgd",
                        StringComparison.OrdinalIgnoreCase) >= 0);

            if (projectManagerAssembly == null)
            {
                editor.WriteMessage(
                    "\nPlant 3D Project Manager assembly is not loaded.");
                return;
            }

            editor.WriteMessage(
                "\nProject Manager assembly: {0}",
                projectManagerAssembly.FullName);

            var activeProjectFound = false;
            object activePlantProject = null;
            var apiMethodNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                Type[] types;
                try
                {
                    types = assembly.GetTypes();
                }
                catch (ReflectionTypeLoadException exception)
                {
                    types = exception.Types.Where(type => type != null).ToArray();
                }

                foreach (var type in types.Where(type =>
                             type.Name.IndexOf(
                                 "PlantApplication",
                                 StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    var currentProjectProperty = type.GetProperty(
                        "CurrentProject",
                        BindingFlags.Public | BindingFlags.Static);
                    if (currentProjectProperty == null)
                    {
                        continue;
                    }

                    object currentProject;
                    try
                    {
                        currentProject = currentProjectProperty.GetValue(null);
                    }
                    catch (System.Exception exception)
                    {
                        editor.WriteMessage(
                            "\nCould not read {0}.CurrentProject: {1}",
                            type.FullName,
                            exception.GetBaseException().Message);
                        continue;
                    }

                    if (currentProject == null)
                    {
                        editor.WriteMessage(
                            "\n{0}.CurrentProject is null.",
                            type.FullName);
                        continue;
                    }

                    activeProjectFound = true;
                    activePlantProject = currentProject;
                    editor.WriteMessage(
                        "\nActive project type: {0}",
                        currentProject.GetType().FullName);
                    editor.WriteMessage(
                        "\nActive project assembly: {0}",
                        currentProject.GetType().Assembly.FullName);
                    editor.WriteMessage("\nProject type hierarchy:");
                    for (var projectType = currentProject.GetType();
                         projectType != null;
                         projectType = projectType.BaseType)
                    {
                        editor.WriteMessage("\n  {0}", projectType.FullName);
                        foreach (var property in projectType.GetProperties(
                                     BindingFlags.Public |
                                     BindingFlags.Instance |
                                     BindingFlags.Static |
                                     BindingFlags.DeclaredOnly)
                                 .OrderBy(property => property.Name))
                        {
                            editor.WriteMessage(
                                "\n    Property {0}: {1}",
                                property.Name,
                                property.PropertyType.FullName);
                        }

                        foreach (var method in projectType.GetMethods(
                                     BindingFlags.Public |
                                     BindingFlags.Instance |
                                     BindingFlags.Static |
                                     BindingFlags.DeclaredOnly)
                                 .Where(IsDrawingPropertyApiMethod)
                                 .OrderBy(method => method.Name))
                        {
                            apiMethodNames.Add(FormatApiMethod(projectType, method));
                        }
                    }

                    foreach (var property in currentProject.GetType().GetProperties(
                                 BindingFlags.Public |
                                 BindingFlags.Instance |
                                 BindingFlags.Static)
                             .Where(property => property.Name.IndexOf(
                                 "path",
                                 StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 property.Name.IndexOf(
                                     "name",
                                     StringComparison.OrdinalIgnoreCase) >= 0)
                             .OrderBy(property => property.Name))
                    {
                        try
                        {
                            editor.WriteMessage(
                                "\n  {0} = {1}",
                                property.Name,
                                property.GetValue(
                                    property.GetMethod.IsStatic ? null : currentProject));
                        }
                        catch (System.Exception exception)
                        {
                            editor.WriteMessage(
                                "\n  {0} could not be read: {1}",
                                property.Name,
                                exception.GetBaseException().Message);
                        }
                    }
                }
            }

            foreach (var assembly in new[] { projectManagerAssembly })
            {
                Type[] types;
                try
                {
                    types = assembly.GetTypes();
                }
                catch (ReflectionTypeLoadException exception)
                {
                    types = exception.Types.Where(type => type != null).ToArray();
                    var loadFailure = exception.LoaderExceptions
                        .FirstOrDefault(loaderException => loaderException != null);
                    if (loadFailure != null)
                    {
                        editor.WriteMessage(
                            "\nSome types in {0} could not be inspected: {1}",
                            assembly.GetName().Name,
                            loadFailure.Message);
                    }
                }

                foreach (var type in types)
                {
                    if (type.FullName == "Autodesk.ProcessPower.ProjectManager.Project" ||
                        type.FullName == "Autodesk.ProcessPower.ProjectManager.DrawingProperty" ||
                        type.FullName == "Autodesk.ProcessPower.ProjectManager.PnPProjectFolder" ||
                        type.FullName == "Autodesk.ProcessPower.ProjectManager.ProjectPartCollection")
                    {
                        editor.WriteMessage(
                            "\nAPI type: {0} [{1}]",
                            type.FullName,
                            type.Assembly.GetName().Name);
                        foreach (var constructor in type.GetConstructors(
                                     BindingFlags.Public |
                                     BindingFlags.Instance)
                                 .OrderBy(constructor => constructor.GetParameters().Length))
                        {
                            editor.WriteMessage(
                                "\n  Constructor {0}",
                                FormatParameters(constructor.GetParameters()));
                        }

                        foreach (var property in type.GetProperties(
                                     BindingFlags.Public |
                                     BindingFlags.Instance |
                                     BindingFlags.Static)
                                 .OrderBy(property => property.Name))
                        {
                            editor.WriteMessage(
                                "\n  Property {0}: {1}",
                                property.Name,
                                property.PropertyType.FullName);
                        }
                    }

                    foreach (var method in type.GetMethods(
                                 BindingFlags.Public |
                                 BindingFlags.Instance |
                                 BindingFlags.Static |
                                 BindingFlags.DeclaredOnly)
                             .Where(IsDrawingPropertyApiMethod))
                    {
                        apiMethodNames.Add(FormatApiMethod(type, method));
                    }
                }
            }

            if (!activeProjectFound)
            {
                editor.WriteMessage(
                    "\nNo active Plant 3D project was found. Open the target project in Project Manager and retry.");
                return;
            }

            if (apiMethodNames.Count == 0)
            {
                editor.WriteMessage(
                    "\nNo public drawing-property API methods were found in loaded Plant/Project Manager assemblies.");
            }
            else
            {
                editor.WriteMessage("\nLoaded Plant/Project Manager drawing-property API methods:");
                foreach (var method in apiMethodNames.OrderBy(method => method))
                {
                    editor.WriteMessage("\n  {0}", method);
                }
            }

            InspectProjectParts(activePlantProject, editor);
            ReportProjectPartDirectories(activePlantProject, editor);
        }

        private static void InspectProjectParts(object plantProject, Editor editor)
        {
            var projectPartsProperty = plantProject.GetType().GetProperty("ProjectParts");
            if (projectPartsProperty == null)
            {
                editor.WriteMessage("\nThe active PlantProject has no public ProjectParts property.");
                return;
            }

            var projectParts = projectPartsProperty.GetValue(plantProject);
            if (projectParts == null)
            {
                editor.WriteMessage("\nThe active PlantProject.ProjectParts collection is null.");
                return;
            }

            var partIndexer = projectParts.GetType().GetProperty(
                "Item",
                new[] { typeof(string) });
            if (partIndexer == null)
            {
                editor.WriteMessage(
                    "\nProjectParts does not expose a string indexer; type: {0}",
                    projectParts.GetType().FullName);
                return;
            }

            foreach (var partName in new[] { "Piping", "PnId", "Ortho", "Iso" })
            {
                object projectPart;
                try
                {
                    projectPart = partIndexer.GetValue(projectParts, new object[] { partName });
                }
                catch (System.Exception exception)
                {
                    editor.WriteMessage(
                        "\nProject part '{0}' could not be read: {1}",
                        partName,
                        exception.GetBaseException().Message);
                    continue;
                }

                if (projectPart == null)
                {
                    continue;
                }

                editor.WriteMessage(
                    "\nProject part '{0}': {1}",
                    partName,
                    projectPart.GetType().FullName);

                var categoriesMethod = projectPart.GetType().GetMethods(
                        BindingFlags.Public | BindingFlags.Instance)
                    .FirstOrDefault(method =>
                    {
                        var parameters = method.GetParameters();
                        return method.Name == "GetCategoryNames" &&
                               parameters.Length == 1 &&
                               parameters[0].ParameterType.IsEnum;
                    });

                if (categoriesMethod == null)
                {
                    editor.WriteMessage("\n  GetCategoryNames(CategoryType) was not found.");
                }
                else
                {
                    var enumType = categoriesMethod.GetParameters()[0].ParameterType;
                    var drawingCategory = Enum.GetNames(enumType)
                        .FirstOrDefault(name => name.Equals(
                            "eDrawing",
                            StringComparison.OrdinalIgnoreCase));
                    if (drawingCategory == null)
                    {
                        editor.WriteMessage(
                            "\n  Drawing category enum member was not found in {0}.",
                            enumType.FullName);
                    }
                    else
                    {
                        try
                        {
                            var categoryNames = categoriesMethod.Invoke(
                                projectPart,
                                new[] { Enum.Parse(enumType, drawingCategory) }) as IEnumerable;
                            if (categoryNames != null)
                            {
                                var metadataMethod = projectPart.GetType().GetMethod(
                                    "GetDrawingPropertyMetadata",
                                    new[] { typeof(string) });
                                foreach (var categoryName in categoryNames)
                                {
                                    var category = Convert.ToString(categoryName);
                                    editor.WriteMessage("\n  Drawing property category: {0}", category);
                                    if (metadataMethod == null)
                                    {
                                        continue;
                                    }

                                    var propertyMetadata = metadataMethod.Invoke(
                                        projectPart,
                                        new object[] { category }) as IEnumerable;
                                    if (propertyMetadata == null)
                                    {
                                        continue;
                                    }

                                    foreach (var metadata in propertyMetadata)
                                    {
                                        editor.WriteMessage(
                                            "\n    {0} | {1} | {2}",
                                            ReadProperty(metadata, "Name"),
                                            ReadProperty(metadata, "DisplayName"),
                                            ReadProperty(metadata, "Description"));
                                    }
                                }
                            }
                        }
                        catch (System.Exception exception)
                        {
                            editor.WriteMessage(
                                "\n  Drawing property metadata could not be read: {0}",
                                exception.GetBaseException().Message);
                        }
                    }
                }

                ReportProjectItems(projectPart, editor);
                if (string.Equals(partName, "PnId", StringComparison.OrdinalIgnoreCase))
                {
                    ReportPnIdFolderPaths(projectPart, editor);
                }
            }
        }

        private static void ReportProjectPartDirectories(
            object plantProject,
            Editor editor)
        {
            var directoriesProperty = plantProject.GetType().GetProperty(
                "ProjectPartDirectories");
            if (directoriesProperty == null)
            {
                editor.WriteMessage("\nProjectPartDirectories property was not found.");
                return;
            }

            var directories = directoriesProperty.GetValue(plantProject) as IEnumerable;
            if (directories == null)
            {
                editor.WriteMessage("\nProjectPartDirectories is not enumerable.");
                return;
            }

            editor.WriteMessage("\nProject-part directory mappings:");
            var mappingCount = 0;
            var pnIdMappingFound = false;
            foreach (var entry in directories)
            {
                mappingCount++;
                var partName = ReadProperty(entry, "Key");
                if (!string.Equals(partName, "PnId", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                pnIdMappingFound = true;
                var fileHelper = entry.GetType().GetProperty("Value")?.GetValue(entry);
                if (fileHelper == null)
                {
                    editor.WriteMessage("\n  PnId mapping has no FileHelper value.");
                    continue;
                }

                editor.WriteMessage(
                    "\n  PnId FileHelper: {0}",
                    fileHelper.GetType().FullName);
                foreach (var property in fileHelper.GetType().GetProperties(
                             BindingFlags.Public | BindingFlags.Instance)
                         .Where(property =>
                             property.Name.IndexOf("path", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             property.Name.IndexOf("folder", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             property.Name.IndexOf("directory", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             property.Name.IndexOf("file", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             property.Name.IndexOf("name", StringComparison.OrdinalIgnoreCase) >= 0)
                         .OrderBy(property => property.Name))
                {
                    try
                    {
                        editor.WriteMessage(
                            "\n    {0} = {1}",
                            property.Name,
                            property.GetValue(fileHelper));
                    }
                    catch (System.Exception exception)
                    {
                        editor.WriteMessage(
                            "\n    {0} could not be read: {1}",
                            property.Name,
                            exception.GetBaseException().Message);
                    }
                }
            }

            if (mappingCount == 0)
            {
                editor.WriteMessage("\n  No project-part directory mappings were exposed.");
            }
            else if (!pnIdMappingFound)
            {
                editor.WriteMessage("\n  No PnId directory mapping was exposed.");
            }
        }

        private static void ReportPnIdFolderPaths(
            object projectPart,
            Editor editor)
        {
            var itemsProperty = projectPart.GetType().GetProperty("ProjectItems");
            var items = itemsProperty?.GetValue(projectPart) as IEnumerable;
            var localPathMethod = projectPart.GetType().GetMethods(
                    BindingFlags.Public | BindingFlags.Instance)
                .FirstOrDefault(method =>
                    method.Name == "GetLocalResolvedFolderPath" &&
                    method.GetParameters().Length == 1);
            var findByPathMethod = projectPart.GetType().GetMethod(
                "FindProjectFolderByPath",
                new[] { typeof(string) });
            if (items == null || localPathMethod == null || findByPathMethod == null)
            {
                editor.WriteMessage(
                    "\nPnId folder-path diagnostics could not find ProjectItems, GetLocalResolvedFolderPath, or FindProjectFolderByPath.");
                return;
            }

            editor.WriteMessage("\nPnId project-folder path mapping:");
            var folderCount = 0;
            foreach (var item in items)
            {
                if (item.GetType().FullName !=
                    "Autodesk.ProcessPower.ProjectManager.PnPProjectFolder")
                {
                    continue;
                }

                string localPath;
                try
                {
                    localPath = Convert.ToString(
                        localPathMethod.Invoke(projectPart, new[] { item }));
                }
                catch (System.Exception exception)
                {
                    editor.WriteMessage(
                        "\n  Folder '{0}' path lookup failed: {1}",
                        ReadProperty(item, "Name"),
                        exception.GetBaseException().Message);
                    continue;
                }

                editor.WriteMessage(
                    "\n  Folder {0}:",
                    ReadProperty(item, "Name"));
                foreach (var propertyName in new[]
                         {
                             "Guid",
                             "FolderGuid",
                             "Folder",
                             "AbsoluteFolderPath",
                             "RelativeFolderPath",
                             "ResolvedFolderPath"
                         })
                {
                    var value = ReadProperty(item, propertyName);
                    if (value != null)
                    {
                        editor.WriteMessage("\n    {0} = {1}", propertyName, value);
                    }
                }
                editor.WriteMessage("\n    GetLocalResolvedFolderPath = {0}", localPath);

                if (folderCount == 0 && !string.IsNullOrWhiteSpace(localPath))
                {
                    var candidatePartRoot = Path.GetDirectoryName(
                        Path.GetFullPath(localPath));
                    var rootFolder = string.IsNullOrWhiteSpace(candidatePartRoot)
                        ? null
                        : findByPathMethod.Invoke(
                            projectPart,
                            new object[] { candidatePartRoot });
                    editor.WriteMessage(
                        "\n    Candidate PnId root = {0}",
                        candidatePartRoot ?? "<null>");
                    editor.WriteMessage(
                        "\n    FindProjectFolderByPath(root) = {0}",
                        rootFolder == null
                            ? "<null; PnId part root is implicit>"
                            : ReadProperty(rootFolder, "Name") + " [" +
                              ReadProperty(rootFolder, "Guid") + "]");
                }

                folderCount++;
                if (folderCount >= 8)
                {
                    editor.WriteMessage("\n  Folder-path listing capped at 8 entries.");
                    break;
                }
            }

            editor.WriteMessage("\n  PnId folders inspected: {0}", folderCount);
        }

        private static void ReportProjectItems(object projectPart, Editor editor)
        {
            var itemsProperty = projectPart.GetType().GetProperty("ProjectItems");
            if (itemsProperty == null)
            {
                editor.WriteMessage("\n  ProjectItems property was not found.");
                return;
            }

            var items = itemsProperty.GetValue(projectPart) as IEnumerable;
            if (items == null)
            {
                editor.WriteMessage("\n  ProjectItems is not enumerable.");
                return;
            }

            var count = 0;
            foreach (var item in items)
            {
                if (count == 40)
                {
                    editor.WriteMessage("\n  ProjectItems listing capped at 40 entries.");
                    break;
                }

                editor.WriteMessage(
                    "\n  Project item ({0}):",
                    item.GetType().FullName);
                foreach (var property in item.GetType().GetProperties(
                             BindingFlags.Public | BindingFlags.Instance)
                         .Where(property =>
                             property.Name.IndexOf("name", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             property.Name.IndexOf("guid", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             property.Name.IndexOf("parent", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             property.Name.IndexOf("type", StringComparison.OrdinalIgnoreCase) >= 0)
                         .OrderBy(property => property.Name))
                {
                    try
                    {
                        editor.WriteMessage(
                            "\n    {0} = {1}",
                            property.Name,
                            property.GetValue(item));
                    }
                    catch (System.Exception exception)
                    {
                        editor.WriteMessage(
                            "\n    {0} could not be read: {1}",
                            property.Name,
                            exception.GetBaseException().Message);
                    }
                }

                count++;
            }

            editor.WriteMessage("\n  ProjectItems inspected: {0}", count);
        }

        private static string ReadProperty(object instance, string name)
        {
            var property = instance.GetType().GetProperty(name);
            if (property == null)
            {
                return "<not exposed>";
            }

            try
            {
                return Convert.ToString(property.GetValue(instance));
            }
            catch (System.Exception exception)
            {
                return "<error: " + exception.GetBaseException().Message + ">";
            }
        }

        private static bool IsDrawingPropertyApiMethod(MethodInfo method)
        {
            return method.Name.IndexOf("AddPnPDrawingFile", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   method.Name.IndexOf("DrawingProperty", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   method.Name.IndexOf("DrawingProperties", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   method.Name.IndexOf("AcceptDrawing", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   method.Name.IndexOf("Folder", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string FormatApiMethod(Type type, MethodInfo method)
        {
            return type.FullName + "." + method.Name +
                   FormatParameters(method.GetParameters()) +
                   " : " + method.ReturnType.FullName;
        }

        private static string FormatParameters(ParameterInfo[] parameters)
        {
            return "(" + string.Join(
                ", ",
                parameters.Select(parameter =>
                    (parameter.IsOut ? "out " : parameter.ParameterType.IsByRef ? "ref " : "") +
                    parameter.Name + ": " +
                    (parameter.ParameterType.IsByRef
                        ? parameter.ParameterType.GetElementType().FullName
                        : parameter.ParameterType.FullName))) + ")";
        }
    }
}
