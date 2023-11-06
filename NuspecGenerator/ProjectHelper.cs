using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace NuspecGenerator
{
    [SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "<Pending>")]
    public static class ProjectHelper
    {
        private static string PackagesConfigFileName = "packages.config";

        private static readonly DTE2 Dte2 = GenerateNuspecCommand.Instance.Dte2;

        internal static Project GetSelectedProject()
        {
            var selectedItems = (Array)Dte2.ToolWindows.SolutionExplorer?.SelectedItems;

            return selectedItems!.Cast<UIHierarchyItem>()
                .Select(i =>
                {
                    var project = i.Object as Project;
                    return project;
                })
                .SingleOrDefault();
        }

        internal static string GetFile(string path, string fileName)
        {
            var file = Directory
                .GetFiles(path!, fileName, SearchOption.AllDirectories)
                .FirstOrDefault();

            return file;
        }

        internal static bool IsCpsProject(IVsHierarchy hierarchy)
        {
            Requires.NotNull(hierarchy, "hierarchy");
            return hierarchy.IsCapabilityMatch("CPS");
        }

        internal static Dictionary<string, string> GetPackageReferences(string projectFile)
        {
            var xml = XDocument.Load(projectFile);

            var packageReferences =
                (from element in xml.Descendants().Where(e => e.Name.LocalName == "PackageReference")
                 let include = !string.IsNullOrWhiteSpace(element.Attribute("Include")?.Value)
                     ? element.Attribute("Include")?.Value
                     : element.Elements().First(x => x.Name.LocalName == "Include").Value
                 let version = !string.IsNullOrWhiteSpace(element.Attribute("Version")?.Value)
                     ? element.Attribute("Version")?.Value
                     : element.Elements().First(x => x.Name.LocalName == "Version").Value
                 select new { Include = include, Version = version }).ToList().ToDictionary(x => x.Include, x => x.Version);

            return packageReferences;
        }

        internal static List<string> GetLines(string path, string[] fileNames)
        {
            var lines = new List<string>();
            foreach (string fileName in fileNames)
                lines.AddRange(File.ReadAllLines(GetFile(path, fileName)).ToList());
            return lines;
        }

        internal static string GetAssemblyFileName(Project project)
        {
            return SupportedProjectsManager.GetInstance().GetAssemblyFileNameByProject(project);
        }

        internal static bool IsSupported(Project project)
        {
            var hierarchy = ProjectHelper.GetProjectHierarchy(project);

            if (project == null)
            {
                Logger.Log("No or multiple project files selected, select one project.");
                return false;
            }

            bool isSupportedGuid = SupportedProjectsManager.GetInstance().CheckProjectIsSupported(project);

            if (!isSupportedGuid || ProjectHelper.IsCpsProject(hierarchy))
            {
                Logger.Log("Project type is not supported (only for NetFramework C# & VB.Net)");
                Logger.Log($"{project.Name} project.kind: {project.Kind}");
                return false;
            }

            var packagesConfigFile =
                ProjectHelper.GetFile(Path.GetDirectoryName(project.FileName), PackagesConfigFileName) != null;

            if (packagesConfigFile)
            {
                Logger.Log("The packages.config method is not supported, migrate to PackageReference first.");
                return false;
            }

            return true;
        }

        internal static IVsHierarchy GetProjectHierarchy(Project project)
        {
            if (Package.GetGlobalService(typeof(SVsSolution)) is IVsSolution solution &&
                solution.GetProjectOfUniqueName(project.UniqueName, out var hierarchy) == 0)
                return hierarchy;

            throw new InvalidOperationException("Failed to get project hierarchy");
        }
    }
}