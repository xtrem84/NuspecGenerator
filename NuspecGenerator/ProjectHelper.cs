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
        private static readonly DTE2 Dte2 = GenerateNuspecCommand.Instance.Dte2;

        internal static Project GetSelectedProject()
        {
            var selectedItems = (Array) Dte2.ToolWindows.SolutionExplorer?.SelectedItems;

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
                    select new {Include = include, Version = version}).ToList().ToDictionary(x => x.Include, x => x.Version);

            return packageReferences;
        }

        internal static string GetLines(string path, string[] fileNames)
        {
            var lines = fileNames.Aggregate(string.Empty,
                (current, fileName) => current + File.ReadAllText(GetFile(path, fileName)));
            return lines;
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