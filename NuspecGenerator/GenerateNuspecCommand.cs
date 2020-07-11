using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace NuspecGenerator
{
    [SuppressMessage("Usage", "VSTHRD010:Invoke single-threaded types on Main thread", Justification = "<Pending>")]
    internal sealed class GenerateNuspecCommand
    {
        public const int CommandId = 0x0100;
        public const string NuspecFileName = ".nuspec";
        public const string AssemblyInfoFileName = "AssemblyInfo.cs";
        public const string PackagesConfigFileName = "packages.config";
        public const string CSharpProjectTypeGuid = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}";

        public static readonly Guid CommandSet = new Guid("21c3ed1d-dcfb-4145-ab59-1db105d28c68");

        public DTE2 Dte2;

        public static GenerateNuspecCommand Instance { get; private set; }


        private GenerateNuspecCommand(OleMenuCommandService commandService, DTE2 dte2)
        {
            Dte2 = dte2;
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var commandId = new CommandID(CommandSet, CommandId);
            var menuCommand = new OleMenuCommand(Execute, commandId);
            commandService.AddCommand(menuCommand);
        }

        private static void AddNuspecFileToProject(Project project)
        {
            var packagesConfigFile =
                ProjectHelper.GetFile(Path.GetDirectoryName(project.FileName), PackagesConfigFileName) != null;

            if (packagesConfigFile)
            {
                Logger.Log("The packages.config method is not supported, migrate to PackageReference first.");
                return;
            }

            var package = CreatePackage(project.FileName);
            var nuspecFilePath = Path.Combine(Path.GetDirectoryName(project.FileName)!, NuspecFileName);

            GenerateXml(package, nuspecFilePath);
            project.ProjectItems.AddFromFile(nuspecFilePath);
        }

        private static package CreatePackage(string projectFile)
        {
            var packageReferences = ProjectHelper.GetPackageReferences(projectFile);
            var dependencies = GetDependencies(packageReferences);
            var tokens = GetTokens(Path.GetDirectoryName(projectFile),
                new[] {Path.GetFileName(projectFile), AssemblyInfoFileName});

            var package = new package
            {
                metadata = CreatePackageMetadata(tokens, dependencies)
            };

            return package;
        }

        private static packageMetadata CreatePackageMetadata(IReadOnlyDictionary<string, string> tokens,
            packageMetadataDependencies dependencies)
        {
            return new packageMetadata
            {
                id = !string.IsNullOrEmpty(tokens["id"]) ? tokens["id"] : "$id$",
                version = !string.IsNullOrEmpty(tokens["version"]) ? tokens["version"] : "$version$",
                authors = !string.IsNullOrEmpty(tokens["author"]) ? tokens["author"] : "$author$",
                owners = !string.IsNullOrEmpty(tokens["author"]) ? tokens["author"] : "$owners$",
                title = !string.IsNullOrEmpty(tokens["title"]) ? tokens["title"] : "$title$",
                description = !string.IsNullOrEmpty(tokens["description"]) ? tokens["description"] : "$description$",
                copyright = !string.IsNullOrEmpty(tokens["copyright"]) ? tokens["copyright"] : "$copyright$",
                requireLicenseAcceptance = false,
                dependencies = dependencies.Items.Length == 0 ? null : dependencies
            };
        }

        private static void Execute(object sender, EventArgs e)
        {
            try
            {
                var project = ProjectHelper.GetSelectedProject();
                var hierarchy = ProjectHelper.GetProjectHierarchy(project);

                if (project == null ||
                    !project.Kind.Equals(CSharpProjectTypeGuid, StringComparison.InvariantCultureIgnoreCase) ||
                    ProjectHelper.IsCpsProject(hierarchy))
                {
                    var messsage = project == null
                        ? "No or multiple project files selected, select one project."
                        : "Project type is not supported.";
                    Logger.Log(messsage);
                    return;
                }

                project.Save();

                var nuspecFile = ProjectHelper.GetFile(Path.GetDirectoryName(project.FileName), NuspecFileName);

                if (nuspecFile == null)
                {
                    Logger.Log("Adding Nuspec projectFile...");
                    AddNuspecFileToProject(project);
                }
                else
                {
                    Logger.Log("Updating Nuspec projectFile...");
                    UpdateNuspecFile(project.FileName, nuspecFile);
                }

                Logger.Log("Done.");
            }
            catch (Exception ex)
            {
                Logger.Log($"An error occured: {ex}");
            }
        }

        private static void GenerateXml(package package, string nuspecFilePath)
        {
            var xmlSerializer = new XmlSerializer(typeof(package));
            TextWriter textWriter = new StreamWriter(nuspecFilePath, false, Encoding.UTF8);
            xmlSerializer.Serialize(textWriter, package);
            textWriter.Close();
        }

        private static packageMetadataDependencies GetDependencies(Dictionary<string, string> packageReferences)
        {
            var dependencies = new packageMetadataDependencies
            {
                Items = packageReferences.Select(x => new dependency {id = x.Key, version = x.Value}).ToArray<object>()
            };
            return dependencies;
        }

        private static Dictionary<string, string> GetTokens(string path, string[] fileNames)
        {
            var lines = ProjectHelper.GetLines(path, fileNames);

            var tokens = new Dictionary<string, string>
            {
                {"id", Regex.Match(lines, @"<AssemblyName>(.*)<\/AssemblyName>").Groups[1].Value},
                {
                    "version",
                    Regex.Match(lines, @"\n\[assembly:\s*AssemblyInformationalVersion\s*\(\s*""([0-9\.\*]*?)""\s*\)")
                        .Groups[1].Success
                        ? Regex.Match(lines,
                                @"\n\[assembly:\s*AssemblyInformationalVersion\s*\(\s*""([0-9\.\*]*?)""\s*\)").Groups[1]
                            .Value
                        : Regex.Match(lines, @"\n\[assembly:\s*AssemblyVersion\s*\(\s*""([0-9\.\*]*?)""\s*\)").Groups[1]
                            .Value
                },
                {"author", Regex.Match(lines, @"\n\[assembly:\s*AssemblyCompany\s*\(\s*""(.*)""\s*\)").Groups[1].Value},
                {"title", Regex.Match(lines, @"\n\[assembly:\s*AssemblyTitle\s*\(\s*""(.*)""\s*\)").Groups[1].Value},
                {
                    "description",
                    Regex.Match(lines, @"\n\[assembly:\s*AssemblyDescription\s*\(\s*""(.*)""\s*\)").Groups[1].Value
                },
                {
                    "copyright",
                    Regex.Match(lines, @"\n\[assembly:\s*AssemblyCopyright\s*\(\s*""(.*)""\s*\)").Groups[1].Value
                }
            };

            return tokens;
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            var dte2 = await package.GetServiceAsync(typeof(DTE)) as DTE2;
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GenerateNuspecCommand(commandService, dte2);
        }

        private static void UpdateNuspecFile(string projectFile, string nuspecFilePath)
        {
            var package = UpdatePackage(projectFile, nuspecFilePath);
            GenerateXml(package, nuspecFilePath);
        }

        private static package UpdatePackage(string projectFile, string nuspecFilePath)
        {
            var xml = XDocument.Load(nuspecFilePath).ToString();
            var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            var xmlTextReader = new XmlTextReader(memoryStream) {Namespaces = true};
            var xmlSerializer = new XmlSerializer(typeof(package));
            var tokens = GetTokens(Path.GetDirectoryName(projectFile),
                new[] {Path.GetFileName(projectFile), AssemblyInfoFileName});
            var packageReferences = ProjectHelper.GetPackageReferences(projectFile);
            var dependencies = GetDependencies(packageReferences);
            var packageMetadata = CreatePackageMetadata(tokens, dependencies);
            var package = (package) xmlSerializer.Deserialize(xmlTextReader);
            var updatedPackage = UpdatePackage(package, packageMetadata);

            return updatedPackage;
        }

        private static package UpdatePackage(package package, packageMetadata packageMetadata)
        {
            package.metadata.id = string.IsNullOrEmpty(package.metadata.id)
                ? packageMetadata.id
                : package.metadata.id;

            package.metadata.version = string.IsNullOrEmpty(package.metadata.version)
                ? packageMetadata.version
                : package.metadata.version;

            package.metadata.authors = string.IsNullOrEmpty(package.metadata.authors)
                ? packageMetadata.authors
                : package.metadata.authors;

            package.metadata.owners = string.IsNullOrEmpty(package.metadata.owners)
                ? packageMetadata.owners
                : package.metadata.owners;

            package.metadata.title = string.IsNullOrEmpty(package.metadata.title)
                ? packageMetadata.title
                : package.metadata.title;

            package.metadata.description = string.IsNullOrEmpty(package.metadata.description)
                ? packageMetadata.description
                : package.metadata.description;

            package.metadata.copyright = string.IsNullOrEmpty(package.metadata.copyright)
                ? packageMetadata.copyright
                : package.metadata.copyright;

            package.metadata.dependencies =
                packageMetadata.dependencies.Items.Length == 0 ? null : packageMetadata.dependencies;

            return package;
        }
    }
}