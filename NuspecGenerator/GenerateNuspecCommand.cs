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
        private const int CommandId = 0x0100;
        private const string NuspecFileName = ".nuspec";
        private const string NuspecNamespace = "http://schemas.microsoft.com/packaging/2013/05/nuspec.xsd";
        private static readonly Guid CommandSet = new Guid("21c3ed1d-dcfb-4145-ab59-1db105d28c68");

        public DTE2 Dte2;

        private GenerateNuspecCommand(OleMenuCommandService commandService, DTE2 dte2)
        {
            Dte2 = dte2;
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var commandId = new CommandID(CommandSet, CommandId);
            var menuCommand = new OleMenuCommand(Execute, commandId);
            commandService.AddCommand(menuCommand);
        }

        public static GenerateNuspecCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            var dte2 = await package.GetServiceAsync(typeof(DTE)) as DTE2;
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GenerateNuspecCommand(commandService, dte2);
        }

        private static string GenerateNuspecFile(Project project)
        {
            Logger.Log("Generating .nuspec file...");
            var package = CreatePackage(project);
            var nuspecFilePath = Path.Combine(Path.GetDirectoryName(project.FileName)!, NuspecFileName);

            GenerateXml(package, nuspecFilePath);
            return nuspecFilePath;
        }

        private static package CreatePackage(Project project)
        {
            var packageReferences = ProjectHelper.GetPackageReferences(project.FileName);
            var dependencies = GetDependencies(packageReferences);
            var tokens = GetTokens(Path.GetDirectoryName(project.FileName),
                new[] { Path.GetFileName(project.FileName), ProjectHelper.GetAssemblyFileName(project) });

            var package = new package
            {
                metadata = CreatePackageMetadata(tokens, dependencies)
            };

            return package;
        }

        private static packageMetadata CreatePackageMetadata(IReadOnlyDictionary<string, string> tokens,
            packageMetadataDependencies dependencies)
        {
            var pck = new packageMetadata();

            if (!CheckField("id", tokens)) return null;
            if (!CheckField("description", tokens)) return null;
            if (!CheckField("version", tokens)) return null;
            if (!CheckField("authors", tokens)) return null;

            pck.id = tokens["id"];
            pck.version = tokens["version"];
            pck.authors = tokens["authors"];
            pck.description = tokens["description"];
            pck.copyright = !string.IsNullOrEmpty(tokens["copyright"]) ? tokens["copyright"] : "$copyright$";
            pck.title = !string.IsNullOrEmpty(tokens["title"]) ? tokens["title"] : "$title$";
            pck.requireLicenseAcceptance = false;
            pck.dependencies = dependencies.Items.Length == 0 ? null : dependencies;

            return pck;
        }
        private static bool CheckField(string fieldName, IReadOnlyDictionary<string, string> tokens)
        {
            if (!tokens.ContainsKey(fieldName))
            {
                Logger.Log($"{fieldName} key not found");
                throw new ArgumentNullException(fieldName);
            }
            if (string.IsNullOrEmpty(tokens[fieldName]))
            {
                Logger.Log($"{fieldName} is neccesary");
                throw new ArgumentNullException(fieldName);
            }
            return true;
        }

        private static void Execute(object sender, EventArgs e)
        {
            try
            {
                var project = ProjectHelper.GetSelectedProject();
                project.Save();

                if (!ProjectHelper.IsSupported(project)) return;

                var nuspecFile = ProjectHelper.GetFile(Path.GetDirectoryName(project.FileName), NuspecFileName);

                if (nuspecFile == null)
                {
                    nuspecFile = GenerateNuspecFile(project);
                }
                else
                {
                    UpdateNuspecFile(project, nuspecFile);
                }

                project.ProjectItems.AddFromFile(nuspecFile);
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
                Items = packageReferences.Select(x => new dependency { id = x.Key, version = x.Value }).ToArray<object>()
            };
            return dependencies;
        }

        private static Dictionary<string, string> GetTokens(string path, string[] fileNames)
        {
            var lines = ProjectHelper.GetLines(path, fileNames);
            var tokens = new Dictionary<string, string>();

            foreach (string line in lines)
            {
                if (line.ToLower().Trim().Contains("AssemblyTitle".ToLower())) tokens["id"] = GetValueFromLine(line);
                if (line.ToLower().Trim().Contains("AssemblyInformationalVersion".ToLower())) tokens["version"] = GetValueFromLine(line);
                if (line.ToLower().Trim().Contains("AssemblyVersion".ToLower())) tokens["version"] = GetValueFromLine(line);
                if (line.ToLower().Trim().Contains("AssemblyCompany".ToLower())) tokens["authors"] = GetValueFromLine(line);
                if (line.ToLower().Trim().Contains("AssemblyTitle".ToLower())) tokens["title"] = GetValueFromLine(line);
                if (line.ToLower().Trim().Contains("AssemblyDescription".ToLower())) tokens["description"] = GetValueFromLine(line);
                if (line.ToLower().Trim().Contains("AssemblyCopyright".ToLower())) tokens["copyright"] = GetValueFromLine(line);
            }
            return tokens;
        }

        private static string GetValueFromLine(string line)
        {
            if (line.Contains("(") && line.Contains(")")) return line.Split(')')[0].Split('(')[1].Replace("\"", String.Empty);
            return string.Empty;
        }

        private static XDocument UpdateNamespace(XDocument xDocument, XNamespace xNamespace)
        {
            foreach (var xElement in xDocument.Descendants())
            {
                xElement.SetAttributeValue("xmlns", xNamespace.NamespaceName);
                xElement.Name = xNamespace + xElement.Name.LocalName;
            }

            return xDocument;
        }

        private static void UpdateNuspecFile(Project project, string nuspecFilePath)
        {
            Logger.Log("Updating .nuspec file...");
            var package = UpdatePackage(project, nuspecFilePath);
            GenerateXml(package, nuspecFilePath);
        }

        private static package UpdatePackage(Project project, string nuspecFilePath)
        {
            var xml = XDocument.Load(nuspecFilePath);
            xml = UpdateNamespace(xml, NuspecNamespace);
            var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()));
            var xmlTextReader = new XmlTextReader(memoryStream) { Namespaces = true };
            var xmlSerializer = new XmlSerializer(typeof(package));
            var tokens = GetTokens(Path.GetDirectoryName(project.FileName),
                new[] { Path.GetFileName(project.FileName), ProjectHelper.GetAssemblyFileName(project) });
            var packageReferences = ProjectHelper.GetPackageReferences(project.FileName);
            var dependencies = GetDependencies(packageReferences);
            var packageMetadata = CreatePackageMetadata(tokens, dependencies);
            var package = (package)xmlSerializer.Deserialize(xmlTextReader);
            var updatedPackage = UpdatePackage(package, packageMetadata);

            return updatedPackage;
        }

        private static package UpdatePackage(package package, packageMetadata packageMetadata)
        {
            package.metadata.id = CompareValues(package.metadata.id, packageMetadata.id);
            package.metadata.version = CompareValues(package.metadata.version, packageMetadata.version);
            package.metadata.authors = CompareValues(package.metadata.authors, packageMetadata.authors);
            package.metadata.owners = CompareValues(package.metadata.owners, packageMetadata.owners);
            package.metadata.title = CompareValues(package.metadata.title, packageMetadata.title);
            package.metadata.description = CompareValues(package.metadata.description, packageMetadata.description);
            package.metadata.copyright = CompareValues(package.metadata.copyright, packageMetadata.copyright);

            package.metadata.dependencies =
                packageMetadata?.dependencies?.Items.Length == 0 ? null : packageMetadata?.dependencies;

            return package;
        }

        private static string CompareValues(string oldValue, string newValue)
        {
            if(string.IsNullOrEmpty(oldValue)) return string.Empty;
            if(string.IsNullOrEmpty(newValue))  return string.Empty;
            if (!oldValue.Equals(newValue)) 
                return newValue;
            return oldValue;
        }
    }
}