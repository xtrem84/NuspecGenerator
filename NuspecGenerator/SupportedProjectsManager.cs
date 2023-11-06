using System;
using System.Collections.Generic;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace NuspecGenerator
{
    public class SupportedProjectsManager
    {
        private static SupportedProjectsManager _instance;
        private  List<SupportedProject> _supportedProjects;
        public  IReadOnlyList<SupportedProject> SupportedProjects => _supportedProjects;
        private SupportedProjectsManager()
        {
            _supportedProjects = new List<SupportedProject>();
            _supportedProjects.Add(new SupportedProject("C# .NetFramework Project", "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}", "AssemblyInfo.cs"));
            _supportedProjects.Add(new SupportedProject("VB.Net .NetFramework Project","{F184B08F-C81C-45F6-A57F-5ABD9991F28F}","AssemblyInfo.vb"));

            foreach(var project in _supportedProjects)
            {
                Logger.Log($"Supported Projects {project.Name}  project.Kind:{project.Guid}");
            }
        }
        public static SupportedProjectsManager GetInstance()
        {
            if (_instance == null) _instance = new SupportedProjectsManager();
            return _instance;
        }
        public string GetAssemblyFileNameByProject(Project project)
        {
            string filename = _supportedProjects.Find(y =>
            y.Guid.Equals(project.Kind, StringComparison.InvariantCultureIgnoreCase))?.AssemblyInfoFileName;
            return filename;
        }
        public bool CheckProjectIsSupported(Project project)
        {
            var result = _supportedProjects.Exists(x => x.Guid.Equals(project.Kind,StringComparison.InvariantCultureIgnoreCase));
            return result;
        }
    }
}