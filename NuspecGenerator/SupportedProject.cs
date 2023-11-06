namespace NuspecGenerator
{
    public class SupportedProject
    {
        public SupportedProject(string name, string guid, string assemblyInfoFileName)
        {
            Name = name;
            Guid = guid;
            AssemblyInfoFileName = assemblyInfoFileName;
        }

        public string Name { get; }
        public string Guid { get; }
        public string AssemblyInfoFileName { get; }
    }
}