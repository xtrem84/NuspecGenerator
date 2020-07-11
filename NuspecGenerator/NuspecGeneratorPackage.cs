using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace NuspecGenerator
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class NuspecGeneratorPackage : AsyncPackage
    {
        public const string PackageGuidString = "1a6861a1-ce0b-40bc-8c9c-dd4a2bbb509a";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            Logger.Initialize(this, "Nuspec Generator");
            await GenerateNuspecCommand.InitializeAsync(this);
        }
    }
}
