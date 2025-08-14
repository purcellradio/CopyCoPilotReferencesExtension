using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace CopyCoPilotReferencesExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class CopyCoPilotReferencesPackage : AsyncPackage
    {
        #region Static

        public const string PackageGuidString = "f5e53b6a-4f5f-4f87-b6d9-167b162f3b8b";

        #endregion

        #region Protected Methods

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await CopyCoPilotReferencesCommand.InitializeAsync(this);
        }

        #endregion
    }
}