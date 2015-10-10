using System;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CopyAsPathCommand
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [InstalledProductRegistration(CopyAsPathCommandPackage.ProductName,
        CopyAsPathCommandPackage.ProductDescription, CopyAsPathCommandPackage.Version)]
    [Guid(CopyAsPathCommandPackage.PackageGuidString)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists)]
    [ProvideOptionPage(typeof(OptionsPage), "Environment", "Copy as Path Command", 0, 0, true)]
    public sealed class CopyAsPathCommandPackage : Package
    {
        public const string Version = "1.0.6";
        public const string ProductName = "Copy as path command";
        public const string ProductDescription = "Right click on a solution explorer item and get its absolute path.";
        public const string PackageGuidString = "d40a488d-23d0-44cc-99fb-f6dd0717ab7d";

        #region Package Members
        protected override void Initialize()
        {
            DTE2 envDte = this.GetService(typeof(DTE)) as DTE2;
            UIHierarchy solutionWindow = envDte.ToolWindows.SolutionExplorer;
            CopyAsPathCommand.Initialize(this, solutionWindow);
            base.Initialize();
        }
        #endregion
    }
}
