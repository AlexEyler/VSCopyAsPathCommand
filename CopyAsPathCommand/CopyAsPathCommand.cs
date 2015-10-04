using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft;
using EnvDTE;
using System.Windows;

namespace CopyAsPathCommand
{
    internal sealed class CopyAsPathCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("cceebeea-e540-447e-8b6d-ffec0ffd311d");
        private UIHierarchy solutionExplorer;
        private CopyAsPathCommandPackage package;

        private CopyAsPathCommand(CopyAsPathCommandPackage package, UIHierarchy solutionExplorer)
        {
            Requires.NotNull(package, nameof(package));
            Requires.NotNull(solutionExplorer, nameof(solutionExplorer));
            this.package = package;
            this.solutionExplorer = solutionExplorer;

            OleMenuCommandService commandService = package.QueryService<IMenuCommandService>() as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            var optionsPage = package.GetDialogPage(typeof(OptionsPage)) as OptionsPage;
            bool useQuotes = optionsPage?.QuotesOption ?? OptionsPage.DefaultQuotesOption;

            Array selectedItems = this.solutionExplorer.SelectedItems as Array;
            if (selectedItems != null)
            {
                foreach (UIHierarchyItem item in selectedItems)
                {
                    ProjectItem prjItem = item.Object as ProjectItem;
                    string path = prjItem.Properties.Item("FullPath").Value.ToString();
                    if (useQuotes)
                    {
                        path = "\"" + path + "\"";
                    }

                    Clipboard.SetText(path);
                }
            }
        }

        public static CopyAsPathCommand Instance { get; private set; }

        public static void Initialize(CopyAsPathCommandPackage package, UIHierarchy solutionExplorer)
        {
            Requires.NotNull(package, nameof(package));
            Requires.NotNull(solutionExplorer, nameof(solutionExplorer));
            Instance = new CopyAsPathCommand(package, solutionExplorer);
        }
    }
}
