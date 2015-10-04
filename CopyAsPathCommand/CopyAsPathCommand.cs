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

        private CopyAsPathCommand(IServiceProvider serviceProvider, UIHierarchy solutionExplorer)
        {
            Requires.NotNull(serviceProvider, nameof(serviceProvider));
            Requires.NotNull(solutionExplorer, nameof(solutionExplorer));

            OleMenuCommandService commandService = serviceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }

            this.solutionExplorer = solutionExplorer;
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            Array selectedItems = this.solutionExplorer.SelectedItems as Array;
            if (selectedItems != null)
            {
                foreach (UIHierarchyItem item in selectedItems)
                {
                    ProjectItem prjItem = item.Object as ProjectItem;
                    string path = "\"" + prjItem.Properties.Item("FullPath").Value.ToString() + "\"";
                    Clipboard.SetText(path);
                }
            }
        }

        public static CopyAsPathCommand Instance { get; private set; }

        public static void Initialize(IServiceProvider serviceProvider, UIHierarchy solutionExplorer)
        {
            Requires.NotNull(serviceProvider, nameof(serviceProvider));
            Requires.NotNull(solutionExplorer, nameof(solutionExplorer));
            Instance = new CopyAsPathCommand(serviceProvider, solutionExplorer);
        }
    }
}
