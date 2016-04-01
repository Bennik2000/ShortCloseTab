using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;

namespace ShortCloseTab
{
    internal sealed class CloseTabCommand
    {
        private const int CommandId = 0x0100;

        private static readonly Guid CommandSet = new Guid("286f10fc-2b99-49a7-b8f9-04ee335b6bba");
        
        private readonly Package _package;
        
        private CloseTabCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            _package = package;

            OleMenuCommandService commandService = ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandId = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandId);
                commandService.AddCommand(menuItem);
            }
        }

        private IServiceProvider ServiceProvider => _package;

        public static void Initialize(Package package)
        {
            // ReSharper disable once ObjectCreationAsStatement
            new CloseTabCommand(package);
        }
        
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var dte = (DTE2) ServiceProvider.GetService(typeof (DTE));
            try
            {
                dte.ExecuteCommand("close");
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { }
        }
    }
}
