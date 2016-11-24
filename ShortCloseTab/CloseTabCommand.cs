using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using ShortCloseTab.Utils;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Reflection;

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
                
                ImportSettings();
            }
        }

        private IServiceProvider ServiceProvider => _package;

        public static void Initialize(Package package)
        {
            // ReSharper disable once ObjectCreationAsStatement
            new CloseTabCommand(package);
        }



        private void ImportSettings()
        {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ShortCloseTab_shortcut_settings.vssettings");

            ExportResource.Export("ShortCloseTab.shortcut_settings.vssettings", path, Assembly.GetAssembly(GetType()));


            var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
            try
            {
                // Save the currently loaded scheme name
                var schemeNameProperty = dte.Properties["Environment", "Keyboard"];
                var schemeName = schemeNameProperty.Item("SchemeName").Value;
                
                // Replace the scheme name in the file
                ExportResource.ReplaceInFile(path, "%{Scheme}%", schemeName as string);


                // Load the settings
                dte.ExecuteCommand("Tools.ImportandExportSettings", $"/import:\"{path}\"");
            }
            // ReSharper disable once EmptyGeneralCatchClause
            catch { }

            File.Delete(path);
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
