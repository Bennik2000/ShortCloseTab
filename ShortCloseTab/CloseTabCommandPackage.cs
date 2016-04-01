using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace ShortCloseTab
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class CloseTabCommandPackage : Package
    {
        private const string PackageGuidString = "c754622d-e485-4cea-b88f-effbd966f960";
        
        #region Package Members
        
        protected override void Initialize()
        {
            CloseTabCommand.Initialize(this);
            
            base.Initialize();
        }

        #endregion
    }
}
