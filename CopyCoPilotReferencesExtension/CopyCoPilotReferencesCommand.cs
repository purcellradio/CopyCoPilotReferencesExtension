using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace CopyCoPilotReferencesExtension
{
    internal sealed class CopyCoPilotReferencesCommand
    {
        #region Static

        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("e903358c-f298-4653-a50b-c7853569396f");

        #endregion

        #region Private Fields

        private readonly AsyncPackage package;

        #endregion

        #region Public Properties

        public static CopyCoPilotReferencesCommand Instance { get; private set; }

        #endregion

        #region Constructors

        private CopyCoPilotReferencesCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        #endregion

        #region Event Handlers

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            if(sender is OleMenuCommand myCommand)
            {
                myCommand.Visible = true;
                myCommand.Enabled = true;
            }
        }

        #endregion

        #region Public Methods

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CopyCoPilotReferencesCommand(package, commandService);
        }

        #endregion

        #region Private Methods

        // Lightweight implementation of Path.GetRelativePath for .NET Framework
        private static string GetRelativePath(string basePath, string targetPath)
        {
            if(string.IsNullOrEmpty(basePath) || string.IsNullOrEmpty(targetPath))
            {
                return targetPath ?? string.Empty;
            }

            try
            {
                // Ensure directories end with separator
                if(!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()) && !basePath.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
                {
                    basePath += Path.DirectorySeparatorChar;
                }

                var baseUri = new Uri(basePath, UriKind.Absolute);
                var targetUri = new Uri(targetPath, UriKind.Absolute);

                // Different volume -> cannot make relative
                if(baseUri.Scheme != targetUri.Scheme || !string.Equals(baseUri.Authority, targetUri.Authority, StringComparison.OrdinalIgnoreCase))
                {
                    return targetPath;
                }

                var relativeUri = baseUri.MakeRelativeUri(targetUri);
                var relativePath = Uri.UnescapeDataString(relativeUri.ToString());

                // Convert URI separators to platform separators first, then we will normalize to '/' for Copilot
                relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);
                return relativePath;
            }
            catch
            {
                return targetPath;
            }
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if(Package.GetGlobalService(typeof(DTE)) is DTE2 dte && dte.Solution != null && !string.IsNullOrEmpty(dte.Solution.FullName))
            {
                var uih = dte.ToolWindows.SolutionExplorer;

                if(uih.SelectedItems is Array selectedItems && selectedItems.Length > 0)
                {
                    var solutionDir = Path.GetDirectoryName(dte.Solution.FullName);

                    var filePaths = selectedItems.Cast<UIHierarchyItem>()
                        .Select(item => item.Object as ProjectItem)
                        .Where(projItem => projItem != null && projItem.FileCount > 0)
                        .Select(projItem =>
                            {
                                var fullPath = projItem.FileNames[1];
                                var relativePath = fullPath;

                                if(!string.IsNullOrEmpty(solutionDir))
                                {
                                    // Always compute relative path (will include .. for files outside solution directory)
                                    relativePath = GetRelativePath(solutionDir, fullPath);
                                }

                                // Normalize to forward slashes for Copilot reference format
                                relativePath = relativePath.Replace('\\', '/');
                                return $"#file:'{relativePath}'";
                            }
                        )
                        .ToList();

                    if(filePaths.Any())
                    {
                        Clipboard.SetText(string.Join(" ", filePaths));
                    }
                }
            }
        }

        #endregion
    }
}