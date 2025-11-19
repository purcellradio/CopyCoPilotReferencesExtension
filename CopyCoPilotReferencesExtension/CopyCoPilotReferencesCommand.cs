using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;
using System.Collections.Generic;

namespace CopyCoPilotReferencesExtension
{
    internal sealed class CopyCoPilotReferencesCommand
    {
        #region Static

        public const int CommandId = 0x0100; // Primary command ID

        private static readonly int[] AdditionalCommandIds = new[]
        {
            0x0110,
            0x0111,
            0x0112,
            0x0113,
            0x0114,
            0x0115
        };

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

            RegisterCommand(commandService, CommandId);

            foreach(var id in AdditionalCommandIds)
            {
                RegisterCommand(commandService, id);
            }
        }

        #endregion

        #region Event Handlers

        private void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var cmd = sender as OleMenuCommand;

            if(cmd == null)
            {
                return;
            }

            if(!(Package.GetGlobalService(typeof(DTE)) is DTE2 dte) || dte.Solution == null)
            {
                cmd.Visible = false;
                cmd.Enabled = false;

                return;
            }

            var uih = dte.ToolWindows.SolutionExplorer;

            if(!(uih.SelectedItems is Array selectedItems) || selectedItems.Length == 0)
            {
                cmd.Visible = false;
                cmd.Enabled = false;

                return;
            }

            var hasUsable = selectedItems.Cast<UIHierarchyItem>()
                .Any(item => IsUsableObject(item.Object));

            cmd.Visible = hasUsable;
            cmd.Enabled = hasUsable;
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

        private void RegisterCommand(OleMenuCommandService commandService, int commandId)
        {
            var menuCommandID = new CommandID(CommandSet, commandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        private static bool IsUsableObject(object obj)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if(obj is ProjectItem pi)
            {
                if(string.Equals(
                       pi.Kind,
                       Constants.vsProjectItemKindPhysicalFolder,
                       StringComparison.OrdinalIgnoreCase
                   ))
                {
                    return true;
                }

                if(pi.FileCount > 0)
                {
                    for(short i = 1; i <= pi.FileCount; i++)
                    {
                        var path = pi.FileNames[i];

                        if(!string.IsNullOrEmpty(path) && File.Exists(path))
                        {
                            return true;
                        }
                    }
                }

                return false;
            }

            if(obj is Project)
            {
                return true;
            }

            if(obj is Solution)
            {
                return true;
            }

            return false;
        }

        private static string GetRelativePath(string basePath, string targetPath)
        {
            if(string.IsNullOrEmpty(basePath) || string.IsNullOrEmpty(targetPath))
            {
                return targetPath ?? string.Empty;
            }

            try
            {
                if(!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()) && !basePath.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
                {
                    basePath += Path.DirectorySeparatorChar;
                }

                var baseUri = new Uri(basePath, UriKind.Absolute);
                var targetUri = new Uri(targetPath, UriKind.Absolute);

                if(baseUri.Scheme != targetUri.Scheme || !string.Equals(
                       baseUri.Authority,
                       targetUri.Authority,
                       StringComparison.OrdinalIgnoreCase
                   ))
                {
                    return targetPath;
                }

                var relativeUri = baseUri.MakeRelativeUri(targetUri);
                var relativePath = Uri.UnescapeDataString(relativeUri.ToString());
                relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);

                return relativePath;
            }
            catch
            {
                return targetPath;
            }
        }

        private static void CollectFilePathsFromProjectItem(ProjectItem projectItem, HashSet<string> results)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if(projectItem == null)
            {
                return;
            }

            if(string.Equals(
                   projectItem.Kind,
                   Constants.vsProjectItemKindPhysicalFolder,
                   StringComparison.OrdinalIgnoreCase
               ))
            {
                foreach(ProjectItem child in projectItem.ProjectItems)
                {
                    CollectFilePathsFromProjectItem(child, results);
                }

                return;
            }

            if(projectItem.FileCount > 0)
            {
                for(short i = 1; i <= projectItem.FileCount; i++)
                {
                    var path = projectItem.FileNames[i];

                    if(!string.IsNullOrEmpty(path) && File.Exists(path))
                    {
                        results.Add(path);
                    }
                }
            }

            foreach(ProjectItem child in projectItem.ProjectItems)
            {
                CollectFilePathsFromProjectItem(child, results);
            }
        }

        private static void CollectFilePathsFromProject(Project project, HashSet<string> results)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if(project == null)
            {
                return;
            }

            try
            {
                foreach(ProjectItem item in project.ProjectItems)
                {
                    CollectFilePathsFromProjectItem(item, results);
                }
            }
            catch
            {
            }
        }

        private static IEnumerable<string> GetFilePathsFromHierarchyObject(object hierarchyObject)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var results = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if(hierarchyObject is ProjectItem projectItem)
            {
                CollectFilePathsFromProjectItem(projectItem, results);
            }
            else if(hierarchyObject is Project project)
            {
                CollectFilePathsFromProject(project, results);
            }
            else if(hierarchyObject is Solution solution)
            {
                foreach(Project proj in solution.Projects)
                {
                    CollectFilePathsFromProject(proj, results);
                }
            }

            return results;
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if(!(Package.GetGlobalService(typeof(DTE)) is DTE2 dte) || dte.Solution == null || string.IsNullOrEmpty(dte.Solution.FullName))
            {
                return;
            }

            var uih = dte.ToolWindows.SolutionExplorer;

            if(!(uih.SelectedItems is Array selectedItems) || selectedItems.Length == 0)
            {
                return;
            }

            var solutionDir = Path.GetDirectoryName(dte.Solution.FullName);

            var allFilePaths = selectedItems.Cast<UIHierarchyItem>()
                .SelectMany(item => GetFilePathsFromHierarchyObject(item.Object))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var formattedReferences = allFilePaths.Select(fullPath =>
                    {
                        var relativePath = fullPath;

                        if(!string.IsNullOrEmpty(solutionDir))
                        {
                            relativePath = GetRelativePath(solutionDir, fullPath);
                        }

                        relativePath = relativePath.Replace('\\', '/');

                        return $"#file:'{relativePath}'";
                    }
                )
                .ToList();

            if(formattedReferences.Any())
            {
                Clipboard.SetText(string.Join(" ", formattedReferences));
            }
        }

        #endregion
    }
}