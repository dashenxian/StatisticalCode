using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;
using System.Collections.Generic;
using System.Linq;

namespace RightCmd
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class StatisticalRowsCmd
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("6c2b604e-77be-4f80-9d43-de486d38dd4b");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="StatisticalRowsCmd"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private StatisticalRowsCmd(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static StatisticalRowsCmd Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in StatisticalRowsCmd's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new StatisticalRowsCmd(package, commandService);
            Instance.Dte2 = await package.GetServiceAsync(typeof(DTE)) as DTE2;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var classFiles = GetClassFile();
            var filePaths = classFiles.Select(i => i.FileNames[0]);
            var rowCount = GetRowCount(filePaths);
            string message = string.Format(CultureInfo.CurrentCulture, $"总行数：{rowCount}", this.GetType().FullName);
            string title = "统计";
            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private IEnumerable<ProjectItem> GetClassFile()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var selects = (Array)Dte2.ToolWindows.SolutionExplorer.SelectedItems;
            var files = new List<ProjectItem>();
            foreach (UIHierarchyItem selItem in selects)
            {
                if (selItem.Object is EnvDTE.ProjectItem item)
                {
                    files.Add(item);
                }
            }

            return files;
        }
        /// <summary>
        /// 计算行数
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        private int GetRowCount(IEnumerable<string> filePaths)
        {
            var count = 0;
            foreach (var filePath in filePaths)
            {
                count += GetRowCount(filePath);
            }
            return count;
        }
        /// <summary>
        /// 计算行数
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private int GetRowCount(string filePath)
        {
            var count = 0;
            var lines = System.IO.File.ReadLines(filePath);
            lines = lines.Select(i => i.Trim())
                .Where(i => !i.StartsWith("//") && !string.IsNullOrEmpty(i))
                ;
            count = lines.Count();

            return count;
        }
        public DTE2 Dte2
        {
            get;
            private set;
        }
    }
}
