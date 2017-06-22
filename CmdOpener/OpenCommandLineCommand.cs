//------------------------------------------------------------------------------
// <copyright file="OpenCommandLine.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.IO;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace CommandLineOpener
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class OpenCommandLineCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("b2cb2c39-5a56-4830-976d-a281951730e1");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenCommandLineCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private OpenCommandLineCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static OpenCommandLineCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new OpenCommandLineCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            try
            {
                var fullPath = GetSourceFilePath();

                if (string.IsNullOrEmpty(fullPath) == false)
                {
                    var directoryPath = Path.GetDirectoryName(fullPath);

                    var startInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        WorkingDirectory = directoryPath,
                        FileName = "cmd",
                        UseShellExecute = false,
                    };

                    System.Diagnostics.Process.Start(startInfo);
                }
            }
            catch
            {
                // ignored
            }
        }

        private static EnvDTE80.DTE2 GetDTE2()
        {
            return Package.GetGlobalService(typeof(DTE)) as EnvDTE80.DTE2;
        }

        private string GetSourceFilePath()
        {
            EnvDTE80.DTE2 _applicationObject = GetDTE2();
            UIHierarchy uih = _applicationObject.ToolWindows.SolutionExplorer;
            Array selectedItems = (Array)uih.SelectedItems;
            if (null != selectedItems)
            {
                foreach (UIHierarchyItem selItem in selectedItems)
                {
                    ProjectItem projectItem = selItem.Object as ProjectItem;
                    string filePath = string.Empty;
                    if (projectItem != null)
                    {
                        filePath = projectItem.Properties.Item("FullPath").Value.ToString();
                    }

                    Project project = selItem.Object as Project;
                    
                    if (project != null)
                    {
                        //If the folder is virtual folder, it will throw an exception.
                        filePath = project.Properties.Item("FullPath").Value.ToString();
                    }

                    Solution solution = selItem.Object as Solution;

                    if (solution != null)
                    {
                        filePath = solution.Properties.Item("Path").Value.ToString();
                    }

                    //System.Windows.Forms.MessageBox.Show(selItem.Name + filePath);
                    return filePath;
                }
            }
            return string.Empty;
        }
    }
}
