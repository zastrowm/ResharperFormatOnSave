#region CodeMaid is Copyright 2007-2014 Steve Cadwallader.

// CodeMaid is free software: you can redistribute it and/or modify it under the terms of the GNU
// Lesser General Public License version 3 as published by the Free Software Foundation.
//
// CodeMaid is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
// even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
// Lesser General Public License for more details <http://www.gnu.org/licenses/>.

#endregion CodeMaid is Copyright 2007-2014 Steve Cadwallader.

using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using SteveCadwallader.CodeMaid.Integration;

namespace SteveCadwallader.CodeMaid.UI.ToolWindows.BuildProgress
{
    /// <summary>
    /// This class implements the build progress tool window and hosts a user control.
    /// </summary>
    [Guid(GuidList.GuidCodeMaidToolWindowBuildProgressString)]
    public class BuildProgressToolWindow : ToolWindowPane
    {
        #region Fields

        private const string DefaultCaption = "Build Progress";

        private readonly BuildProgressViewModel _viewModel;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="BuildProgressToolWindow" /> class.
        /// </summary>
        public BuildProgressToolWindow()
            : base(null)
        {
            // Set the tool window caption.
            Caption = DefaultCaption;

            // Set the tool window image from resources.
            BitmapResourceID = 502;
            BitmapIndex = 0;

            // Create the view model.
            _viewModel = new BuildProgressViewModel();

            // Set the tool window content.
            base.Content = new BuildProgressView { DataContext = _viewModel };
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets the last known build action.
        /// </summary>
        private vsBuildAction BuildAction { get; set; }

        /// <summary>
        /// Gets or sets the last known build scope.
        /// </summary>
        private vsBuildScope BuildScope { get; set; }

        /// <summary>
        /// Gets or sets the number of projects built.
        /// </summary>
        private int NumberOfProjectsBuilt { get; set; }

        /// <summary>
        /// Gets or sets the number of projects to be built.
        /// </summary>
        private int NumberOfProjectsToBeBuilt { get; set; }

        /// <summary>
        /// Gets or sets the package that owns the tool window.
        /// </summary>
        private new CodeMaidPackage Package { get { return base.Package as CodeMaidPackage; } }

        /// <summary>
        /// Gets the progress percentage, otherwise zero if cannot be determined.
        /// </summary>
        private double ProgressPercentage
        {
            get
            {
                double toBeBuilt = NumberOfProjectsToBeBuilt;

                return toBeBuilt > 0
                           ? NumberOfProjectsBuilt / toBeBuilt
                           : 0;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// This method can be overriden by the derived class to execute any code that needs to run
        /// after the IVsWindowFrame is created. If the toolwindow has a toolbar with a combobox, it
        /// should make sure its command handler are set by the time they return from this method.
        /// This is called when someone set the Frame property.
        /// </summary>
        public override void OnToolWindowCreated()
        {
            base.OnToolWindowCreated();

            // Pass the package over to the view model, not available during constructor.
            _viewModel.Package = Package;
        }

        /// <summary>
        /// A method called to notify the tool window that a build has begun.
        /// </summary>
        /// <param name="scope">The scope.</param>
        /// <param name="action">The action.</param>
        internal void NotifyBuildBegin(vsBuildScope scope, vsBuildAction action)
        {
            BuildAction = action;
            BuildScope = scope;
            NumberOfProjectsBuilt = 0;

            if (BuildScope == vsBuildScope.vsBuildScopeSolution)
            {
                NumberOfProjectsToBeBuilt = GetNumberOfProjectsToBeBuilt();
                _viewModel.IsProgressIndeterminate = false;
            }
            else
            {
                NumberOfProjectsToBeBuilt = 0;
                _viewModel.IsProgressIndeterminate = true;
            }

            Caption = DefaultCaption;
            _viewModel.HasBuildFailed = false;
            _viewModel.IsBuildActive = true;
            _viewModel.ProgressPercentage = 0;
        }

        /// <summary>
        /// A method called to notify the tool window that an individual project build has begun.
        /// </summary>
        /// <param name="project">The project.</param>
        /// <param name="projectConfig">The project config.</param>
        /// <param name="platform">The platform.</param>
        /// <param name="solutionConfig">The solution config.</param>
        internal void NotifyBuildProjConfigBegin(string project, string projectConfig, string platform, string solutionConfig)
        {
            string projectName = ExtractProjectName(project);
            string buildString = GetBuildTypeString(BuildScope, BuildAction);

            string progressString = string.Empty;
            if (NumberOfProjectsToBeBuilt > 0)
            {
                string projectsString = NumberOfProjectsToBeBuilt.ToString(CultureInfo.CurrentUICulture);
                string completeString = (++NumberOfProjectsBuilt).ToString(CultureInfo.CurrentUICulture).PadLeft(projectsString.Length);

                progressString = string.Format(" {0} of {1}", completeString, projectsString);
            }

            Caption = String.Format("{0}: {1}{2} \"{3}\"...",
                                    DefaultCaption, buildString, progressString, projectName);
            _viewModel.ProgressPercentage = ProgressPercentage;
        }

        /// <summary>
        /// A method called to notify the the tool window that an individual project build is done.
        /// </summary>
        /// <param name="project">The project.</param>
        /// <param name="projectConfig">The project config.</param>
        /// <param name="platform">The platform.</param>
        /// <param name="solutionConfig">The solution config.</param>
        /// <param name="success">True if project build was successful, otherwise false.</param>
        internal void NotifyBuildProjConfigDone(string project, string projectConfig, string platform, string solutionConfig, bool success)
        {
            if (!success)
            {
                _viewModel.HasBuildFailed = true;
            }
        }

        /// <summary>
        /// A method called to notify the tool window that a build is done.
        /// </summary>
        /// <param name="scope">The scope.</param>
        /// <param name="action">The action.</param>
        internal void NotifyBuildDone(vsBuildScope scope, vsBuildAction action)
        {
            Caption = DefaultCaption;
            _viewModel.HasBuildFailed = false;
            _viewModel.IsBuildActive = false;
            _viewModel.IsProgressIndeterminate = false;
            _viewModel.ProgressPercentage = 0;
        }

        /// <summary>
        /// Extracts the project name from the specified project string.
        /// </summary>
        /// <param name="project">The raw project string to process.</param>
        /// <returns>The extracted project name.</returns>
        private static string ExtractProjectName(string project)
        {
            int lastBackslash = project.LastIndexOf('\\') + 1;
            int lastPeriod = project.LastIndexOf('.') + 1;

            return project.Substring(lastBackslash, lastPeriod - lastBackslash - 1);
        }

        /// <summary>
        /// Gets the build type string from the specified build scope and action.
        /// </summary>
        /// <param name="buildScope">The build scope.</param>
        /// <param name="buildAction">The build action.</param>
        /// <returns>The build type string.</returns>
        private static string GetBuildTypeString(vsBuildScope buildScope, vsBuildAction buildAction)
        {
            StringBuilder stringBuilder = new StringBuilder();

            // First append the word 'Batch ' if this is a batch build event.
            if (buildScope == vsBuildScope.vsBuildScopeBatch)
            {
                stringBuilder.Append("Batch ");
            }

            // Next append the action-specific build string.
            switch (buildAction)
            {
                case vsBuildAction.vsBuildActionBuild:
                    stringBuilder.Append("Building");
                    break;

                case vsBuildAction.vsBuildActionClean:
                    stringBuilder.Append("Cleaning");
                    break;

                case vsBuildAction.vsBuildActionDeploy:
                    stringBuilder.Append("Deploying");
                    break;

                case vsBuildAction.vsBuildActionRebuildAll:
                    stringBuilder.Append("Rebuilding");
                    break;
            }

            return stringBuilder.ToString();
        }

        /// <summary>
        /// Gets the number of projects to be built based on the active solution configuration.
        /// </summary>
        private int GetNumberOfProjectsToBeBuilt()
        {
            var solutionContexts = Package.IDE.Solution.SolutionBuild.ActiveConfiguration.SolutionContexts;
            int count = 0;

            for (int i = 0; i < solutionContexts.Count; i++)
            {
                try
                {
                    if (solutionContexts.Item(i + 1).ShouldBuild)
                    {
                        count++;
                    }
                }
                catch (ArgumentException)
                {
                    // This is a work-around for a known issue with the
                    // SolutionContexts.GetEnumerator with unloaded projects in VS2010.
                }
            }

            return count;
        }

        #endregion Methods
    }
}