#region CodeMaid is Copyright 2007-2014 Steve Cadwallader.

// CodeMaid is free software: you can redistribute it and/or modify it under the terms of the GNU
// Lesser General Public License version 3 as published by the Free Software Foundation.
//
// CodeMaid is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
// even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
// Lesser General Public License for more details <http://www.gnu.org/licenses/>.

#endregion CodeMaid is Copyright 2007-2014 Steve Cadwallader.

using System;
using System.Collections.Generic;
using System.Linq;
using EnvDTE;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Outlining;
using Microsoft.VisualStudio.TextManager.Interop;
using SteveCadwallader.CodeMaid.Model.CodeItems;
using IServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace SteveCadwallader.CodeMaid.Logic.Digging
{
    /// <summary>
    /// A manager class for controlling the synchronization of outlining states between the code
    /// document and Spade.
    /// </summary>
    internal class OutliningSynchronizationManager : IDisposable
    {
        #region Fields

        private readonly CodeMaidPackage _package;
        private readonly IVsEditorAdaptersFactoryService _editorAdaptersFactoryService;
        private readonly IOutliningManagerService _outliningManagerService;
        private readonly ServiceProvider _serviceProvider;

        private Document _document;
        private IOutliningManager _outliningManager;
        private IWpfTextView _wpfTextView;

        private IEnumerable<ICodeItemParent> _codeItemParents;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OutliningSynchronizationManager" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        public OutliningSynchronizationManager(CodeMaidPackage package)
        {
            _package = package;

            // Retrieve services needed for outlining from the package.
            _editorAdaptersFactoryService = _package.IComponentModel.GetService<IVsEditorAdaptersFactoryService>();
            _outliningManagerService = _package.IComponentModel.GetService<IOutliningManagerService>();
            _serviceProvider = new ServiceProvider((IServiceProvider)_package.IDE);
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets the document.
        /// </summary>
        public Document Document
        {
            get { return _document; }
            set
            {
                if (_document != value)
                {
                    if (_document != null && _outliningManager != null)
                    {
                        // Unregister from outlining events on the previous document.
                        _outliningManager.RegionsCollapsed -= OnCodeRegionsCollapsed;
                        _outliningManager.RegionsExpanded -= OnCodeRegionsExpanded;
                    }

                    _document = value;
                    _outliningManager = GetOutliningManager(_document);

                    if (_document != null && _outliningManager != null)
                    {
                        // Register for outlining events on the new document.
                        _outliningManager.RegionsCollapsed += OnCodeRegionsCollapsed;
                        _outliningManager.RegionsExpanded += OnCodeRegionsExpanded;
                    }
                }
            }
        }

        #endregion Properties

        #region Public Methods

        /// <summary>
        /// Updates the code items whose outlining is being synchronized by this manager.
        /// </summary>
        /// <param name="codeItems">The code items.</param>
        public void UpdateCodeItems(SetCodeItems codeItems)
        {
            TearDownCodeItemParents();

            // Retrieve and cache an updated list of code item parents.
            _codeItemParents = RecursivelyGetAllCodeItemParents(codeItems);

            InitializeCodeItemParents();
        }

        #endregion Public Methods

        #region Event Handlers

        /// <summary>
        /// An event handler raised when a code item parent's expanded state has changed.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="eventArgs">The event arguments.</param>
        private void OnCodeItemParentIsExpandedChanged(object sender, EventArgs eventArgs)
        {
            var codeItemParent = sender as ICodeItemParent;
            if (codeItemParent != null)
            {
                var iCollapsible = FindCollapsibleFromCodeItemParent(codeItemParent);
                if (iCollapsible != null)
                {
                    if (codeItemParent.IsExpanded && iCollapsible.IsCollapsed)
                    {
                        _outliningManager.Expand(iCollapsible as ICollapsed);
                    }
                    else if (!codeItemParent.IsExpanded && !iCollapsible.IsCollapsed)
                    {
                        _outliningManager.TryCollapse(iCollapsible);
                    }
                }
            }
        }

        /// <summary>
        /// An event handler raised when code region(s) have been collapsed.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event arguments.</param>
        private void OnCodeRegionsCollapsed(object sender, RegionsCollapsedEventArgs e)
        {
            foreach (var collapsedRegion in e.CollapsedRegions)
            {
                var codeItemParent = FindCodeItemParentFromCollapsible(collapsedRegion);
                if (codeItemParent != null)
                {
                    codeItemParent.IsExpanded = false;
                }
            }
        }

        /// <summary>
        /// An event handler raised when code region(s) have been expanded.
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The event arguments.</param>
        private void OnCodeRegionsExpanded(object sender, RegionsExpandedEventArgs e)
        {
            foreach (var expandedRegion in e.ExpandedRegions)
            {
                var codeItemParent = FindCodeItemParentFromCollapsible(expandedRegion);
                if (codeItemParent != null)
                {
                    codeItemParent.IsExpanded = true;
                }
            }
        }

        #endregion Event Handlers

        #region Helper Methods

        /// <summary>
        /// Attempts to find a <see cref="ICodeItemParent" /> associated with the specified <see
        /// cref="ICollapsible" />.
        /// </summary>
        /// <param name="collapsible">The collapsible region.</param>
        /// <returns>
        /// The <see cref="ICodeItemParent" /> on the same starting line, otherwise null.
        /// </returns>
        private ICodeItemParent FindCodeItemParentFromCollapsible(ICollapsible collapsible)
        {
            var startLine = GetStartLineForCollapsible(collapsible);

            return _codeItemParents.FirstOrDefault(x => x.StartLine == startLine);
        }

        /// <summary>
        /// Attempts to find a <see cref="ICollapsible" /> associated with the specified <see
        /// cref="ICodeItemParent" />.
        /// </summary>
        /// <param name="parent">The code item parent.</param>
        /// <returns>The <see cref="ICollapsible" /> on the same starting line, otherwise null.</returns>
        private ICollapsible FindCollapsibleFromCodeItemParent(ICodeItemParent parent)
        {
            if (_outliningManager == null || _wpfTextView == null)
            {
                return null;
            }

            var snapshotLine = _wpfTextView.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(parent.StartLine);
            var collapsibles = _outliningManager.GetAllRegions(snapshotLine.Extent);

            return (from collapsible in collapsibles
                    let startLine = GetStartLineForCollapsible(collapsible)
                    where startLine == parent.StartLine
                    select collapsible).FirstOrDefault();
        }

        /// <summary>
        /// Gets the start line for the specified collapsible.
        /// </summary>
        /// <remarks>
        /// The +1 offset is to accomdate for the 0-based code indexing vs. 1-based code item indexing.
        /// </remarks>
        /// <param name="collapsible">The collapsible region.</param>
        /// <returns>The starting line.</returns>
        private static int GetStartLineForCollapsible(ICollapsible collapsible)
        {
            var startPoint = collapsible.Extent.GetStartPoint(collapsible.Extent.TextBuffer.CurrentSnapshot);
            var line = startPoint.Snapshot.GetLineNumberFromPosition(startPoint.Position) + 1;

            return line;
        }

        /// <summary>
        /// Attempts to get the outlining manager associated with the specified document.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <returns>The associated outlining manager, otherwise null.</returns>
        private IOutliningManager GetOutliningManager(Document document)
        {
            _wpfTextView = GetWpfTextView(document);
            if (_wpfTextView != null && _outliningManagerService != null)
            {
                return _outliningManagerService.GetOutliningManager(_wpfTextView);
            }

            return null;
        }

        /// <summary>
        /// Attempts to get the WPF text view associated with the specified document.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <returns>The associated WPF text view, otherwise null.</returns>
        private IWpfTextView GetWpfTextView(Document document)
        {
            var textView = GetTextView(document);
            if (textView != null && _editorAdaptersFactoryService != null)
            {
                return _editorAdaptersFactoryService.GetWpfTextView(textView);
            }

            return null;
        }

        /// <summary>
        /// Attempts to get the text view associated with the specified document.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <returns>The associated text view, otherwise null.</returns>
        private IVsTextView GetTextView(Document document)
        {
            if (_serviceProvider == null || document == null)
            {
                return null;
            }

            IVsUIHierarchy hierarchy;
            uint itemID;
            IVsWindowFrame windowFrame;

            if (VsShellUtilities.IsDocumentOpen(_serviceProvider, document.FullName, Guid.Empty, out hierarchy, out itemID, out windowFrame))
            {
                return VsShellUtilities.GetTextView(windowFrame);
            }

            return null;
        }

        /// <summary>
        /// Recursively retrives all code item parents within the specified code items.
        /// </summary>
        /// <param name="codeItems">The code items.</param>
        /// <returns>The code item parents.</returns>
        private static IEnumerable<ICodeItemParent> RecursivelyGetAllCodeItemParents(SetCodeItems codeItems)
        {
            if (codeItems == null)
            {
                return Enumerable.Empty<ICodeItemParent>();
            }

            var parents = codeItems.OfType<ICodeItemParent>().ToList();

            return parents.Union(parents.SelectMany(x => RecursivelyGetAllCodeItemParents(x.Children)));
        }

        /// <summary>
        /// Initializes the code item parents by synchronizing their current state and registering
        /// for events.
        /// </summary>
        private void InitializeCodeItemParents()
        {
            foreach (var codeItemParent in _codeItemParents ?? Enumerable.Empty<ICodeItemParent>())
            {
                var iCollapsible = FindCollapsibleFromCodeItemParent(codeItemParent);
                if (iCollapsible != null)
                {
                    codeItemParent.IsExpanded = !iCollapsible.IsCollapsed;
                }

                // Register for expansion changes on the code item parent.
                codeItemParent.IsExpandedChanged += OnCodeItemParentIsExpandedChanged;
            }
        }

        /// <summary>
        /// Tears down the code item parents by unregistering from events.
        /// </summary>
        private void TearDownCodeItemParents()
        {
            foreach (var codeItemParent in _codeItemParents ?? Enumerable.Empty<ICodeItemParent>())
            {
                codeItemParent.IsExpandedChanged -= OnCodeItemParentIsExpandedChanged;
            }

            _codeItemParents = null;
        }

        #endregion Helper Methods

        #region Implementation of IDisposable

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting
        /// unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            TearDownCodeItemParents();

            Document = null;
        }

        #endregion Implementation of IDisposable
    }
}