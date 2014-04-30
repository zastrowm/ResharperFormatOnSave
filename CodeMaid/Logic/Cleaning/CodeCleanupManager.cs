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
using SteveCadwallader.CodeMaid.Helpers;
using SteveCadwallader.CodeMaid.Logic.Reorganizing;
using SteveCadwallader.CodeMaid.Model;
using SteveCadwallader.CodeMaid.Model.CodeItems;
using SteveCadwallader.CodeMaid.Properties;

namespace SteveCadwallader.CodeMaid.Logic.Cleaning
{
    /// <summary>
    /// A manager class for cleaning up code.
    /// </summary>
    /// <remarks>
    ///
    /// Note: All POSIXRegEx text replacements search against '\n' but insert/replace with
    ///       Environment.NewLine. This handles line endings correctly.
    /// </remarks>
    internal class CodeCleanupManager
    {
        #region Fields

        private readonly CodeMaidPackage _package;

        private readonly CodeModelManager _codeModelManager;
        private readonly CodeReorderManager _codeReorderManager;
        private readonly UndoTransactionHelper _undoTransactionHelper;

        private readonly CodeCleanupAvailabilityLogic _codeCleanupAvailabilityLogic;
        private readonly CommentFormatLogic _commentFormatLogic;
        private readonly InsertBlankLinePaddingLogic _insertBlankLinePaddingLogic;
        private readonly InsertExplicitAccessModifierLogic _insertExplicitAccessModifierLogic;
        private readonly InsertWhitespaceLogic _insertWhitespaceLogic;
        private readonly RemoveWhitespaceLogic _removeWhitespaceLogic;
        private readonly UpdateLogic _updateLogic;
        private readonly UsingStatementCleanupLogic _usingStatementCleanupLogic;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// The singleton instance of the <see cref="CodeCleanupManager" /> class.
        /// </summary>
        private static CodeCleanupManager _instance;

        /// <summary>
        /// Gets an instance of the <see cref="CodeCleanupManager" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        /// <returns>An instance of the <see cref="CodeCleanupManager" /> class.</returns>
        internal static CodeCleanupManager GetInstance(CodeMaidPackage package)
        {
            return _instance ?? (_instance = new CodeCleanupManager(package));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CodeCleanupManager" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        private CodeCleanupManager(CodeMaidPackage package)
        {
            _package = package;

            _codeModelManager = CodeModelManager.GetInstance(_package);
            _codeReorderManager = CodeReorderManager.GetInstance(_package);
            _undoTransactionHelper = new UndoTransactionHelper(_package, "CodeMaid Cleanup");

            _codeCleanupAvailabilityLogic = CodeCleanupAvailabilityLogic.GetInstance(_package);
            _commentFormatLogic = CommentFormatLogic.GetInstance(_package);
            _insertBlankLinePaddingLogic = InsertBlankLinePaddingLogic.GetInstance(_package);
            _insertExplicitAccessModifierLogic = InsertExplicitAccessModifierLogic.GetInstance();
            _insertWhitespaceLogic = InsertWhitespaceLogic.GetInstance(_package);
            _removeWhitespaceLogic = RemoveWhitespaceLogic.GetInstance(_package);
            _updateLogic = UpdateLogic.GetInstance(_package);
            _usingStatementCleanupLogic = UsingStatementCleanupLogic.GetInstance(_package);
        }

        #endregion Constructors

        #region Internal Methods

        /// <summary>
        /// Attempts to run code cleanup on the specified project item.
        /// </summary>
        /// <param name="projectItem">The project item for cleanup.</param>
        internal void Cleanup(ProjectItem projectItem)
        {
            if (!_codeCleanupAvailabilityLogic.ShouldCleanup(projectItem)) return;

            // Attempt to open the document if not already opened.
            bool wasOpen = projectItem.IsOpen[Constants.vsViewKindTextView];
            if (!wasOpen)
            {
                try
                {
                    projectItem.Open(Constants.vsViewKindTextView);
                }
                catch (Exception)
                {
                    // OK if file cannot be opened (ex: deleted from disk, non-text based type.)
                }
            }

            if (projectItem.Document != null)
            {
                Cleanup(projectItem.Document, false);

                // Close the document if it was opened for cleanup.
                if (Settings.Default.Cleaning_AutoSaveAndCloseIfOpenedByCleanup && !wasOpen)
                {
                    projectItem.Document.Close(vsSaveChanges.vsSaveChangesYes);
                }
            }
        }

        /// <summary>
        /// Attempts to run code cleanup on the specified document.
        /// </summary>
        /// <param name="document">The document for cleanup.</param>
        /// <param name="isAutoSave">A flag indicating if occurring due to auto-save.</param>
        internal void Cleanup(Document document, bool isAutoSave = false)
        {
            if (!_codeCleanupAvailabilityLogic.ShouldCleanup(document, true)) return;

            // Make sure the document to be cleaned up is active, required for some commands like
            // format document.
            document.Activate();

            if (_package.IDE.ActiveDocument != document)
            {
                OutputWindowHelper.WriteLine(document.Name + " did not complete activation before cleaning started.");
            }

            // Conditionally start cleanup with reorganization.
            if (Settings.Default.Reorganizing_RunAtStartOfCleanup)
            {
                _codeReorderManager.Reorganize(document, isAutoSave);
            }

            _undoTransactionHelper.Run(
                () => !(isAutoSave && Settings.Default.General_SkipUndoTransactionsDuringAutoCleanupOnSave),
                delegate
                {
                    var cleanupMethod = FindCodeCleanupMethod(document);
                    if (cleanupMethod != null)
                    {
                        _package.IDE.StatusBar.Text = String.Format("CodeMaid is cleaning '{0}'...", document.Name);

                        // Perform the set of configured cleanups based on the language.
                        cleanupMethod(document, isAutoSave);

                        _package.IDE.StatusBar.Text = String.Format("CodeMaid cleaned '{0}'.", document.Name);
                    }
                },
                delegate(Exception ex)
                {
                    OutputWindowHelper.WriteLine(String.Format("CodeMaid stopped cleaning '{0}': {1}", document.Name, ex));
                    _package.IDE.StatusBar.Text = String.Format("CodeMaid stopped cleaning '{0}'.  See output window for more details.", document.Name);
                });
        }

        #endregion Internal Methods

        #region Private Language Methods

        /// <summary>
        /// Finds a code cleanup method appropriate for the specified document, otherwise null.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <returns>The code cleanup method, otherwise null.</returns>
        private Action<Document, bool> FindCodeCleanupMethod(Document document)
        {
            switch (document.Language)
            {
                case "CSharp":
                    return RunCodeCleanupCSharp;

                case "C/C++":
                case "CSS":
                case "JavaScript":
                case "JScript":
                case "LESS":
                case "TypeScript":
                    return RunCodeCleanupC;

                case "HTML":
                case "HTMLX":
                case "XAML":
                case "XML":
                    return RunCodeCleanupMarkup;

                case "Basic":
                case "F#":
                    return RunCodeCleanupGeneric;

                default:
                    OutputWindowHelper.WriteLine(String.Format(
                        "CodeMaid does not support document language '{0}'.", document.Language));
                    return null;
            }
        }

        /// <summary>
        /// Attempts to run code cleanup on the specified CSharp document.
        /// </summary>
        /// <param name="document">The document for cleanup.</param>
        /// <param name="isAutoSave">A flag indicating if occurring due to auto-save.</param>
        private void RunCodeCleanupCSharp(Document document, bool isAutoSave)
        {
            var textDocument = (TextDocument)document.Object("TextDocument");
            bool isExternal = _codeCleanupAvailabilityLogic.IsDocumentExternal(document);

            // Perform any actions that can modify the file code model first.
            RunExternalFormatting(textDocument);
            if (!isExternal)
            {
                _usingStatementCleanupLogic.RemoveUnusedUsingStatements(textDocument, isAutoSave);
                _usingStatementCleanupLogic.SortUsingStatements(isAutoSave);
            }

            // Interpret the document into a collection of elements.
            var codeItems = _codeModelManager.RetrieveAllCodeItems(document);

            var regions = codeItems.OfType<CodeItemRegion>().ToList();
            var usingStatements = codeItems.OfType<CodeItemUsingStatement>().ToList();
            var namespaces = codeItems.OfType<CodeItemNamespace>().ToList();
            var classes = codeItems.OfType<CodeItemClass>().ToList();
            var delegates = codeItems.OfType<CodeItemDelegate>().ToList();
            var enumerations = codeItems.OfType<CodeItemEnum>().ToList();
            var events = codeItems.OfType<CodeItemEvent>().ToList();
            var fields = codeItems.OfType<CodeItemField>().ToList();
            var interfaces = codeItems.OfType<CodeItemInterface>().ToList();
            var methods = codeItems.OfType<CodeItemMethod>().ToList();
            var properties = codeItems.OfType<CodeItemProperty>().ToList();
            var structs = codeItems.OfType<CodeItemStruct>().ToList();

            // Build up more complicated collections.
            var usingStatementBlocks = CodeModelHelper.GetCodeItemBlocks(usingStatements).ToList();
            var usingStatementsThatStartBlocks = (from IEnumerable<CodeItemUsingStatement> block in usingStatementBlocks select block.First()).ToList();
            var usingStatementsThatEndBlocks = (from IEnumerable<CodeItemUsingStatement> block in usingStatementBlocks select block.Last()).ToList();
            var fieldsWithComments = fields.Where(x => x.StartPoint.Line < x.EndPoint.Line).ToList();

            // Perform removal cleanup.
            _removeWhitespaceLogic.RemoveEOLWhitespace(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAtTop(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAtBottom(textDocument);
            _removeWhitespaceLogic.RemoveEOFTrailingNewLine(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAfterAttributes(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAfterOpeningBrace(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesBeforeClosingBrace(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesBetweenChainedStatements(textDocument);
            _removeWhitespaceLogic.RemoveMultipleConsecutiveBlankLines(textDocument);

            // Perform insertion of blank line padding cleanup.
            _insertBlankLinePaddingLogic.InsertPaddingBeforeRegionTags(regions);
            _insertBlankLinePaddingLogic.InsertPaddingAfterRegionTags(regions);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeEndRegionTags(regions);
            _insertBlankLinePaddingLogic.InsertPaddingAfterEndRegionTags(regions);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(usingStatementsThatStartBlocks);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(usingStatementsThatEndBlocks);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(namespaces);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(namespaces);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(classes);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(classes);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(delegates);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(delegates);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(enumerations);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(enumerations);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(events);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(events);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(fieldsWithComments);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(fieldsWithComments);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(interfaces);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(interfaces);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(methods);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(methods);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(properties);
            _insertBlankLinePaddingLogic.InsertPaddingBetweenMultiLinePropertyAccessors(properties);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(properties);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCodeElements(structs);
            _insertBlankLinePaddingLogic.InsertPaddingAfterCodeElements(structs);

            _insertBlankLinePaddingLogic.InsertPaddingBeforeCaseStatements(textDocument);
            _insertBlankLinePaddingLogic.InsertPaddingBeforeSingleLineComments(textDocument);

            // Perform insertion of explicit access modifier cleanup.
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnClasses(classes);
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnDelegates(delegates);
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnEnumerations(enumerations);
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnEvents(events);
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnFields(fields);
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnInterfaces(interfaces);
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnMethods(methods);
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnProperties(properties);
            _insertExplicitAccessModifierLogic.InsertExplicitAccessModifiersOnStructs(structs);

            // Perform update cleanup.
            _updateLogic.UpdateEndRegionDirectives(textDocument);
            _updateLogic.UpdateEventAccessorsToBothBeSingleLineOrMultiLine(events);
            _updateLogic.UpdatePropertyAccessorsToBothBeSingleLineOrMultiLine(properties);
            _updateLogic.UpdateSingleLineMethods(methods);

            // Perform comment cleaning.
            _commentFormatLogic.FormatComments(textDocument);
        }

        /// <summary>
        /// Attempts to run code cleanup on the specified C/C++ document.
        /// </summary>
        /// <param name="document">The document for cleanup.</param>
        /// <param name="isAutoSave">A flag indicating if occurring due to auto-save.</param>
        private void RunCodeCleanupC(Document document, bool isAutoSave)
        {
            var textDocument = (TextDocument)document.Object("TextDocument");

            RunExternalFormatting(textDocument);

            // Perform removal cleanup.
            _removeWhitespaceLogic.RemoveEOLWhitespace(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAtTop(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAtBottom(textDocument);
            _removeWhitespaceLogic.RemoveEOFTrailingNewLine(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAfterOpeningBrace(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesBeforeClosingBrace(textDocument);
            _removeWhitespaceLogic.RemoveMultipleConsecutiveBlankLines(textDocument);
        }

        /// <summary>
        /// Attempts to run code cleanup on the specified markup document.
        /// </summary>
        /// <param name="document">The document for cleanup.</param>
        /// <param name="isAutoSave">A flag indicating if occurring due to auto-save.</param>
        private void RunCodeCleanupMarkup(Document document, bool isAutoSave)
        {
            var textDocument = (TextDocument)document.Object("TextDocument");

            RunExternalFormatting(textDocument);

            // Perform removal cleanup.
            _removeWhitespaceLogic.RemoveEOLWhitespace(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAtTop(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAtBottom(textDocument);
            _removeWhitespaceLogic.RemoveEOFTrailingNewLine(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesBeforeClosingTag(textDocument);
            _removeWhitespaceLogic.RemoveBlankSpacesBeforeClosingAngleBracket(textDocument);
            _removeWhitespaceLogic.RemoveMultipleConsecutiveBlankLines(textDocument);

            // Perform insertion cleanup.
            _insertWhitespaceLogic.InsertBlankSpaceBeforeSelfClosingAngleBracket(textDocument);
        }

        /// <summary>
        /// Attempts to run code cleanup on the specified generic document.
        /// </summary>
        /// <param name="document">The document for cleanup.</param>
        /// <param name="isAutoSave">A flag indicating if occurring due to auto-save.</param>
        private void RunCodeCleanupGeneric(Document document, bool isAutoSave)
        {
            var textDocument = (TextDocument)document.Object("TextDocument");

            RunExternalFormatting(textDocument);

            // Perform removal cleanup.
            _removeWhitespaceLogic.RemoveEOLWhitespace(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAtTop(textDocument);
            _removeWhitespaceLogic.RemoveBlankLinesAtBottom(textDocument);
            _removeWhitespaceLogic.RemoveEOFTrailingNewLine(textDocument);
            _removeWhitespaceLogic.RemoveMultipleConsecutiveBlankLines(textDocument);
        }

        #endregion Private Language Methods

        #region Private Cleanup Methods

        /// <summary>
        /// Runs external formatting tools (e.g. Visual Studio, ReSharper).
        /// </summary>
        /// <param name="textDocument">The text document to cleanup.</param>
        private void RunExternalFormatting(TextDocument textDocument)
        {
            RunVisualStudioFormatDocument(textDocument);
            RunReSharperSilentCleanup(textDocument);
        }

        /// <summary>
        /// Runs the Visual Studio built-in format document command.
        /// </summary>
        /// <param name="textDocument">The text document to cleanup.</param>
        private void RunVisualStudioFormatDocument(TextDocument textDocument)
        {
            if (!Settings.Default.Cleaning_RunVisualStudioFormatDocumentCommand) return;

            try
            {
                using (new CursorPositionRestorer(textDocument))
                {
                    _package.IDE.ExecuteCommand("Edit.FormatDocument", String.Empty);
                }
            }
            catch
            {
                // OK if fails, not available for some file types.
            }
        }

        /// <summary>
        /// Runs the ReSharper silent cleanup command.
        /// </summary>
        /// <param name="textDocument">The text document to cleanup.</param>
        private void RunReSharperSilentCleanup(TextDocument textDocument)
        {
            if (!Settings.Default.Compatibility_UseReSharperSilentCleanup) return;

            try
            {
                using (new CursorPositionRestorer(textDocument))
                {
                    _package.IDE.ExecuteCommand("ReSharper_SilentCleanupCode", String.Empty);
                }
            }
            catch
            {
                // OK if fails, not available for some file types.
            }
        }

        #endregion Private Cleanup Methods
    }
}