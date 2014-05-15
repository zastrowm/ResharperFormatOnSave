// Copyright (c) 2013 Cognex Corporation. All Rights Reserved

using System;
using System.Collections.Generic;
using System.Linq;
using EnvDTE;
using ReSharperFormatOnSave.Helpers;
using ReSharperFormatOnSave.Properties;

namespace ReSharperFormatOnSave.Logic.Cleaning
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

    private readonly UndoTransactionHelper _undoTransactionHelper;

    private readonly CodeCleanupAvailabilityLogic _codeCleanupAvailabilityLogic;

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
      _undoTransactionHelper = new UndoTransactionHelper(_package, "CodeMaid Cleanup");
      _codeCleanupAvailabilityLogic = CodeCleanupAvailabilityLogic.GetInstance(_package);
    }

    #endregion Constructors

    /// <summary>
    /// Attempts to run code cleanup on the specified document.
    /// </summary>
    /// <param name="document">The document for cleanup.</param>
    internal void Cleanup(Document document)
    {
      if (!_codeCleanupAvailabilityLogic.ShouldCleanup(document, true))
        return;

      // Make sure the document to be cleaned up is active, required for some commands like
      // format document.
      document.Activate();

      if (_package.IDE.ActiveDocument != document)
      {
        //OutputWindowHelper.WriteLine(document.Name + " did not complete activation before cleaning started.");
      }

      _undoTransactionHelper.Run(
        () => false,
        () => PerformFormat(document),
        ex => HandleException(document, ex));
    }

    /// <summary> Perform formatting on the specified document. </summary>
    private void PerformFormat(Document document)
    {
      _package.IDE.StatusBar.Text = String.Format("ReSharperAutoSave is formatting '{0}'...", document.Name);
      // Perform the set of configured cleanups based on the language.
      var textDocument = (TextDocument)document.Object("TextDocument");
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

      _package.IDE.StatusBar.Text = String.Format("ReSharperAutoSave formatted '{0}'.", document.Name);
    }

    /// <summary> Handle any exception that occurs while attempting to cleanup the document. </summary>
    private void HandleException(Document document, Exception ex)
    {
      // OutputWindowHelper.WriteLine(String.Format("ReSharperAutoSave stopped formatting '{0}': {1}", document.Name, ex));
      _package.IDE.StatusBar.Text =
        String.Format("ReSharperAutoSave stopped formatting '{0}': See output window for more details. {1}",
                      document.Name,
                      ex);
    }
  }
}