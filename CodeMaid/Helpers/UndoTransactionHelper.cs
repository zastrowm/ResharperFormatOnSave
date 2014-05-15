// Copyright (c) 2013 Cognex Corporation. All Rights Reserved

using System;
using ReSharperFormatOnSave.Properties;

namespace ReSharperFormatOnSave.Helpers
{
  /// <summary>
  /// A helper class for performing actions within the context of an undo transaction.
  /// </summary>
  public class UndoTransactionHelper
  {
    #region Fields

    private readonly CodeMaidPackage _package;
    private readonly string _transactionName;

    #endregion Fields

    #region Constructors

    /// <summary>
    /// Initializes a new instance of the <see cref="UndoTransactionHelper" /> class.
    /// </summary>
    /// <param name="package">The hosting package.</param>
    /// <param name="transactionName">The name of the transaction.</param>
    public UndoTransactionHelper(CodeMaidPackage package, string transactionName)
    {
      _package = package;
      _transactionName = transactionName;
    }

    #endregion Constructors

    #region Methods

    /// <summary>
    /// Runs the specified try action within a try block.
    /// </summary>
    /// <param name="tryAction">The action to be performed within a try block.</param>
    public void Run(Action tryAction)
    {
      Run(() => true, tryAction, ex => { });
    }

    /// <summary>
    /// Runs the specified try action within a try block, and conditionally the catch action
    /// within a catch block.
    /// </summary>
    /// <param name="tryAction">The action to be performed within a try block.</param>
    /// <param name="catchAction">The action to be performed wihin a catch block.</param>
    public void Run(Action tryAction, Action<Exception> catchAction)
    {
      Run(() => true, tryAction, catchAction);
    }

    /// <summary>
    /// Runs the specified try action within a try block, and conditionally the catch action
    /// within a catch block all conditionally within the context of an undo transaction.
    /// </summary>
    /// <param name="undoConditions">
    /// A set of additional conditions for wrapping in an undo context.
    /// </param>
    /// <param name="tryAction">The action to be performed within a try block.</param>
    /// <param name="catchAction">The action to be performed wihin a catch block.</param>
    public void Run(Func<bool> undoConditions, Action tryAction, Action<Exception> catchAction)
    {
      // Start an undo transaction (unless inside one already or other undo conditions are not met).
      // bool shouldCloseUndoContext = false;
      //if (!_package.IDE.UndoContext.IsOpen && undoConditions() && Settings.Default.Compatibility_UseUndoTransactions)
      //{
      //  _package.IDE.UndoContext.Open(_transactionName);
      //  shouldCloseUndoContext = true;
      //}

      try
      {
        tryAction();
      }
      catch (Exception ex)
      {
        catchAction(ex);

        //if (shouldCloseUndoContext)
        //{
        //  _package.IDE.UndoContext.SetAborted();
        //  shouldCloseUndoContext = false;
        //}
      }
      finally
      {
        //// Always close the undo transaction to prevent ongoing interference with the IDE.
        //if (shouldCloseUndoContext)
        //{
        //  _package.IDE.UndoContext.Close();
        //}
      }
    }

    #endregion Methods
  }
}