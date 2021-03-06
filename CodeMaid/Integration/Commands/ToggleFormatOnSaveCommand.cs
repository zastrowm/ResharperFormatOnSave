// Copyright (c) 2013 Cognex Corporation. All Rights Reserved

using System.ComponentModel.Design;
using ReSharperFormatOnSave.Properties;

namespace ReSharperFormatOnSave.Integration.Commands
{
  /// <summary>
  /// A command that provides for launching the CodeMaid configuration to the general cleanup page.
  /// </summary>
  internal class ToggleFormatOnSaveCommand : BaseCommand
  {
    /// <summary>
    /// Initializes a new instance of the <see cref="ToggleFormatOnSaveCommand" /> class.
    /// </summary>
    /// <param name="package">The hosting package.</param>
    internal ToggleFormatOnSaveCommand(CodeMaidPackage package)
      : base(package, new CommandID(GuidList.Commands, PkgCmdIDList.ToggleFormatOnSave))
    {
    }

    /// <summary>
    /// Called to update the current status of the command.
    /// </summary>
    protected override void OnBeforeQueryStatus()
    {
      Checked = Settings.Default.EnableFormatOnSave;
    }

    /// <summary>
    /// Called to execute the command.
    /// </summary>
    protected override void OnExecute()
    {
      bool value = !Settings.Default.EnableFormatOnSave;
      Settings.Default.EnableFormatOnSave = value;
      Settings.Default.Save();
    }
  }
}