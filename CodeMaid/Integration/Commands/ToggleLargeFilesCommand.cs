using System.ComponentModel.Design;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System;
using ReSharperFormatOnSave.Properties;

namespace ReSharperFormatOnSave.Integration.Commands
{
  /// <summary>
  /// Enables or disables cleaning of large files.
  /// </summary>
  internal class ToggleLargeFilesCommand : BaseCommand
  {
    /// <summary>
    /// Initializes a new instance of the <see cref="ToggleFormatOnSaveCommand" /> class.
    /// </summary>
    /// <param name="package">The hosting package.</param>
    internal ToggleLargeFilesCommand(CodeMaidPackage package)
      : base(package, new CommandID(GuidList.Commands, PkgCmdIDList.ToggleLargeFiles))
    {
    }

    /// <summary>
    /// Called to update the current status of the command.
    /// </summary>
    protected override void OnBeforeQueryStatus()
    {
      Checked = Settings.Default.FormatLargeFiles;
    }

    /// <summary>
    /// Called to execute the command.
    /// </summary>
    protected override void OnExecute()
    {
      bool value = !Settings.Default.FormatLargeFiles;
      Settings.Default.FormatLargeFiles = value;
      Settings.Default.Save();
    }
  }
}