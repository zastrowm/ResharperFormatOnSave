// Copyright (c) 2013 Cognex Corporation. All Rights Reserved

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ReSharperFormatOnSave.Helpers;
using ReSharperFormatOnSave.Integration;
using ReSharperFormatOnSave.Integration.Commands;
using ReSharperFormatOnSave.Integration.Events;
using ReSharperFormatOnSave.Logic.Cleaning;

namespace ReSharperFormatOnSave
{
  /// <summary>
  /// This is the class that implements the package exposed by this assembly.
  ///
  /// The minimum requirement for a class to be considered a valid package for Visual Studio is to
  /// implement the IVsPackage interface and register itself with the shell. This package uses the
  /// helper classes defined inside the Managed Package Framework (MPF) to do it: it derives from
  /// the Package class that provides the implementation of the IVsPackage interface and uses the
  /// registration attributes defined in the framework to register itself and its components with
  /// the shell.
  /// </summary>
  [PackageRegistration(UseManagedResourcesOnly = true)]
  // Tells Visual Studio utilities that this is a package that needs registered.
  [InstalledProductRegistration("#110", "#112", "#114", IconResourceID = 400,
    LanguageIndependentName = "ReSharperFormatOnSave")] // VS Help/About details (Name, Description, Version, Icon).
  [ProvideAutoLoad("ADFC4E64-0397-11D1-9F4E-00A0C911004F")]
  // Force CodeMaid to load on startup so menu items can determine their state.
  [ProvideBindingPath]
  [ProvideMenuResource(1000, 1)] // This attribute is needed to let the shell know that this package exposes some menus.
  [Guid(GuidList.GuidCodeMaidPackageString)] // Package unique GUID.
  public sealed class CodeMaidPackage : Package, IVsInstalledProduct
  {
    #region Fields

    /// <summary>
    /// An internal collection of the commands registered by this package.
    /// </summary>
    private readonly ICollection<BaseCommand> _commands = new List<BaseCommand>();

    /// <summary>
    /// The IComponentModel service.
    /// </summary>
    private IComponentModel _iComponentModel;

    /// <summary>
    /// The top level application instance of the VS IDE that is executing this package.
    /// </summary>
    private DTE2 _ide;

    #endregion Fields

    /// <summary>
    /// Gets or sets the code cleanup availability logic.
    /// </summary>
    private CodeCleanupAvailabilityLogic CodeCleanupAvailabilityLogic { get; set; }

    /// <summary>
    /// Gets or sets the code cleanup manager.
    /// </summary>
    private CodeCleanupManager CodeCleanupManager { get; set; }

    #region Constructors

    /// <summary>
    /// Default constructor of the package. Inside this method you can place any initialization
    /// code that does not require any Visual Studio service because at this point the package
    /// object is created but not sited yet inside Visual Studio environment. The place to do
    /// all the other initialization is the Initialize method.
    /// </summary>
    public CodeMaidPackage()
    {
      Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this));

      CodeCleanupAvailabilityLogic = CodeCleanupAvailabilityLogic.GetInstance(this);
      CodeCleanupManager = CodeCleanupManager.GetInstance(this);
    }

    #endregion Constructors

    #region Public Integration Properties

    /// <summary>
    /// Gets the top level application instance of the VS IDE that is executing this package.
    /// </summary>
    public DTE2 IDE
    {
      get { return _ide ?? (_ide = (DTE2) GetService(typeof(DTE))); }
    }

    /// <summary>
    /// Gets the menu command service.
    /// </summary>
    private OleMenuCommandService MenuCommandService
    {
      get { return GetService(typeof(IMenuCommandService)) as OleMenuCommandService; }
    }

    #endregion Public Integration Properties

    #region Private Event Listener Properties

    /// <summary>
    /// Gets or sets the running document table event listener.
    /// </summary>
    private RunningDocumentTableEventListener RunningDocumentTableEventListener { get; set; }

    /// <summary>
    /// Gets or sets the shell event listener.
    /// </summary>
    private ShellEventListener ShellEventListener { get; set; }

    #endregion Private Event Listener Properties

    #region Private Service Properties

    /// <summary>
    /// Gets the shell service.
    /// </summary>
    private IVsShell ShellService
    {
      get { return GetService(typeof(SVsShell)) as IVsShell; }
    }

    #endregion Private Service Properties

    #region Package Members

    /// <summary>
    /// Initialization of the package; this method is called right after the package is sited,
    /// so this is the place where you can put all the initialization code that rely on services
    /// provided by VisualStudio.
    /// </summary>
    protected override void Initialize()
    {
      Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this));
      base.Initialize();

      RegisterCommands();
      RegisterShellEventListener();
    }

    #endregion Package Members

    #region IVsInstalledProduct Members

    public int IdBmpSplash(out uint pIdBmp)
    {
      pIdBmp = 400;
      return VSConstants.S_OK;
    }

    public int IdIcoLogoForAboutbox(out uint pIdIco)
    {
      pIdIco = 400;
      return VSConstants.S_OK;
    }

    public int OfficialName(out string pbstrName)
    {
      pbstrName = GetResourceString("@110");
      return VSConstants.S_OK;
    }

    public int ProductDetails(out string pbstrProductDetails)
    {
      pbstrProductDetails = GetResourceString("@112");
      return VSConstants.S_OK;
    }

    public int ProductID(out string pbstrPID)
    {
      pbstrPID = GetResourceString("@114");
      return VSConstants.S_OK;
    }

    private string GetResourceString(string resourceName)
    {
      string resourceValue;
      var resourceManager = (IVsResourceManager) GetService(typeof(SVsResourceManager));
      if (resourceManager == null)
      {
        throw new InvalidOperationException(
          "Could not get SVsResourceManager service. Make sure that the package is sited before calling this method");
      }

      Guid packageGuid = GetType().GUID;
      int hr = resourceManager.LoadResourceString(
        ref packageGuid,
        -1,
        resourceName,
        out resourceValue);
      ErrorHandler.ThrowOnFailure(hr);

      return resourceValue;
    }

    #endregion IVsInstalledProduct Members

    #region Private Methods

    /// <summary>
    /// Register the package commands (which must exist in the .vsct file).
    /// </summary>
    private void RegisterCommands()
    {
      var menuCommandService = MenuCommandService;
      if (menuCommandService != null)
      {
        // Create the individual commands, which internally register for command events.
        _commands.Add(new ToggleFormatOnSaveCommand(this));

        // Add all commands to the menu command service.
        foreach (var command in _commands)
        {
          menuCommandService.AddCommand(command);
        }
      }
    }

    /// <summary>
    /// Registers the shell event listener.
    /// </summary>
    /// <remarks>
    /// This event listener is registered by itself and first to wait for the shell to be ready
    /// for other event listeners to be registered.
    /// </remarks>
    private void RegisterShellEventListener()
    {
      ShellEventListener = new ShellEventListener(this, ShellService);
      ShellEventListener.ShellAvailable += RegisterNonShellEventListeners;
    }

    /// <summary>
    /// Register the package event listeners.
    /// </summary>
    /// <remarks>
    /// This must occur after the DTE service is available since many of the events are based
    /// off of the DTE object.
    /// </remarks>
    private void RegisterNonShellEventListeners()
    {
      // Create event listeners and register for events.
      var menuCommandService = MenuCommandService;
      if (menuCommandService != null)
      {
        RunningDocumentTableEventListener = new RunningDocumentTableEventListener(this);
        RunningDocumentTableEventListener.BeforeSave += OnBeforeDocumentSave;
      }
    }

    /// <summary>
    /// Called before a document is saved in order to potentially run code cleanup.
    /// </summary>
    /// <param name="document">The document about to be saved.</param>
    internal void OnBeforeDocumentSave(Document document)
    {
      if (!CodeCleanupAvailabilityLogic.ShouldCleanup(document))
        return;

      using (new ActiveDocumentRestorer(this))
      {
        CodeCleanupManager.Cleanup(document, true);
      }
    }

    #endregion Private Methods

    #region IDisposable Members

    /// <summary>
    /// Releases unmanaged and - optionally - managed resources
    /// </summary>
    /// <param name="disposing">
    /// <c>true</c> to release both managed and unmanaged resources; <c>false</c> to release
    /// only unmanaged resources.
    /// </param>
    protected override void Dispose(bool disposing)
    {
      base.Dispose(disposing);

      if (RunningDocumentTableEventListener != null)
      {
        RunningDocumentTableEventListener.Dispose();
      }

      if (ShellEventListener != null)
      {
        ShellEventListener.Dispose();
      }
    }

    #endregion IDisposable Members
  }
}