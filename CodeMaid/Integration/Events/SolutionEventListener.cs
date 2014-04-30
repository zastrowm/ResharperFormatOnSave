#region CodeMaid is Copyright 2007-2014 Steve Cadwallader.

// CodeMaid is free software: you can redistribute it and/or modify it under the terms of the GNU
// Lesser General Public License version 3 as published by the Free Software Foundation.
//
// CodeMaid is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
// even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
// Lesser General Public License for more details <http://www.gnu.org/licenses/>.

#endregion CodeMaid is Copyright 2007-2014 Steve Cadwallader.

using System;
using EnvDTE;

namespace SteveCadwallader.CodeMaid.Integration.Events
{
    /// <summary>
    /// A class that encapsulates listening for solution events.
    /// </summary>
    internal class SolutionEventListener : BaseEventListener
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SolutionEventListener" /> class.
        /// </summary>
        /// <param name="package">The package hosting the event listener.</param>
        internal SolutionEventListener(CodeMaidPackage package)
            : base(package)
        {
            // Store access to the solutions events, otherwise events will not register properly via DTE.
            SolutionEvents = Package.IDE.Events.SolutionEvents;
            SolutionEvents.Opened += SolutionEvents_Opened;
            SolutionEvents.AfterClosing += SolutionEvents_AfterClosing;
        }

        #endregion Constructors

        #region Internal Events

        /// <summary>
        /// An event raised when a solution has opened.
        /// </summary>
        internal event Action OnSolutionOpened;

        /// <summary>
        /// An event raised when a solution has closed.
        /// </summary>
        internal event Action OnSolutionClosed;

        #endregion Internal Events

        #region Private Properties

        /// <summary>
        /// Gets or sets a pointer to the IDE solution events.
        /// </summary>
        private SolutionEvents SolutionEvents { get; set; }

        #endregion Private Properties

        #region Private Methods

        /// <summary>
        /// An event handler for a solution being opened.
        /// </summary>
        private void SolutionEvents_Opened()
        {
            if (OnSolutionOpened != null)
            {
                OnSolutionOpened();
            }
        }

        /// <summary>
        /// An event handler for a solution being closed.
        /// </summary>
        private void SolutionEvents_AfterClosing()
        {
            if (OnSolutionClosed != null)
            {
                OnSolutionClosed();
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
            if (!IsDisposed)
            {
                IsDisposed = true;

                if (disposing && SolutionEvents != null)
                {
                    SolutionEvents.Opened -= SolutionEvents_Opened;
                    SolutionEvents.AfterClosing -= SolutionEvents_AfterClosing;
                }
            }
        }

        #endregion IDisposable Members
    }
}