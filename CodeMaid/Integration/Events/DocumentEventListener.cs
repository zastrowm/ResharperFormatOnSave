﻿#region CodeMaid is Copyright 2007-2014 Steve Cadwallader.

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
    /// A class that encapsulates listening for document events.
    /// </summary>
    internal class DocumentEventListener : BaseEventListener
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentEventListener" /> class.
        /// </summary>
        /// <param name="package">The package hosting the event listener.</param>
        internal DocumentEventListener(CodeMaidPackage package)
            : base(package)
        {
            // Store access to the document events, otherwise events will not register properly via DTE.
            DocumentEvents = Package.IDE.Events.DocumentEvents;
            DocumentEvents.DocumentClosing += DocumentEvents_DocumentClosing;
        }

        #endregion Constructors

        #region Internal Events

        /// <summary>
        /// An event raised when a document is closing.
        /// </summary>
        internal event Action<Document> OnDocumentClosing;

        #endregion Internal Events

        #region Private Properties

        /// <summary>
        /// Gets or sets a pointer to the IDE document events.
        /// </summary>
        private DocumentEvents DocumentEvents { get; set; }

        #endregion Private Properties

        #region Private Event Handlers

        /// <summary>
        /// An event handler for a document closing.
        /// </summary>
        /// <param name="document">The document that is closing.</param>
        private void DocumentEvents_DocumentClosing(Document document)
        {
            if (OnDocumentClosing != null)
            {
                OnDocumentClosing(document);
            }
        }

        #endregion Private Event Handlers

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

                if (disposing && DocumentEvents != null)
                {
                    DocumentEvents.DocumentClosing -= DocumentEvents_DocumentClosing;
                }
            }
        }

        #endregion IDisposable Members
    }
}