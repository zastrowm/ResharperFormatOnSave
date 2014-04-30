﻿#region CodeMaid is Copyright 2007-2014 Steve Cadwallader.

// CodeMaid is free software: you can redistribute it and/or modify it under the terms of the GNU
// Lesser General Public License version 3 as published by the Free Software Foundation.
//
// CodeMaid is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
// even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
// Lesser General Public License for more details <http://www.gnu.org/licenses/>.

#endregion CodeMaid is Copyright 2007-2014 Steve Cadwallader.

using EnvDTE;
using SteveCadwallader.CodeMaid.Helpers;
using SteveCadwallader.CodeMaid.Properties;

namespace SteveCadwallader.CodeMaid.Logic.Cleaning
{
    /// <summary>
    /// A class for encapsulating the logic of inserting whitespace.
    /// </summary>
    internal class InsertWhitespaceLogic
    {
        #region Fields

        private readonly CodeMaidPackage _package;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// The singleton instance of the <see cref="InsertWhitespaceLogic" /> class.
        /// </summary>
        private static InsertWhitespaceLogic _instance;

        /// <summary>
        /// Gets an instance of the <see cref="InsertWhitespaceLogic" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        /// <returns>An instance of the <see cref="InsertWhitespaceLogic" /> class.</returns>
        internal static InsertWhitespaceLogic GetInstance(CodeMaidPackage package)
        {
            return _instance ?? (_instance = new InsertWhitespaceLogic(package));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="InsertWhitespaceLogic" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        private InsertWhitespaceLogic(CodeMaidPackage package)
        {
            _package = package;
        }

        #endregion Constructors

        #region Methods

        /// <summary>
        /// Inserts a single blank space before a self-closing angle bracket.
        /// </summary>
        /// <param name="textDocument">The text document to cleanup.</param>
        internal void InsertBlankSpaceBeforeSelfClosingAngleBracket(TextDocument textDocument)
        {
            if (!Settings.Default.Cleaning_InsertBlankSpaceBeforeSelfClosingAngleBrackets) return;

            string pattern = _package.UsePOSIXRegEx
                                 ? @"{[^:b]}/\>"
                                 : @"([^ \t])/>";
            string replacement = _package.UsePOSIXRegEx
                                     ? @"\1 />"
                                     : @"$1 />";

            TextDocumentHelper.SubstituteAllStringMatches(textDocument, pattern, replacement);
        }

        #endregion Methods
    }
}