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
using EnvDTE80;
using SteveCadwallader.CodeMaid.Helpers;

namespace SteveCadwallader.CodeMaid.Model.CodeItems
{
    /// <summary>
    /// The representation of a code event.
    /// </summary>
    public class CodeItemEvent : BaseCodeItemElement
    {
        #region Fields

        private readonly Lazy<bool> _isExplicitInterfaceImplementation;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CodeItemEvent" /> class.
        /// </summary>
        public CodeItemEvent()
        {
            // Make exceptions for explicit interface implementations - which report private access
            // but really do not have a meaningful access level.
            _Access = LazyTryDefault(
                () => CodeEvent != null && !IsExplicitInterfaceImplementation ? CodeEvent.Access : vsCMAccess.vsCMAccessPublic);

            _Attributes = LazyTryDefault(
                () => CodeEvent != null ? CodeEvent.Attributes : null);

            _DocComment = LazyTryDefault(
                () => CodeEvent != null ? CodeEvent.DocComment : null);

            _isExplicitInterfaceImplementation = LazyTryDefault(
                () => CodeEvent != null && ExplicitInterfaceImplementationHelper.IsExplicitInterfaceImplementation(CodeEvent));

            _IsStatic = LazyTryDefault(
                () => CodeEvent != null && CodeEvent.IsShared);

            _TypeString = LazyTryDefault(
                () => CodeEvent != null && CodeEvent.Type != null ? CodeEvent.Type.AsString : null);
        }

        #endregion Constructors

        #region BaseCodeItem Overrides

        /// <summary>
        /// Gets the kind.
        /// </summary>
        public override KindCodeItem Kind
        {
            get { return KindCodeItem.Event; }
        }

        /// <summary>
        /// Loads all lazy initialized values immediately.
        /// </summary>
        public override void LoadLazyInitializedValues()
        {
            base.LoadLazyInitializedValues();

            var ieii = IsExplicitInterfaceImplementation;
        }

        #endregion BaseCodeItem Overrides

        #region Properties

        /// <summary>
        /// Gets or sets the VSX CodeEvent.
        /// </summary>
        public CodeEvent CodeEvent { get; set; }

        /// <summary>
        /// Gets a flag indicating if this property is an explicit interface implementation.
        /// </summary>
        public bool IsExplicitInterfaceImplementation { get { return _isExplicitInterfaceImplementation.Value; } }

        #endregion Properties
    }
}