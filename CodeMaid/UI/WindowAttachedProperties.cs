﻿#region CodeMaid is Copyright 2007-2014 Steve Cadwallader.

// CodeMaid is free software: you can redistribute it and/or modify it under the terms of the GNU
// Lesser General Public License version 3 as published by the Free Software Foundation.
//
// CodeMaid is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
// even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
// Lesser General Public License for more details <http://www.gnu.org/licenses/>.

#endregion CodeMaid is Copyright 2007-2014 Steve Cadwallader.

using System.Windows;

namespace SteveCadwallader.CodeMaid.UI
{
    /// <summary>
    /// A helper class for attached properties on <see cref="Window" />.
    /// </summary>
    /// <remarks>DialogResult attached property based on: http://blog.excastle.com/2010/07/25/mvvm-and-dialogresult-with-no-code-behind/.</remarks>
    public static class WindowAttachedProperties
    {
        #region DialogResult (Attached Property)

        /// <summary>
        /// The dependency property definition for the DialogResult attached property.
        /// </summary>
        public static DependencyProperty DialogResultProperty = DependencyProperty.RegisterAttached(
            "DialogResult", typeof(bool?), typeof(WindowAttachedProperties),
            new FrameworkPropertyMetadata(OnDialogResultChanged));

        /// <summary>
        /// Gets the DialogResult value from the specified target.
        /// </summary>
        /// <param name="target">The target.</param>
        /// <returns>The value.</returns>
        public static bool? GetDialogResult(Window target)
        {
            return (bool?)target.GetValue(DialogResultProperty);
        }

        /// <summary>
        /// Sets the DialogResult value on the specified target.
        /// </summary>
        /// <param name="target">The target.</param>
        /// <param name="value">The value.</param>
        public static void SetDialogResult(Window target, bool? value)
        {
            target.SetValue(DialogResultProperty, value);
        }

        /// <summary>
        /// Called when the DialogResult attached property has changed.
        /// </summary>
        /// <param name="obj">The dependency object where the value has changed.</param>
        /// <param name="e">
        /// The <see cref="System.Windows.DependencyPropertyChangedEventArgs" /> instance containing
        /// the event data.
        /// </param>
        private static void OnDialogResultChanged(DependencyObject obj, DependencyPropertyChangedEventArgs e)
        {
            var window = obj as Window;
            if (window != null)
            {
                window.DialogResult = e.NewValue as bool?;
            }
        }

        #endregion DialogResult (Attached Property)
    }
}