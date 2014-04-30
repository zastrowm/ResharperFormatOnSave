#region CodeMaid is Copyright 2007-2014 Steve Cadwallader.

// CodeMaid is free software: you can redistribute it and/or modify it under the terms of the GNU
// Lesser General Public License version 3 as published by the Free Software Foundation.
//
// CodeMaid is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
// even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
// Lesser General Public License for more details <http://www.gnu.org/licenses/>.

#endregion CodeMaid is Copyright 2007-2014 Steve Cadwallader.

using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using SteveCadwallader.CodeMaid.Logic.Cleaning;

namespace SteveCadwallader.CodeMaid.UI.Dialogs.CleanupProgress
{
    /// <summary>
    /// The view model representing the state and commands available for cleanup progress.
    /// </summary>
    public class CleanupProgressViewModel : Bindable
    {
        #region Fields

        private readonly BackgroundWorker _backgroundWorker;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="CleanupProgressViewModel" /> class.
        /// </summary>
        /// <param name="package">The hosting package.</param>
        /// <param name="items">The items to cleanup.</param>
        public CleanupProgressViewModel(CodeMaidPackage package, IEnumerable<object> items)
        {
            CodeCleanupManager = CodeCleanupManager.GetInstance(package);

            // Initialize UI elements.
            CountTotal = items.Count();

            // Initialize background worker.
            _backgroundWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            _backgroundWorker.DoWork += backgroundWorker_DoWork;
            _backgroundWorker.ProgressChanged += backgroundWorker_ProgressChanged;
            _backgroundWorker.RunWorkerCompleted += backgroundWorker_RunWorkerCompleted;

            _backgroundWorker.RunWorkerAsync(items);
        }

        #endregion Constructors

        #region Properties

        private string _currentFileName;

        /// <summary>
        /// Gets or sets the name of the current file being cleaned.
        /// </summary>
        public string CurrentFileName
        {
            get { return _currentFileName; }
            set
            {
                if (_currentFileName != value)
                {
                    _currentFileName = value;
                    NotifyPropertyChanged("CurrentFileName");
                }
            }
        }

        private int _countProgress;

        /// <summary>
        /// Gets or sets the progress count
        /// </summary>
        public int CountProgress
        {
            get { return _countProgress; }
            set
            {
                if (_countProgress != value)
                {
                    _countProgress = value;
                    NotifyPropertyChanged("CountProgress");
                }
            }
        }

        private int _countTotal;

        /// <summary>
        /// Gets or sets the total count
        /// </summary>
        public int CountTotal
        {
            get { return _countTotal; }
            set
            {
                if (_countTotal != value)
                {
                    _countTotal = value;
                    NotifyPropertyChanged("CountTotal");
                }
            }
        }

        private bool? _dialogResult;

        /// <summary>
        /// Gets or sets the dialog result.
        /// </summary>
        public bool? DialogResult
        {
            get { return _dialogResult; }
            set
            {
                if (_dialogResult != value)
                {
                    _dialogResult = value;
                    NotifyPropertyChanged("DialogResult");
                }
            }
        }

        private bool _isCanceling;

        /// <summary>
        /// Gets or sets a flag indicating if the operation is being canceled.
        /// </summary>
        public bool IsCanceling
        {
            get { return _isCanceling; }
            set
            {
                if (_isCanceling != value)
                {
                    _isCanceling = value;
                    NotifyPropertyChanged("IsCanceling");
                }
            }
        }

        /// <summary>
        /// Gets or sets the code cleanup manager.
        /// </summary>
        private CodeCleanupManager CodeCleanupManager { get; set; }

        #endregion Properties

        #region Cancel Command

        private DelegateCommand _cancelCommand;

        /// <summary>
        /// Gets the cancel command.
        /// </summary>
        public DelegateCommand CancelCommand
        {
            get { return _cancelCommand ?? (_cancelCommand = new DelegateCommand(OnCancelCommandExecuted, OnCancelCommandCanExecute)); }
        }

        /// <summary>
        /// Called when the <see cref="CancelCommand" /> needs to determine if it can execute.
        /// </summary>
        /// <param name="parameter">The command parameter.</param>
        /// <returns>True if the command can execute, otherwise false.</returns>
        private bool OnCancelCommandCanExecute(object parameter)
        {
            return !IsCanceling;
        }

        /// <summary>
        /// Called when the <see cref="CancelCommand" /> is executed.
        /// </summary>
        /// <param name="parameter">The command parameter.</param>
        private void OnCancelCommandExecuted(object parameter)
        {
            IsCanceling = true;
            CancelCommand.RaiseCanExecuteChanged();

            _backgroundWorker.CancelAsync();
        }

        #endregion Cancel Command

        #region Methods

        /// <summary>
        /// Handles the DoWork event of the backgroundWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">
        /// The <see cref="System.ComponentModel.DoWorkEventArgs" /> instance containing the event data.
        /// </param>
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var bw = (BackgroundWorker)sender;
            var items = (IEnumerable<object>)e.Argument;
            int i = 0;

            foreach (dynamic item in items)
            {
                if (bw.CancellationPending)
                {
                    e.Cancel = true;
                    break;
                }

                bw.ReportProgress(++i, item);

                CodeCleanupManager.Cleanup(item);
            }
        }

        /// <summary>
        /// Handles the ProgressChanged event of the backgroundWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">
        /// The <see cref="System.ComponentModel.ProgressChangedEventArgs" /> instance containing
        /// the event data.
        /// </param>
        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int currentCount = e.ProgressPercentage;
            dynamic currentItem = e.UserState;

            CountProgress = currentCount;
            CurrentFileName = currentItem.Name;
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the backgroundWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">
        /// The <see cref="System.ComponentModel.RunWorkerCompletedEventArgs" /> instance containing
        /// the event data.
        /// </param>
        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Close the dialog.
            DialogResult = true;
        }

        #endregion Methods
    }
}