using src.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Threading;

namespace src.Presenters
{
    /// <summary>
    ///     Feeds data to the View to be displayed.
    /// </summary>
    /// <typeparam name="TView">The View this presenter interacts with.</typeparam>
    public class MainFormPresenter<TView> : IPresenter<TView> where TView : IMainFormView
    {
        public TView View { get; set; }
        public Timer Timer { get; private set; }
        private readonly IOfficeApplicationSaver _saver;
        private readonly IOfficeProcessDetector _detector;
        public BindingList<string> OpenOfficeProcesses { get; private set; }

        public int CurrentAutoSaveFrequency { get; private set; }

        /// <summary>
        ///     A dictonary with
        ///     Key: autosave frequencies in string format.
        ///     Value: autosave frequencies in milliseconds.
        /// </summary>
        private readonly Dictionary<string, int> _autoSaveFrequencies = new Dictionary<string, int>
        {
            {"1 Minute", 6000 },
            {"10 Minutes", 10000 },
            {"30 Minutes", 30000 },
            {"1 Hour", 60000 }
        };

        public MainFormPresenter(IOfficeApplicationSaver saver, IOfficeProcessDetector detector)
        {
            _saver = saver;
            _detector = detector;
            GetOpenOfficeProcesses();
        }

        private void GetOpenOfficeProcesses()
        {
            foreach(var process in _detector.FetchOpenOfficeProcesses())
            {
                OpenOfficeProcesses.Add(process.MainModule.FileName);
            }
        }

        public void SubscribeToViewEvents()
        {
            View.AutoSaveFrequencySelected += View_AutoSaveFrequencySelected;
        }

        public void View_AutoSaveFrequencySelected(object sender, EventArgs e)
        {
            ChangeAutoSaveFrequency();
        }

        public void ChangeAutoSaveFrequency()
        {
            var newFreq = _autoSaveFrequencies[View.DropDownAutoSaveFrequency.SelectedItem.ToString()];
            CurrentAutoSaveFrequency = newFreq;
        }
    }
}