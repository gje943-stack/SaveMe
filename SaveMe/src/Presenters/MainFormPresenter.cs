using src.Services;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace src.Presenters
{
    /// <summary>
    ///     Feeds data to the View to be displayed.
    /// </summary>
    /// <typeparam name="TView">The View this presenter interacts with.</typeparam>
    internal class MainFormPresenter<TView> : IPresenter<TView> where TView : IMainFormView
    {
        public TView View { get; set; }

        public Timer Timer { get; private set; }

        private readonly IOfficeApplicationService _officeApplicationService;

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

        public MainFormPresenter(IOfficeApplicationService officeApplicationService)
        {
            PopulateListOfOpenOfficeApplications();
            SubscribeToViewEvents();
            InitializeTimer();
            _officeApplicationService = officeApplicationService;
        }

        private void PopulateListOfOpenOfficeApplications()
        {
        }

        private void InitializeTimer()
        {
            Timer = new Timer(SaveSelectedApplications, null, 0, CurrentAutoSaveFrequency);
        }

        private void SaveSelectedApplications(object state)
        {
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