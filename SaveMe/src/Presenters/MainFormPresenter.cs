using Prism.Events;
using src.Events;
using src.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
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
        public int CurrentAutoSaveFrequency { get; private set; }
        public BindingList<string> OpenOfficeAppNames;
        private readonly IOfficeAppService _service;
        private IEventAggregator _ea;

        public MainFormPresenter(IOfficeAppService service, IEventAggregator ea)
        {
            _service = service;
            _service.OpenOfficeProcesses.ForEach(p => OpenOfficeAppNames.Add(p.MainModule.FileName));
            _ea = ea;
            _ea.GetEvent<OfficeAppClosedEvent>().Subscribe(HandleAppClosed);
            _ea.GetEvent<OfficeAppOpenedEvent>().Subscribe(HandleAppOpened);
        }

        private void HandleAppOpened(Process p)
        {
            OpenOfficeAppNames.Add(p.MainModule.FileName);
        }

        private void HandleAppClosed(Process p)
        {
            OpenOfficeAppNames.Remove(p.MainModule.FileName);
        }

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