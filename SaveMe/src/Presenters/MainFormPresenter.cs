using Prism.Events;
using src.Events;
using src.Models;
using src.Services;
using src.Static;
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
        private readonly IOfficeAppService _service;
        private IEventAggregator _ea;

        public MainFormPresenter(IOfficeAppService service, IEventAggregator ea)
        {
            _service = service;
            _ea = ea;
        }

        public void InitialSetup()
        {
            _ea.GetEvent<OfficeAppClosedEvent>().Subscribe(HandleAppClosed);
            _ea.GetEvent<OfficeAppOpenedEvent>().Subscribe(HandleAppOpened);
            View.Load += View_Load;
        }

        private void View_Load(object sender, EventArgs e)
        {
            
        }

        private void HandleAppOpened(IOfficeApplication appOpened)
        {
            View.OpenOfficeAppNames.Add(appOpened.FullName);
        }

        private void HandleAppClosed(IOfficeApplication appClosed)
        {
            View.OpenOfficeAppNames.Remove(appClosed.FullName);
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
            var newFreq = AutoSaveFrequencies._[View.DropDownAutoSaveFrequency.SelectedItem.ToString()];
            CurrentAutoSaveFrequency = newFreq;
        }
    }
}