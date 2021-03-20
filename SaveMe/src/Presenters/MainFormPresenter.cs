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
using System.Windows.Forms;

namespace src.Presenters
{
    /// <summary>
    ///     Feeds data to the View to be displayed.
    /// </summary>
    /// <typeparam name="TView">The View this presenter interacts with.</typeparam>
    public class MainFormPresenter<TView> : IPresenter<TView> where TView : IMainFormView
    {
        public TView View { get; set; }
        public System.Threading.Timer Timer { get; private set; }
        public int CurrentAutoSaveFrequency { get; private set; }
        private readonly IOfficeAppService _service;

        public MainFormPresenter(IOfficeAppService service)
        {
            _service = service;
            _service.NewAppStartedEvent += _service_NewAppStartedEvent;
            _service.AppClosedEvent += _service_AppClosedEvent;
        }

        private void _service_AppClosedEvent(object sender, EventArgs e)
        {
            var closedApp = sender as IOfficeApplication;
            View.Invoke(new MethodInvoker(delegate ()
            {
                View.ListOfOpenOfficeApplications.Items.Remove(closedApp.FullName);
            }));
        }

        private void _service_NewAppStartedEvent(object sender, EventArgs e)
        {
            var newApp = sender as IOfficeApplication;
            View.Invoke(new MethodInvoker(delegate ()
            {
                foreach(var app in _service.OpenOfficeApps)
                {
                    if (!View.ListOfOpenOfficeApplications.Items.Contains(app.FullName))
                    {
                        View.ListOfOpenOfficeApplications.Items.Add(newApp.FullName);
                    }
                }
            }));
        }

        public void InitialSetup()
        {
            View.Load += View_Load;
        }

        private void View_Load(object sender, EventArgs e)
        {
            PopulateOpenAppNames();
        }

        private void PopulateOpenAppNames()
        {
            foreach (var appName in _service.GetOpenAppNames())
            {
                View.Invoke(new MethodInvoker(delegate()
                {
                    View.ListOfOpenOfficeApplications.Items.Add(appName);
                }));
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
            var newFreq = AutoSaveFrequencies._[View.DropDownAutoSaveFrequency.SelectedItem.ToString()];
            CurrentAutoSaveFrequency = newFreq;
        }
    }
}