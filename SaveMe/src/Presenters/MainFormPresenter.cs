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
        private List<IOfficeApp> OpenApps { get; set; } = new();

        private readonly IOfficeAppProvider _provider;

        public MainFormPresenter(IOfficeAppProvider provider)
        {
            _provider = provider;
            _provider.AppClosedEvent += HandleAppClosed;
            _provider.AppOpenedEvent += HandleAppOpened;
        }

        private void HandleAppOpened(object sender, AppOpenedEventArgs e)
        {
            var openAppNames = OpenApps.Select(app => app.Name).ToList();

            switch (e.AppType)
            {
                case OfficeAppType.Excel:
                    AddNewApps(_provider.FetchNewExcelApplications, openAppNames);
                    break;

                case OfficeAppType.Word:
                    AddNewApps(_provider.FetchNewWordApplications, openAppNames);
                    break;

                case OfficeAppType.PowerPoint:
                    AddNewApps(_provider.FetchNewPowerPointApplications, openAppNames);
                    break;
            };
        }

        private void HandleAppClosed(object sender, AppClosedEventArgs e)
        {
            OpenApps.RemoveAll(app => app.Name == e.Name);
            RemoveItemFromViewListBox($"{e.Type} - {e.Name}");
        }

        public void AddNewApps(Func<List<string>, IEnumerable<IOfficeApp>> getApps, List<string> openAppNames)
        {
            foreach (var app in getApps(openAppNames))
            {
                OpenApps.Add(app);
                AddNewAppToViewListBox($"{app.Type} - {app.Name}");
            }
        }

        private void AddNewAppToViewListBox(string newApp)
        {
            View.Invoke(new MethodInvoker(delegate ()
            {
                View.ListOfOpenOfficeApplications.Items.Add(newApp);
            }));
        }

        private void RemoveItemFromViewListBox(string closedApp)
        {
            View.Invoke(new MethodInvoker(delegate ()
            {
                View.ListOfOpenOfficeApplications.Items.Remove($"{closedApp}");
            }));
        }

        public void InitialSetup()
        {
            View.Load += View_Load;
        }

        private void View_Load(object sender, EventArgs e)
        {
            var openAppNames = OpenApps.Select(app => app.Name).ToList();
            AddNewApps(_provider.FetchNewExcelApplications, openAppNames);
            AddNewApps(_provider.FetchNewWordApplications, openAppNames);
            AddNewApps(_provider.FetchNewWordApplications, openAppNames);
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