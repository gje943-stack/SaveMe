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
using System.Threading.Tasks;
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
            var numOfWordAppsKnown = OpenApps.Where(app => app.Type == OfficeAppType.Word).Count();
            var numOfExcelAppsKnown = OpenApps.Where(app => app.Type == OfficeAppType.Excel).Count();
            var numOfPowerPointAppsKnown = OpenApps.Where(app => app.Type == OfficeAppType.PowerPoint).Count();
            switch (e.AppType)
            {
                case OfficeAppType.Excel:
                    AddNewlyStartedApp(_provider.FetchNewExcelApplication, numOfExcelAppsKnown);
                    break;

                case OfficeAppType.Word:
                    AddNewlyStartedApp(_provider.FetchNewWordApplication, numOfWordAppsKnown);
                    break;

                case OfficeAppType.PowerPoint:
                    AddNewlyStartedApp(_provider.FetchNewPowerPointApplication, numOfPowerPointAppsKnown);
                    break;
            };
        }

        private void HandleAppClosed(object sender, AppClosedEventArgs e)
        {
            OpenApps.RemoveAll(app => app.Name == e.Name);
            RemoveItemFromViewListBox($"{e.Type} - {e.Name}");
        }

        private void AddNewlyStartedApp(Func<int, IOfficeApp> getNewApp, int numAppsCurrentlyKnown)
        {
            var newApp = getNewApp(numAppsCurrentlyKnown);
            OpenApps.Add(newApp);
            AddNewAppToViewListBox(newApp);
        }

        public void AddAppsOnStartup(Func<IEnumerable<IOfficeApp>> getApps)
        {
            foreach (var app in getApps())

            {
                OpenApps.Add(app);
                AddNewAppToViewListBox(app);
            }
        }

        private void AddNewAppToViewListBox(IOfficeApp app)
        {
            View.Invoke(new MethodInvoker(delegate ()
            {
                View.ListOfOpenOfficeApplications.Items.Add($"{app.Type} - {app.Name}");
            }));
        }

        private void RemoveItemFromViewListBox(string closedAppName)
        {
            View.Invoke(new MethodInvoker(delegate ()
            {
                View.ListOfOpenOfficeApplications.Items.Remove(closedAppName);
            }));
        }

        public void InitialSetup()
        {
            View.Load += View_Load;
        }

        private void View_Load(object sender, EventArgs e)
        {
            AddAppsOnStartup(_provider.TryFetchExcelApplicationsOnStartup);
            AddAppsOnStartup(_provider.TryFetchPowerPointApplicationsOnStartup);
            AddAppsOnStartup(_provider.TryFetchWordApplicationsOnStartup);
            SubscribeToViewEvents();
        }

        public void SubscribeToViewEvents()
        {
            View.AutoSaveFrequencySelected += View_AutoSaveFrequencySelected;
            View.SaveAllButtonClicked += async (s, e) => await SaveAllOpenApps(s, e);
            View.SaveSelectedButtonClicked += async (s, e) => await SaveSingleApp(s, e);
        }

        private async Task SaveSingleApp(object s, EventArgs e)
        {
            var index = View.ListOfOpenOfficeApplications.SelectedIndex;
            var appToSave = OpenApps[index];
            await Task.Run(() => OfficeAppSaver.SaveSingleApplication(appToSave));
        }

        private async Task SaveAllOpenApps(object sender, EventArgs e)
        {
            await Task.Run(() => OfficeAppSaver.SaveAll());
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