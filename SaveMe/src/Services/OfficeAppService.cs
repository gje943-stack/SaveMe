using Prism.Events;
using src.Events;
using src.Models;
using src.Repo;
using src.Services.Process_Managers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace src.Services
{
    public class OfficeAppService : IOfficeAppService
    {
        private readonly IOfficeApplicationProvider _provider;
        private readonly IProcessWatcher _watcher;
        public event EventHandler NewAppStartedEvent;
        public event EventHandler AppClosedEvent;
        public List<IOfficeApplication> OpenOfficeApps { get; set; } = new();

        public OfficeAppService(IOfficeApplicationProvider provider, IProcessWatcher watcher)
        {
            _provider = provider;
            _watcher = watcher;
            _watcher.AppStartedEvent += _watcher_AppStartedEvent;
            SetOpenOfficeApplications(_provider.FetchOpenWordApplications);
            SetOpenOfficeApplications(_provider.FetchOpenPowerPointApplications);
            SetOpenOfficeApplications(_provider.FetchOpenExcelApplications);
        }

        private void _watcher_AppStartedEvent(object sender, OfficeAppOpenedEventArgs e)
        {
            switch (e.AppType)
            {
                case OfficeAppType.Excel:
                    PublishNewAppStartedEvent(_provider.FetchOpenExcelApplications);
                    break;
                case OfficeAppType.PowerPoint:
                    PublishNewAppStartedEvent(_provider.FetchOpenPowerPointApplications);
                    break;
                case OfficeAppType.Word:
                    PublishNewAppStartedEvent(_provider.FetchOpenWordApplications);
                    break;
            }
        }

        private void PublishNewAppStartedEvent(Func<IEnumerable<IOfficeApplication>> getApps)
        {
            var newApp = GetNewOfficeApplication(getApps);
            NewAppStartedEvent?.Invoke(newApp, EventArgs.Empty);
        }

        private IOfficeApplication GetNewOfficeApplication(Func<IEnumerable<IOfficeApplication>> getApps)
        {
            while (true)
            {
                foreach (var app in getApps())
                {
                    if (!OpenOfficeApps.Contains(app))
                    {
                        OpenOfficeApps.Add(app);
                        return app;
                    }
                }
            }
        }

        public List<string> GetOpenAppNames()
        {
            return OpenOfficeApps.Select(app => app.FullName).ToList();
        }


        public async Task SaveApps(List<string> appsToSave)
        {
            await Task.Run(() =>
            {
                foreach (var app in OpenOfficeApps)
                {
                    if (appsToSave.Contains(app.FullName))
                    {
                        app.Save();
                    }
                }
            });
        }

        public void SetOpenOfficeApplications(Func<IEnumerable<IOfficeApplication>> getApps)
        {
            foreach (var app in getApps())
            {
                if (!OpenOfficeApps.Contains(app))
                {
                    app.AppClosed += HandleAppClosed;
                    OpenOfficeApps.Add(app);
                }
            }
        }

        private void HandleAppClosed(object sender, EventArgs e)
        {
            OpenOfficeApps.Remove(sender as IOfficeApplication);
            AppClosedEvent?.Invoke(sender as IOfficeApplication, EventArgs.Empty);
        }
    }
}