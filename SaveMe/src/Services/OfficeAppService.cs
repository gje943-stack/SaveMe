using src.Events;
using src.Models;
using src.Services.Process_Managers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;

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
            AddOpenPowerPointApps();
            AddOpenExcelApps();
            AddOpenWordApps();
        }

        private void _watcher_AppStartedEvent(object sender, OfficeAppOpenedEventArgs e)
        {
            switch (e.AppType)
            {
                case OfficeAppType.Excel:
                    AddOpenExcelApps();
                    break;
                case OfficeAppType.PowerPoint:
                    AddOpenPowerPointApps();
                    break;
                case OfficeAppType.Word:
                    AddOpenWordApps();
                    break;
            }
            NewAppStartedEvent?.Invoke(this, EventArgs.Empty);
        }

        private void AddOpenWordApps()
        {
            while (true)
            {
                foreach (var doc in _provider.FetchOpenWordApplications())
                {
                    if (!AppAlreadyCreated(doc.FullName))
                    {
                        var newApp = new WordApplication(doc);
                        newApp.AppClosed += HandleAppClosed;
                        OpenOfficeApps.Add(newApp);
                    }
                }
            }
        }

        private void AddOpenExcelApps()
        {
            while (true)
            {
                foreach (var wb in _provider.FetchOpenExcelApplications())
                {
                    if (!AppAlreadyCreated(wb.FullName))
                    {
                        var newApp = new ExcelApplication(wb);
                        newApp.AppClosed += HandleAppClosed;
                        OpenOfficeApps.Add(newApp);
                    }
                }
            }
        }

        private void AddOpenPowerPointApps()
        {
            while (true)
            {
                foreach (var pres in _provider.FetchOpenPowerPointApplications())
                {
                    if (!AppAlreadyCreated(pres.FullName))
                    {
                        var newApp = new PowerPointApplication(pres);
                        newApp.AppClosed += HandleAppClosed;
                        OpenOfficeApps.Add(newApp);
                    }
                }
            }
        }

        public List<string> GetOpenAppNames()
        {
            return OpenOfficeApps.Select(app => $"{app.AppType} - {app.FullName}").ToList();
        }


        public async Task SaveApps()
        {
            await Task.Run(() =>
            {
                foreach (var app in OpenOfficeApps)
                {
                    app.Save();
                }
            });
        }

        private void HandleAppClosed(object sender, EventArgs e)
        {
            OpenOfficeApps.Remove(sender as IOfficeApplication);
            AppClosedEvent?.Invoke(sender as IOfficeApplication, EventArgs.Empty);
        }

        private bool AppAlreadyCreated(string appName)
        {
            return OpenOfficeApps.Select(app => app.FullName).ToList().Contains(appName);
        }
    }
}