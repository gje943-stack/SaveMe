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
using System.Threading.Tasks;

namespace src.Services
{
    public class OfficeAppService : IOfficeAppService
    {
        private IEventAggregator _ea;
        private readonly IOfficeAppRepo _repo;
        private readonly IOfficeApplicationManager _manager;

        public OfficeAppService(IEventAggregator ea,
            IOfficeAppRepo repo,
            IOfficeApplicationManager manager)
        {
            _repo = repo;
            _manager = manager;
            _ea = ea;
            _ea.GetEvent<OfficeAppClosedEvent>().Subscribe(HandleAppClosed);
            _ea.GetEvent<OfficeAppOpenedEvent>().Subscribe(HandleAppOpened);
            AddOpenOfficeApplicationsToRepo(_manager.FetchOpenExcelProcesses);
            AddOpenOfficeApplicationsToRepo(_manager.FetchOpenExcelProcesses);
            AddOpenOfficeApplicationsToRepo(_manager.FetchOpenExcelProcesses);
        }

        public List<string> GetOpenAppNames()
        {
            return _repo.OpenOfficeApps.Select(app => app.FullName).ToList();
        }

        public async Task SaveApps(List<string> appsToSave)
        {
            await Task.Run(() =>
            {
                foreach (var app in _repo.OpenOfficeApps)
                {
                    if (appsToSave.Contains(app.FullName))
                    {
                        app.Save();
                    }
                }
            });
        }

        private void HandleAppOpened(IOfficeApplication app)
        {
            _repo.Insert(app);
        }

        public void AddOpenOfficeApplicationsToRepo(Func<IEnumerable<IOfficeApplication>> getApps)
        {
            foreach (var app in getApps())
            {
                _repo.Insert(app);
            }
        }

        private void HandleAppClosed(IOfficeApplication app)
        {
            _repo.Delete(app);
        }
    }
}