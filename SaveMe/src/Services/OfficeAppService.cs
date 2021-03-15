using Prism.Events;
using src.Events;
using src.Models;
using src.Repo;
using src.Services.Process_Managers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace src.Services
{
    public class OfficeAppService : IOfficeAppService
    {
        private readonly IOfficeApplicationSaver _saver;
        private IEventAggregator _ea;
        private readonly IOfficeAppRepo _repo;
        private readonly IOfficeProcessManager _manager;

        public event EventHandler AppClosedEvent;
        public event EventHandler AppOpenedEvent;

        public OfficeAppService(IOfficeApplicationSaver saver,
            IEventAggregator ea,
            IOfficeAppRepo repo, IOfficeProcessManager manager)
        {
            _saver = saver;
            _repo = repo;
            _manager = manager;
            _ea.GetEvent<OfficeAppClosedEvent>().Subscribe(HandleAppClosed);
        }

        public List<string> GetOppenAppNames()
        {
            var openAppNames = new List<string>();
            foreach(var app in _repo.OpenOfficeProcesses)
            {
                openAppNames.Add(app.FullName);
            }
            return openAppNames;
        }

        private void HandleAppOpened(IOfficeApplication app)
        {
            _repo.Insert(app);
            AppOpenedEvent?.Invoke(app, EventArgs.Empty);
        }

        public void AddOpenOfficeApplicationsToRepo(Func<IEnumerable<IOfficeApplication>> getApps)
        {
            foreach(var app in getApps())
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