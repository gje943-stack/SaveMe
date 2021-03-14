using Prism.Events;
using src.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace src.Services
{
    public class OfficeAppService : IOfficeAppService
    {
        private readonly IOfficeApplicationSaver _saver;
        private readonly IOfficeProcessDetector _detector;
        private IEventAggregator _ea;
        public List<Process> OpenOfficeProcesses { get; private set; }

        public OfficeAppService(IOfficeApplicationSaver saver, IOfficeProcessDetector detector, IEventAggregator ea)
        {
            _saver = saver;
            _detector = detector;
            _ea = ea;
            _ea.GetEvent<OfficeAppClosedEvent>().Subscribe(HandleAppClosed);
            _ea.GetEvent<ExcelAppOpenedEvent>().Subscribe(HandleAppOpened);
            OpenOfficeProcesses = _detector.FetchOpenOfficeProcesses();
        }

        private void HandleAppOpened(Process p)
        {
            OpenOfficeProcesses.Add(p);
        }

        private void HandleAppClosed(Process p)
        {
            OpenOfficeProcesses.Remove(p);
        }
    }
}