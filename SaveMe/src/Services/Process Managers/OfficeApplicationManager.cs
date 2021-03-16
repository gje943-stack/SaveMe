using Prism.Events;
using src.Services.Process_Managers;
using System;
using System.Collections.Generic;
using System.Management;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using src.Models;
using src.Services.Providers;
using src.Events;

namespace src.Services.Process_Managers
{
    public class OfficeApplicationManager : IOfficeApplicationManager
    {
        private readonly IProcessWatcher _watcher;
        private IEventAggregator _ea;
        private readonly IOfficeAppProvider _provider;

        public OfficeApplicationManager(IEventAggregator ea, IProcessWatcher watcher, IOfficeAppProvider provider)
        {
            _ea = ea;
            _watcher = watcher;
            _watcher.ProcessStartedEvent += HandleProcessStartedEvent;
            _provider = provider;
        }

        private void HandleProcessStartedEvent(object sender, EventArgs e)
        {
            var processName = sender as string;
            var newApp = GetNewApp(processName);
            if (newApp != null)
            {
                _ea.GetEvent<OfficeAppOpenedEvent>().Publish(newApp);
            }
        }

        public IOfficeApplication GetNewApp(string processName)
        {
            IOfficeApplication? newApp;
            if (processName.Contains("Excel"))
            {
                var workbook = _provider.GenerateNewWorkbookFromName(processName);
                newApp = new ExcelApplication(workbook, _ea);
            }
            else if (processName.Contains("Word"))
            {
                var doc = _provider.GenerateNewDocumentFromName(processName);
                newApp = new WordApplication(doc, _ea);
            }
            else
            {
                var presentation = _provider.GenerateNewPresentationFromName(processName);
                newApp = new PowerPointApplication(presentation, _ea);
            }
            return newApp;
        }

        public IEnumerable<IOfficeApplication> FetchOpenWordProcesses()
        {
            var app = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            foreach (Word.Document doc in app.Documents)
            {
                yield return new WordApplication(doc, _ea);
            }
        }

        public IEnumerable<IOfficeApplication> FetchOpenExcelProcesses()
        {
            var app = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            foreach (Excel.Workbook wb in app.Workbooks)
            {
                yield return new ExcelApplication(wb, _ea);
            }
        }

        public IEnumerable<IOfficeApplication> FetchOpenPowerPointApplications()
        {
            var app = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            foreach (PowerPoint.Presentation p in app.Presentations)
            {
                yield return new PowerPointApplication(p, _ea);
            }
        }
    }
}