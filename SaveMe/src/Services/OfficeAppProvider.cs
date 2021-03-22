using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using src.Events;
using src.Models;
using System.Threading;

namespace src.Services
{
    public class OfficeAppProvider : IOfficeAppProvider
    {
        public event EventHandler<AppClosedEventArgs> AppClosedEvent;

        public event EventHandler<AppOpenedEventArgs> AppOpenedEvent;

        public IProcessWatcher _watcher;

        private Word.Application? wordApp = (Word.Application)Marshal2.GetActiveObject("Word.Application");
        private Excel.Application? xlApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
        private PowerPoint.Application? ppApp = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");

        public OfficeAppProvider(IProcessWatcher watcher)
        {
            _watcher = watcher;
            _watcher.NewProcessEvent += NewProcessEventHandler;
        }

        private void NewProcessEventHandler(object sender, NewProcessEventArgs e)
        {
            AppOpenedEvent?.Invoke(this, new AppOpenedEventArgs(e.ProcessType));
        }

        public IEnumerable<IOfficeApp> TryFetchWordApplicationsOnStartup()
        {
            if (wordApp != null && wordApp.Documents.Count > 0)
            {
                RegisterWordAppClosedEvent(wordApp);
                foreach (Word.Document doc in wordApp.Documents)
                {
                    yield return new OfficeApp(OfficeAppType.Word, doc.FullName);
                }
            }
        }

        public IEnumerable<IOfficeApp> TryFetchExcelApplicationsOnStartup()
        {
            if (xlApp != null && xlApp.Workbooks.Count > 0)
            {
                RegisterExcelAppClosedEvent(xlApp);
                foreach (Excel.Workbook wb in xlApp.Workbooks)
                {
                    yield return new OfficeApp(OfficeAppType.Excel, wb.FullName);
                }
            }
        }

        public IEnumerable<IOfficeApp> TryFetchPowerPointApplicationsOnStartup()
        {
            if (ppApp != null && ppApp.Presentations.Count > 0)
            {
                RegisterPowerPointAppClosedEvent(ppApp);
                foreach (PowerPoint.Presentation p in ppApp.Presentations)
                {
                    yield return new OfficeApp(OfficeAppType.PowerPoint, p.FullName);
                }
            }
        }

        public IOfficeApp FetchNewWordApplication(int numAppsAlreadyKnown)
        {
            while (wordApp.Documents.Count == numAppsAlreadyKnown || wordApp == null)
            {
                Thread.Sleep(100);
                wordApp = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            }
            RegisterWordAppClosedEvent(wordApp);
            var doc = wordApp.Documents[1];
            return new OfficeApp(OfficeAppType.Word, doc.FullName);
        }

        public IOfficeApp FetchNewExcelApplication(int numAppsAlreadyKnown)
        {
            while (xlApp.Workbooks.Count == numAppsAlreadyKnown || xlApp == null)
            {
                Thread.Sleep(100);
                xlApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            }
            RegisterExcelAppClosedEvent(xlApp);
            var wb = xlApp.Workbooks[xlApp.Workbooks.Count];
            var x = wb.Name;
            var y = wb.Names;
            return new OfficeApp(OfficeAppType.Excel, wb.FullName);
        }

        public IOfficeApp FetchNewPowerPointApplication(int numAppsAlreadyKnown)
        {
            while (ppApp.Presentations.Count == numAppsAlreadyKnown || ppApp == null)
            {
                Thread.Sleep(100);
                ppApp = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            }
            RegisterPowerPointAppClosedEvent(ppApp);
            var pres = wordApp.Documents[ppApp.Presentations.Count];
            return new OfficeApp(OfficeAppType.PowerPoint, pres.FullName);
        }

        private void RegisterWordAppClosedEvent(Word.Application app)
        {
            app.DocumentBeforeClose += (Word.Document doc, ref bool cancel) => AppClosedEvent?.Invoke(this, new AppClosedEventArgs(doc.FullName, OfficeAppType.Word));
        }

        private void RegisterExcelAppClosedEvent(Excel.Application app)
        {
            app.WorkbookBeforeClose += (Excel.Workbook wb, ref bool cancel) => AppClosedEvent?.Invoke(this, new AppClosedEventArgs(wb.FullName, OfficeAppType.Excel));
        }

        private void RegisterPowerPointAppClosedEvent(PowerPoint.Application app)
        {
            app.PresentationBeforeClose += (PowerPoint.Presentation pres, ref bool cancel) => AppClosedEvent?.Invoke(this, new AppClosedEventArgs(pres.FullName, OfficeAppType.Word));
        }
    }
}