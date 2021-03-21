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

        public OfficeAppProvider(IProcessWatcher watcher)
        {
            _watcher = watcher;
            _watcher.NewProcessEvent += NewProcessEventHandler;
        }

        private void NewProcessEventHandler(object sender, NewProcessEventArgs e)
        {
            AppOpenedEvent?.Invoke(this, new AppOpenedEventArgs(e.ProcessType));
        }

        public IEnumerable<IOfficeApp> FetchNewWordApplications(List<string> openApps)
        {
            var wordApp = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            if (wordApp != null)
            {
                RegisterWordAppClosedEvent(wordApp);
                foreach (Word.Document doc in wordApp.Documents)
                {
                    if (!openApps.Contains(doc.FullName))
                    {
                        yield return new OfficeApp(OfficeAppType.Word, doc.FullName);
                    }
                }
            }
        }

        public IEnumerable<IOfficeApp> FetchNewExcelApplications(List<string> openApps)
        {
            var xlApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            if (xlApp != null)
            {
                while (xlApp.Workbooks.Count == openApps.Count)
                {
                    xlApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
                    Thread.Sleep(10);
                };
                RegisterExcelAppClosedEvent(xlApp);
                foreach (Excel.Workbook wb in xlApp.Workbooks)
                {
                    if (!openApps.Contains(wb.FullName))
                    {
                        yield return new OfficeApp(OfficeAppType.Excel, wb.FullName);
                    }
                }
            }
        }

        public IEnumerable<IOfficeApp> FetchNewPowerPointApplications(List<string> openApps)
        {
            var ppApp = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            if (ppApp != null)
            {
                RegisterPowerPointAppClosedEvent(ppApp);
                foreach (PowerPoint.Presentation p in ppApp.Presentations)
                {
                    if (!openApps.Contains(p.FullName))
                        yield return new OfficeApp(OfficeAppType.PowerPoint, p.FullName);
                }
            }
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