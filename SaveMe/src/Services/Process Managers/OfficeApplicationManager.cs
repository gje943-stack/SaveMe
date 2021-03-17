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
using src.Events;

namespace src.Services.Process_Managers
{
    public class OfficeApplicationManager : IOfficeApplicationManager
    {
        private IEventAggregator _ea;
        private Excel.Application xlApp;
        private Word.Application wordApp;
        private PowerPoint.Application ppApp;

        public OfficeApplicationManager(IEventAggregator ea)
        {
            _ea = ea;
        }

        public IEnumerable<IOfficeApplication> FetchOpenWordProcesses()
        {
            wordApp = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            if (wordApp != null)
            {
                wordApp.DocumentOpen += WordApp_DocumentOpen;
                foreach (Word.Document doc in wordApp.Documents)
                {
                    yield return new WordApplication(doc, _ea);
                }
            }
        }

        public IEnumerable<IOfficeApplication> FetchOpenExcelProcesses()
        {
            xlApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            if (xlApp != null)
            {
                xlApp.WorkbookOpen += XlApp_WorkbookOpen;
                foreach (Excel.Workbook wb in xlApp.Workbooks)
                {
                    yield return new ExcelApplication(wb, _ea);
                }
            }
        }

        public IEnumerable<IOfficeApplication> FetchOpenPowerPointApplications()
        {
            ppApp = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            if (ppApp != null)
            {
                ppApp.PresentationOpen += PpApp_PresentationOpen;
                foreach (PowerPoint.Presentation p in ppApp.Presentations)
                {
                    yield return new PowerPointApplication(p, _ea);
                }
            }
        }

        private void WordApp_DocumentOpen(Word.Document Doc)
        {
            var app = new WordApplication(Doc, _ea);
            _ea.GetEvent<OfficeAppOpenedEvent>().Publish(app);
        }

        private void PpApp_PresentationOpen(PowerPoint.Presentation Pres)
        {
            var app = new PowerPointApplication(Pres, _ea);
            _ea.GetEvent<OfficeAppOpenedEvent>().Publish(app);
        }

        private void XlApp_WorkbookOpen(Excel.Workbook Wb)
        {
            var app = new ExcelApplication(Wb, _ea);
            _ea.GetEvent<OfficeAppOpenedEvent>().Publish(app);
        }
    }
}