using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using src.Models;

namespace src.Services.Process_Managers
{
    public class OfficeApplicationProvider : IOfficeApplicationProvider
    {
        public IEnumerable<IOfficeApplication> FetchOpenWordApplications()
        {
            var wordApp = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            if (wordApp != null)
            {
                foreach (Word.Document doc in wordApp.Documents)
                {
                    yield return new WordApplication(doc);
                }
            }
        }

        public IEnumerable<IOfficeApplication> FetchOpenExcelApplications()
        {
            var xlApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            if (xlApp != null)
            {
                foreach (Excel.Workbook wb in xlApp.Workbooks)
                {
                    yield return new ExcelApplication(wb);
                }
            }
        }

        public IEnumerable<IOfficeApplication> FetchOpenPowerPointApplications()
        {
            var ppApp = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            if (ppApp != null)
            {
                foreach (PowerPoint.Presentation p in ppApp.Presentations)
                {
                    yield return new PowerPointApplication(p);
                }
            }
        }
    }
}