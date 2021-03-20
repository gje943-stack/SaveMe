using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;

namespace src.Services.Process_Managers
{
    public class OfficeApplicationProvider : IOfficeApplicationProvider
    {
        public IEnumerable<Word.Document> FetchOpenWordApplications()
        {
            var wordApp = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            if (wordApp != null)
            {
                foreach (Word.Document doc in wordApp.Documents)
                {
                    yield return doc;
                }
            }
        }

        public IEnumerable<Excel.Workbook> FetchOpenExcelApplications()
        {
            var xlApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            if (xlApp != null)
            {
                foreach (Excel.Workbook wb in xlApp.Workbooks)
                {
                    yield return wb;
                }
            }
        }

        public IEnumerable<PowerPoint.Presentation> FetchOpenPowerPointApplications()
        {
            var ppApp = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            if (ppApp != null)
            {
                foreach (PowerPoint.Presentation p in ppApp.Presentations)
                {
                    yield return p;
                }
            }
        }
    }
}