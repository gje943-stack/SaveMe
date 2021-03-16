using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using src.Models;


namespace src.Services.Providers
{
    public class OfficeAppProvider : IOfficeAppProvider
    {
        public Excel.Workbook? GenerateNewWorkbookFromName(string processName)
        {
            var app = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            foreach(Excel.Workbook wb in app.Workbooks)
            {
                if(wb.Name == processName)
                {
                    return wb;
                }
            }
            return null;
        }

        public Word.Document? GenerateNewDocumentFromName(string processName)
        {
            var app = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            foreach (Word.Document doc in app.Documents)
            {
                if (doc.Name == processName)
                {
                    return doc;
                }
            }
            return null;
        }

        public PowerPoint.Presentation? GenerateNewPresentationFromName(string processName)
        {
            var app = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            foreach(PowerPoint.Presentation pres in app.Presentations)
            {
                if(pres.Name == processName)
                {
                    return pres;
                }
            }
            return null;
        }
    }
}
