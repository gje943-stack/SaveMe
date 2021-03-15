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
        public Excel.Workbook GenerateNewWorkbook(string processName)
        {
            var res = new Excel.Workbook();
            var app = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            foreach(Excel.Workbook wb in app.Workbooks)
            {
                if(wb.Name == processName)
                {
                    res = wb;
                }
            }
            return res;
        }

        public Word.Document GenerateNewDocument(string processName)
        {
            var res = new Word.Document();
            var app = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            foreach (Word.Document doc in app.Documents)
            {
                if (doc.Name == processName)
                {
                    res = doc;
                }
            }
            return res;
        }
    }
}
