using System;
using System.Collections.Generic;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace src.Services
{
    public class OfficeApplicationSaver : IOfficeApplicationSaver
    {
        public void SaveApplication<TApplication>(TApplication app, string fileName)
        {
            if (app is Excel.Application)
            {
                SaveExcelApplication(app as Excel.Application, fileName);
            }
            if (app is PowerPoint.Application)
            {
                SavePowerPointApplication(app as PowerPoint.Application, fileName);
            }
            if (app is Word.Application)
            {
                SaveWordApplication(app as Word.Application, fileName);
            }
        }

        private void SaveExcelApplication(Excel.Application app, string fileName)
        {
            foreach (Excel.Workbook wb in app.Workbooks)
            {
                if (wb.Name == fileName)
                {
                    wb.Save();
                    //TODO: CONFIGURE DIRECTORY TO SAVE FILES
                }
            }
        }

        private void SaveWordApplication(Word.Application app, string fileName)
        {
            foreach (Word.Document doc in app.Documents)
            {
                if (doc.Name == fileName)
                {
                    doc.Save();
                    //TODO: CONFIGURE DIRECTORY TO SAVE FILES
                }
            }
        }

        private void SavePowerPointApplication(PowerPoint.Application app, string fileName)
        {
            foreach (PowerPoint.Presentation p in app.Presentations)
            {
                if (p.Name == fileName)
                {
                    p.Save();
                    //TODO: CONFIGURE DIRECTORY TO SAVE FILES
                }
            }
        }
    }
}