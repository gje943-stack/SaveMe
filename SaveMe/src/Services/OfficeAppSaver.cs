using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using src.Models;
using System.Threading.Tasks;

namespace src.Services
{
    public static class OfficeAppSaver
    {
        private static void SaveSingleWorkbook(IOfficeApp appToSave)
        {
            var app = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            foreach(Excel.Workbook wb in app.Workbooks)
            {
                if(wb.FullName == appToSave.Name)
                {
                    wb.SaveCopyAs(GenerateFileDirectory(wb.FullName));
                }
            }
        }

        private static void SaveSingleDocument(IOfficeApp appToSave)
        {
            var app = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            foreach(Word.Document doc in app.Documents)
            {
                if(doc.FullName == appToSave.Name)
                {
                    doc.SaveCopyAs(GenerateFileDirectory(doc.FullName));
                }
            }
        }

        private static void SaveSinglePresentation(IOfficeApp appToSave)
        {
            var app = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            foreach(PowerPoint.Presentation pres in app.Presentations)
            {
                if(pres.FullName == appToSave.Name)
                {
                    pres.SaveCopyAs(GenerateFileDirectory(pres.FullName));
                }
            }
        }

        public static void SaveSingleApplication(IOfficeApp appToSave)
        {
            switch (appToSave.Type)
            {
                case OfficeAppType.Word:
                    SaveSingleDocument(appToSave);
                    break;
                case OfficeAppType.Excel:
                    SaveSingleWorkbook(appToSave);
                    break;
                case OfficeAppType.PowerPoint:
                    SaveSinglePresentation(appToSave);
                    break;
            }
        }

        private static void TrySaveAllWordApps()
        {
            var app = (Word.Application)Marshal2.GetActiveObject("Word.Application");
            if(app != null)
            {
                foreach (Word.Document doc in app.Documents)
                {
                    doc.SaveCopyAs(GenerateFileDirectory($"{doc.FullName}.docx"));
                }
            }
        }

        private static void TrySaveAllPowerPointApps()
        {
            var app = (PowerPoint.Application)Marshal2.GetActiveObject("PowerPoint.Application");
            if (app != null)
            {
                foreach (PowerPoint.Presentation pres in app.Presentations)
                {
                    pres.SaveCopyAs(GenerateFileDirectory($"{pres.FullName}.ppt"));
                }
            }
        }

        private static void TrySaveAllExcelApps()
        {
            var app = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            if (app != null)
            {
                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    wb.SaveCopyAs(GenerateFileDirectory($"{wb.FullName}.xlsx"));
                }
            }
        }

        public static void SaveAll()
        {
            TrySaveAllExcelApps();
            TrySaveAllWordApps();
            TrySaveAllPowerPointApps();
        }

        private static string GenerateFileDirectory(string fileName)
        {
            var root = VaultLocationConfig.VaultLocation;
            var uniqueId = Guid.NewGuid().ToString();
            return $@"{root}\{fileName}-{uniqueId}";
        }
    }
}
