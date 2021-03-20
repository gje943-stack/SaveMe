using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace src.Services.Process_Managers
{
    public interface IOfficeApplicationProvider
    {
        IEnumerable<Workbook> FetchOpenExcelApplications();
        IEnumerable<Presentation> FetchOpenPowerPointApplications();
        IEnumerable<Document> FetchOpenWordApplications();
    }
}