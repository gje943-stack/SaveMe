using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using src.Models;

namespace src.Services.Providers
{
    public interface IOfficeAppProvider
    {
        Document GenerateNewDocumentFromName(string processName);
        Presentation GenerateNewPresentationFromName(string processName);
        Workbook GenerateNewWorkbookFromName(string processName);
    }
}