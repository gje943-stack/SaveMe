using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using src.Models;

namespace src.Services.Providers
{
    public interface IOfficeAppProvider
    {
        Document GenerateNewDocument(string processName);
        Workbook GenerateNewWorkbook(string processName);
    }
}