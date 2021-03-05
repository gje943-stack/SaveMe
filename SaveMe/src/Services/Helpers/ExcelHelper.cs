using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace src.Services.Helpers
{
    /// <summary>
    ///     Interacts with all currently open Excel instances.
    /// </summary>
    public class ExcelHelper : IOfficeHelper<ExcelHelper>
    {
        public Workbooks OpenExcelInstances { get; set; }
        public string OfficeAppName { get; } = "Excel";

        public ExcelHelper()
        {
            SetOpenApplicationInstances();
        }

        public void SaveApplication(string name)
        {
            foreach (Workbook wb in OpenExcelInstances)
            {
                if (wb.Name.ToString() == name)
                {
                    wb.Save();
                }
            }
        }

        private Application GetOpenExcelInstances()
        {
            return (Application)Marshal2.GetActiveObject("Excel.Application");
        }

        public void SetOpenApplicationInstances()
        {
            var openInstances = GetOpenExcelInstances();
            OpenExcelInstances = openInstances.Workbooks;
        }

        public IEnumerable<(string appName, string instanceName)> GetOpenApplicationNames()
        {
            foreach (Workbook wb in OpenExcelInstances)
            {
                yield return (OfficeAppName, wb.Name);
            }
        }
    }
}