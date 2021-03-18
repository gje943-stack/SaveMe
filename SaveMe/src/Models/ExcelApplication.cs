using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Prism.Events;
using src.Events;

namespace src.Models
{
    public class ExcelApplication : IOfficeApplication
    {
        private IEventAggregator _ea;

        public event EventHandler AppClosed;

        public ExcelApplication(Workbook Wb)
        {
            FileDirectory = Wb.Path;
            FullName = $"{AppType} - {Wb.FullName}";
            this.Wb = Wb;
            Wb.Application.WindowDeactivate += Application_WindowDeactivate; ;
        }

        private void Application_WindowDeactivate(Workbook Wb, Window Wn)
        {
            AppClosed?.Invoke(this, EventArgs.Empty);
        }

        public void Save()
        {
            Wb.Save();
        }

        public string FileDirectory { get; set; }
        public string FullName { get; set; }
        public Workbook Wb { get; set; }

        public OfficeAppType AppType { get; private set; } = OfficeAppType.Excel;
    }
}