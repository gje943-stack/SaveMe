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
        public ExcelApplication(Workbook Wb, IEventAggregator ea)
        {
            FileDirectory = FileDirectory;
            FullName = FullName;
            this.Wb = Wb;
            Wb.Application.WindowDeactivate += Application_WindowDeactivate; ;
            _ea = ea;
        }

        private void Application_WindowDeactivate(Workbook Wb, Window Wn)
        {
            _ea.GetEvent<OfficeAppClosedEvent>().Publish(this);
        }

        public string FileDirectory { get; set; }
        public string FullName { get; set; }
        public OfficeAppTypes AppType { get; set; } = OfficeAppTypes.Excel;

        public Workbook Wb { get; set; }
    }
}
