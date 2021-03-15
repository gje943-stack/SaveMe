using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Prism.Events;
using src.Events;

namespace src.Models
{
    public class PowerPointApplication : IOfficeApplication
    {
        public PowerPointApplication(Presentation Pre, IEventAggregator ea)
        {
            FileDirectory = FileDirectory;
            FullName = FullName;
            this.Pre = Pre;
            _ea = ea;
            Pre.Application.WindowDeactivate += Application_WindowDeactivate;
        }

        private void Application_WindowDeactivate(Presentation Pres, DocumentWindow Wn)
        {
            _ea.GetEvent<OfficeAppClosedEvent>().Publish(this);
        }

        public string FileDirectory { get; set; }
        public string FullName { get; set; }
        public OfficeAppTypes AppType { get; set; } = OfficeAppTypes.PowerPoint;

        public Presentation Pre { get; set; }

        public event EventHandler AppClosed;
    }
}
