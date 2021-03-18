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
        public event EventHandler AppClosed;

        public PowerPointApplication(Presentation Pre)
        {
            FileDirectory = Pre.Path;
            FullName = $"{AppType} - {Pre.FullName}";
            this.Pre = Pre;
            Pre.Application.WindowDeactivate += Application_WindowDeactivate;
        }

        private void Application_WindowDeactivate(Presentation Pres, DocumentWindow Wn)
        {
            if(Pres.FullName == Pre.FullName || Pre == null)
            {
                AppClosed?.Invoke(this, EventArgs.Empty);
            }
        }

        public void Save()
        {
            Pre.Save();
        }

        public string FileDirectory { get; set; }
        public string FullName { get; set; }
        public Presentation Pre { get; set; }
        public OfficeAppType AppType { get; private set; } = OfficeAppType.PowerPoint;
    }
}