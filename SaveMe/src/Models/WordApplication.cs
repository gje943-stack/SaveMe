using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Word;
using Prism.Events;
using src.Events;

namespace src.Models
{
    public class WordApplication : IOfficeApplication
    {

        public WordApplication(Document Doc)
        {
            this.Doc = Doc;
            FullName = Doc.FullName;
            FileDirectory = Doc.Path;
            Doc.Application.WindowDeactivate += Application_WindowDeactivate;
        }

        public event EventHandler AppClosed;

        private void Application_WindowDeactivate(Document Doc, Window Wn)
        {
            AppClosed?.Invoke(this, EventArgs.Empty);
        }

        public void Save()
        {
            Doc.Save();
        }

        public string FileDirectory { get; set; }
        public string FullName { get; set; }
        public Document Doc { get; set; }

        public OfficeAppType AppType { get; private set; } = OfficeAppType.Word;
    }
}