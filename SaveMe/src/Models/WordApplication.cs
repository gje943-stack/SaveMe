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
        private IEventAggregator _ea;
        public WordApplication(Document Doc, IEventAggregator ea)
        {
            this.Doc = Doc;
            FileDirectory = Doc.Path;
            FullName = Doc.FullName;
            _ea = ea;
            Doc.Application.WindowDeactivate += Application_WindowDeactivate;
        }

        private void Application_WindowDeactivate(Document Doc, Window Wn)
        {
            _ea.GetEvent<OfficeAppClosedEvent>().Publish(this);
        }

        public string FileDirectory { get; set; }
        public string FullName { get; set; }
        public OfficeAppTypes AppType { get; set; } = OfficeAppTypes.Word;

        public Document Doc { get; set; }

        public event EventHandler AppClosed;
    }
}
