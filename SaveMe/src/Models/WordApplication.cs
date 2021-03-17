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

        public WordApplication(Document Doc, IEventAggregator _ea)
        {
            this.Doc = Doc;
            FullName = $"{AppType} - {Doc.FullName}";
            FileDirectory = Doc.Path;
            this._ea = _ea;
            Doc.Application.WindowDeactivate += Application_WindowDeactivate;
        }

        private void Application_WindowDeactivate(Document Doc, Window Wn)
        {
            _ea.GetEvent<OfficeAppClosedEvent>().Publish(this);
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