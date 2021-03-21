using Prism.Events;
using src.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace src.Events
{
    public class AppOpenedEventArgs : EventArgs
    {
        public OfficeAppType AppType { get; set; }

        public AppOpenedEventArgs(OfficeAppType appType)
        {
            AppType = appType;
        }
    }
}