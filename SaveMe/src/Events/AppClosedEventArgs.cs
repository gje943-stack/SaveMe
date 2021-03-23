using Prism.Events;
using src.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace src.Events
{
    public class AppClosedEventArgs : EventArgs
    {
        public string Name { get; set; }
        public OfficeAppType Type { get; set; }

        public AppClosedEventArgs(string name, OfficeAppType type)
        {
            Name = name;
            Type = type;
        }
    }
}