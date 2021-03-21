using src.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace src.Events
{
    public class NewProcessEventArgs : EventArgs
    {
        public OfficeAppType ProcessType { get; set; }

        public NewProcessEventArgs(OfficeAppType processType)
        {
            ProcessType = processType;
        }
    }
}