using Prism.Events;
using src.Events;
using src.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;

namespace src.Services
{
    public class ProcessWatcher : IProcessWatcher
    {
        private readonly ManagementEventWatcher _processStartEvent = new("SELECT * FROM Win32_ProcessStartTrace");
        private Dictionary<string, Action> NewProcessEvents { get; set; }

        public event EventHandler<NewProcessEventArgs> NewProcessEvent;

        public ProcessWatcher()
        {
            _processStartEvent.EventArrived += _processStartEvent_EventArrived;
            _processStartEvent.Start();
            NewProcessEvents = new()
            {
                { "EXCEL", () => NewProcessEvent?.Invoke(this, new NewProcessEventArgs(OfficeAppType.Excel)) },
                { "POWERPNT", () => NewProcessEvent?.Invoke(this, new NewProcessEventArgs(OfficeAppType.PowerPoint)) },
                { "WORD", () => NewProcessEvent?.Invoke(this, new NewProcessEventArgs(OfficeAppType.Word)) }
            };
        }

        private void _processStartEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            var newProcessName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            var newAppName = NewProcessEvents.Keys.ToList().Find(c => newProcessName.Contains(c));
            var newProcessId = Convert.ToInt32(e.NewEvent.Properties["ProcessId"].Value);
            if (newAppName != null)
            {
                NewProcessEvents[newAppName]();
            }
        }
    }
}