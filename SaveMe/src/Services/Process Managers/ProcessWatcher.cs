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

namespace src.Services.Process_Managers
{
    public class ProcessWatcher : IProcessWatcher
    {
        private readonly ManagementEventWatcher _processStartEvent = new("SELECT * FROM Win32_ProcessStartTrace");
        private Dictionary<string, Action> NewProcessEvents { get; set; }

        public event EventHandler<OfficeAppOpenedEventArgs> AppStartedEvent;

        public ProcessWatcher()
        {
            _processStartEvent.EventArrived += _processStartEvent_EventArrived;
            _processStartEvent.Start();
            NewProcessEvents = new()
            {
                { "EXCEL", () => AppStartedEvent?.Invoke(this, new OfficeAppOpenedEventArgs(OfficeAppType.Excel)) },
                { "POWERPNT", () => AppStartedEvent?.Invoke(this, new OfficeAppOpenedEventArgs(OfficeAppType.PowerPoint)) },
                { "WORD", () => AppStartedEvent?.Invoke(this, new OfficeAppOpenedEventArgs(OfficeAppType.Word)) }
            };
        }

        private void _processStartEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            var newProcessName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            var newAppName = NewProcessEvents.Keys.ToList().Find(c => newProcessName.Contains(c));
            if (newAppName != null)
            {
                NewProcessEvents[newAppName]();
            }
        }
    }
}