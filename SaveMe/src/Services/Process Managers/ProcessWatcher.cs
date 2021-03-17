using Prism.Events;
using src.Events;
using src.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Text;

namespace src.Services.Process_Managers
{
    public class ProcessWatcher : IProcessWatcher
    {
        private readonly ManagementEventWatcher _processStartEvent = new("SELECT * FROM Win32_ProcessStartTrace");
        public List<string> ProcessNames { get; set; } = new() { "EXCEL", "POWERPNT", "WORD" };

        public event EventHandler ProcessStartedEvent;

        public ProcessWatcher()
        {
            _processStartEvent.EventArrived += _processStartEvent_EventArrived;
            _processStartEvent.Start();
        }

        private void _processStartEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            var newProcessName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            var appStartedName = ProcessNames.Find(c => newProcessName.Contains(c));
            if (appStartedName.Length > 0)
            {
                ProcessStartedEvent?.Invoke(appStartedName, EventArgs.Empty);
            }
        }
    }
}