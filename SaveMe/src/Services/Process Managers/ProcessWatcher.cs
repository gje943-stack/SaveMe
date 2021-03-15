using Prism.Events;
using src.Events;
using src.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;

namespace src.Services.Process_Managers
{
    public class ProcessWatcher : IProcessWatcher
    {
        private readonly ManagementEventWatcher _processStartEvent = new("SELECT * FROM Win32_ProcessStartTrace");
        public List<string> ProcessNames { get; set; } = new(){ "Excel", "PowerPoint", "Word" };
        public event EventHandler ProcessStartedEvent;

        public ProcessWatcher(IEventAggregator ea)
        {
            _processStartEvent.EventArrived += _processStartEvent_EventArrived;
            _processStartEvent.Start();
        }

        private void _processStartEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            var newProcessName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            if (ProcessNames.Any(c => newProcessName.Contains(c)))
            {
                ProcessStartedEvent?.Invoke(newProcessName, EventArgs.Empty);
            }
        }
    }
}
