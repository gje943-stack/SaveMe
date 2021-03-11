using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management;
using System.Text;

namespace src.Services
{
    public class OfficeProcessDetector : IOfficeProcessDetector
    {
        private readonly ManagementEventWatcher _processStartEvent = new("SELECT * FROM Win32_ProcessStartTrace");
        private readonly ManagementEventWatcher _processStopEvent = new("SELECT * FROM Win32_ProcessStopTrace");
        public event EventHandler ApplicationClosedEvent;

        public OfficeProcessDetector()
        {
            _processStartEvent.EventArrived += _processStartEvent_EventArrived;
            _processStopEvent.EventArrived += _processStopEvent_EventArrived;
            _processStartEvent.Start();
        }

        private void _processStopEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            var processName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            if ()
        }

        private void _processStartEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            var processName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            
        }

        public IEnumerable<Process> FetchOpenOfficeProcesses()
        {
            foreach(var appType in Enum.GetValues(typeof(OfficeApplicationType)))
            {
                foreach(var process in Process.GetProcessesByName($"{appType}.exe"))
                {
                    yield return process;
                }
            }
        }
    }
}
