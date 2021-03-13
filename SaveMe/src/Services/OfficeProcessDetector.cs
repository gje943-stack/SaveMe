using Prism.Events;
using src.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management;
using System.Text;
using System.Threading;

namespace src.Services
{
    public class OfficeProcessDetector : IOfficeProcessDetector
    {
        private readonly ManagementEventWatcher _processStartEvent = new("SELECT * FROM Win32_ProcessStartTrace");
        private readonly ManagementEventWatcher _processStopEvent = new("SELECT * FROM Win32_ProcessStopTrace");
        private IEventAggregator _ea;

        public OfficeProcessDetector(IEventAggregator ea)
        {
            _processStartEvent.EventArrived += _processStartEvent_EventArrived;
            _processStopEvent.EventArrived += _processStopEvent_EventArrived;
            _processStartEvent.Start();
            _processStopEvent.Start();
            _ea = ea;
        }

        private void _processStopEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            (var closedProcessName, var closedProcessId) = GetDataFromProcessEvent(e);
            if (OfficeProcesses.NamesWithExe.Contains(closedProcessName))
            {
                var closedProcess = Process.GetProcessById(closedProcessId);
                _ea.GetEvent<OfficeAppClosedEvent>().Publish(closedProcess);
            }
        }

        private void _processStartEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            (var newProcessName, var newProcessId) = GetDataFromProcessEvent(e);
            if (OfficeProcesses.NamesWithExe.Contains(newProcessName))
            {
                Thread.Sleep(5000);
                var newProcess = Process.GetProcessById(newProcessId);
                _ea.GetEvent<OfficeAppOpenedEvent>().Publish(newProcess);
            }
        }

        private (string, int) GetDataFromProcessEvent(EventArrivedEventArgs e)
        {
            var processName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            var processId = Int32.Parse(e.NewEvent.Properties["ProcessId"].Value.ToString());
            return (processName, processId);
        }

        public List<Process> FetchOpenOfficeProcesses()
        {
            var processes = new List<Process>();
            foreach (var process in Process.GetProcesses())
            {
                if (OfficeProcesses.NamesWithoutExe.Contains(process.ProcessName))
                {
                    processes.Add(process);
                }
            }
            return processes;
        }
    }
}