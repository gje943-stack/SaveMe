using Prism.Events;
using src.Events;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management;
using System.Text;
using System.Threading;

namespace src.Services
{
    public class ExcelProcessManager : IOfficeProcessManager<ExcelProcessManager>
    {
        private readonly ManagementEventWatcher _processStartEvent = new("SELECT * FROM Win32_ProcessStartTrace");
        private readonly ManagementEventWatcher _processStopEvent = new("SELECT * FROM Win32_ProcessStopTrace");
        private IEventAggregator _ea;

        public ExcelProcessManager(IEventAggregator ea)
        {
            //_processStartEvent.EventArrived += _processStartEvent_EventArrived;
            //_processStopEvent.EventArrived += _processStopEvent_EventArrived;
            _processStartEvent.Start();
            _processStopEvent.Start();
            _ea = ea;
        }

        private void _processStopEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            (var closedProcessName, var closedProcessId) = GetDataFromProcessEvent(e);
            if (closedProcessName.Contains("Excel"))
            {
                _ea.GetEvent<ExcelAppClosedEvent>().Publish(closedProcessName);
            }
        }

        private void _processStartEvent_EventArrived(object sender, EventArrivedEventArgs e)
        {
            (var newProcessName, var newProcessId) = GetDataFromProcessEvent(e);
            if (newProcessName.Contains("Excel"))
            {
                var newProcess = Process.GetProcessById(newProcessId);
                newProcess.WaitForInputIdle();
                _ea.GetEvent<ExcelAppOpenedEvent>().Publish(newProcessName);
            }
        }

        public (string, int) GetDataFromProcessEvent(EventArrivedEventArgs e)
        {
            var processName = e.NewEvent.Properties["ProcessName"].Value.ToString();
            var processId = Int32.Parse(e.NewEvent.Properties["ProcessId"].Value.ToString());
            return (processName, processId);
        }

        public List<string> FetchOpenProcesses()
        {
            var res = new List<string>();
            var app = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
            foreach (Excel.Workbook wb in app.Workbooks)
            {
                res.Add(wb.FullName);
            }
            return res;
        }
    }
}