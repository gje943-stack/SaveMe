using src.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace src.Services
{
    public interface IOfficeProcessDetector
    {
        List<Process> FetchOpenOfficeProcesses();
    }
}