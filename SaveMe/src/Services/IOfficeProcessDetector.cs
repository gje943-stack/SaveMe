using System.Collections.Generic;
using System.Diagnostics;

namespace src.Services
{
    public interface IOfficeProcessDetector
    {
        IEnumerable<Process> FetchOpenOfficeProcesses();
    }
}