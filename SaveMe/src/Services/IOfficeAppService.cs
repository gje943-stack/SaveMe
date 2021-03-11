using System.Collections.Generic;
using System.Diagnostics;

namespace src.Services
{
    public interface IOfficeAppService
    {
        List<Process> OpenOfficeProcesses { get; }
    }
}