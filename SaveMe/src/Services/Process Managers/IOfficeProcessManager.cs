using src.Models;
using System.Collections.Generic;

namespace src.Services.Process_Managers
{
    public interface IOfficeProcessManager
    {
        IEnumerable<IOfficeApplication> FetchOpenExcelProcesses();
        IEnumerable<IOfficeApplication> FetchOpenPowerPointApplications();
        IEnumerable<IOfficeApplication> FetchOpenWordProcesses();
    }
}