using System.Collections.Generic;
using System.Management;

namespace src.Services
{
    public interface IOfficeProcessManager<TProcessManager> where TProcessManager : IOfficeProcessManager<TProcessManager>
    {
        List<string> FetchOpenProcesses();

        (string, int) GetDataFromProcessEvent(EventArrivedEventArgs e);
    }
}