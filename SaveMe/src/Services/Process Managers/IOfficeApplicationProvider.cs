using src.Models;
using System.Collections.Generic;

namespace src.Services.Process_Managers
{
    public interface IOfficeApplicationProvider
    {
        IEnumerable<IOfficeApplication> FetchOpenExcelApplications();
        IEnumerable<IOfficeApplication> FetchOpenPowerPointApplications();
        IEnumerable<IOfficeApplication> FetchOpenWordApplications();
    }
}