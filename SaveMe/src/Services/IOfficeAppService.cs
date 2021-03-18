using src.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace src.Services
{
    public interface IOfficeAppService
    {
        List<IOfficeApplication> OpenOfficeApps { get; set; }

        event EventHandler NewAppStartedEvent;
        event EventHandler AppClosedEvent;

        List<string> GetOpenAppNames();
        Task SaveApps(List<string> appsToSave);
        void SetOpenOfficeApplications(Func<IEnumerable<IOfficeApplication>> getApps);
    }
}