using src.Models;
using System;
using System.Collections.Generic;

namespace src.Services
{
    public interface IOfficeAppService
    {
        void AddOpenOfficeApplicationsToRepo(Func<IEnumerable<IOfficeApplication>> getApps);
        List<string> GetOpenAppNames();
    }
}