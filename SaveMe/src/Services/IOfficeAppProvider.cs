using src.Events;
using src.Models;
using System;
using System.Collections.Generic;

namespace src.Services
{
    public interface IOfficeAppProvider
    {
        event EventHandler<AppClosedEventArgs> AppClosedEvent;
        event EventHandler<AppOpenedEventArgs> AppOpenedEvent;

        IEnumerable<IOfficeApp> FetchNewExcelApplications(List<string> openApps);
        IEnumerable<IOfficeApp> FetchNewPowerPointApplications(List<string> openApps);
        IEnumerable<IOfficeApp> FetchNewWordApplications(List<string> openApps);
    }
}