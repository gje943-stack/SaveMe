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

        IOfficeApp FetchNewExcelApplication(int numAppsAlreadyKnown);
        IOfficeApp FetchNewPowerPointApplication(int numAppsAlreadyKnown);
        IOfficeApp FetchNewWordApplication(int numAppsAlreadyKnown);
        IEnumerable<IOfficeApp> TryFetchExcelApplicationsOnStartup();
        IEnumerable<IOfficeApp> TryFetchPowerPointApplicationsOnStartup();
        IEnumerable<IOfficeApp> TryFetchWordApplicationsOnStartup();
    }
}