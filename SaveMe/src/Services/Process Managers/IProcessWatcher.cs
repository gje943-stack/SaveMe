using src.Events;
using System;
using System.Collections.Generic;

namespace src.Services.Process_Managers
{
    public interface IProcessWatcher
    {
        event EventHandler<OfficeAppOpenedEventArgs> AppStartedEvent;
    }
}