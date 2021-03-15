using System;
using System.Collections.Generic;
using System.Management;

namespace src.Services.Process_Managers
{
    public interface IProcessWatcher
    {
        List<string> ProcessNames { get; set; }

        event EventHandler ProcessStartedEvent;
    }
}