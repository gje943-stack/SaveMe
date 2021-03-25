using src.Events;
using System;
using System.Management;

namespace src.Services
{
    public interface IProcessWatcher
    {
        event EventHandler<NewProcessEventArgs> NewProcessEvent;
    }
}