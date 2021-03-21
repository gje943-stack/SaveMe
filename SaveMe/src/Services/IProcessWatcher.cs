using src.Events;
using System;

namespace src.Services
{
    public interface IProcessWatcher
    {
        event EventHandler<NewProcessEventArgs> NewProcessEvent;
    }
}