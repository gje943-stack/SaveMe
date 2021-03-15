using Prism.Events;
using src.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace src.Events
{
    public class OfficeAppClosedEvent : PubSubEvent<IOfficeApplication> { }
}