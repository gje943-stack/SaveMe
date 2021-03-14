using Prism.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace src.Events
{
    public class ExcelAppOpenedEvent : PubSubEvent<string> { }
}