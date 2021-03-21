using System;
using System.Collections.Generic;
using System.Text;

namespace src.Models
{
    public class OfficeApp : IOfficeApp
    {
        public OfficeAppType Type { get; set; }

        public string Name { get; set; }

        public OfficeApp(OfficeAppType type, string name)
        {
            Type = type;
            Name = name;
        }
    }
}