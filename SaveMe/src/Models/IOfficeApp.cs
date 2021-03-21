using System;
using System.Collections.Generic;
using System.Text;

namespace src.Models
{
    public interface IOfficeApp
    {
        public string Name { get; set; }
        OfficeAppType Type { get; }
    }
}