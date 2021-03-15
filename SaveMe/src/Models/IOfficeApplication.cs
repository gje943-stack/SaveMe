using System;
using System.Collections.Generic;
using System.Text;

namespace src.Models
{
    public interface IOfficeApplication
    {
        public string FileDirectory { get; set; }

        public string FullName { get; set; }

        public OfficeAppTypes AppType { get; set; }

        event EventHandler AppClosed;
    }
}
