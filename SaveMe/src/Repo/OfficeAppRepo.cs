using src.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace src.Repo
{
    public class OfficeAppRepo : IOfficeAppRepo
    {
        public List<IOfficeApplication> OpenOfficeProcesses { get; private set; } = new();

        public void Insert(IOfficeApplication newApp)
        {
            OpenOfficeProcesses.Add(newApp);
        }

        public bool Delete(IOfficeApplication appToDelete)
        {
            return OpenOfficeProcesses.Remove(appToDelete);
        }
    }
}
