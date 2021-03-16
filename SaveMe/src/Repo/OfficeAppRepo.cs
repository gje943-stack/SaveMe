using src.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace src.Repo
{
    public class OfficeAppRepo : IOfficeAppRepo
    {
        public List<IOfficeApplication> OpenOfficeApps { get; private set; } = new();

        public void Insert(IOfficeApplication newApp)
        {
            OpenOfficeApps.Add(newApp);
        }

        public bool Delete(IOfficeApplication appToDelete)
        {
            return OpenOfficeApps.Remove(appToDelete);
        }
    }
}
