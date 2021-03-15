using src.Models;
using System.Collections.Generic;

namespace src.Repo
{
    public interface IOfficeAppRepo
    {
        List<IOfficeApplication> OpenOfficeProcesses { get; }

        bool Delete(IOfficeApplication appToDelete);
        void Insert(IOfficeApplication newApp);
    }
}