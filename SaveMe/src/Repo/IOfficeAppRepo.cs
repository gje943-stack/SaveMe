using src.Models;
using System.Collections.Generic;

namespace src.Repo
{
    public interface IOfficeAppRepo
    {
        List<IOfficeApplication> OpenOfficeApps { get; }

        bool Delete(IOfficeApplication appToDelete);
        void Insert(IOfficeApplication newApp);
    }
}