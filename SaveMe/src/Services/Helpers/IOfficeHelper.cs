using System.Collections.Generic;

namespace src.Services.Helpers
{
    public interface IOfficeHelper<TOfficeApplication> where TOfficeApplication : class
    {
        string OfficeAppName { get; }

        IEnumerable<(string appName, string instanceName)> GetOpenApplicationNames();

        void SaveApplication(string name);

        void SetOpenApplicationInstances();
    }
}