using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace src.Services.Helpers
{
    public class PowerPointHelper : IOfficeHelper<PowerPointHelper>
    {

        private Presentations OpenPowerPointInstances;
        public string OfficeAppName { get; } = "PowerPoint";

        public PowerPointHelper()
        {
            SetOpenApplicationInstances();
        }

        public void SetOpenApplicationInstances()
        {
            var openApps = GetOpenApplicationInstances();
            OpenPowerPointInstances = openApps.Presentations;
        }

        private Application GetOpenApplicationInstances()
        {
            return (Application)Marshal2.GetActiveObject("PowerPoint.Application");
        }

        public void SaveApplication(string name)
        {
            foreach(Presentation p in OpenPowerPointInstances)
            {
                if(p.Name.ToString() == name)
                {
                    p.Save();
                }
            }
        }

        public IEnumerable<(string appName, string instanceName)> GetOpenApplicationNames()
        {
            foreach(Presentation p in OpenPowerPointInstances)
            {
                yield return (OfficeAppName, p.Name);
            }
        }

    }
}
