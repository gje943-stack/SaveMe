using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace src.Services.Helpers
{
    public class WordHelper : IOfficeHelper<WordHelper>
    {
        private Documents OpenWordInstances;
        public string OfficeAppName { get; } = "Word";

        public WordHelper()
        {
            SetOpenApplicationInstances();
        }

        public void SetOpenApplicationInstances()
        {
            var openApps = GetOpenWordInstances();
            OpenWordInstances = openApps.Documents;
        }

        private Application GetOpenWordInstances()
        {
            return (Application)Marshal2.GetActiveObject("Word.Application");
        }

        public void SaveApplication(string name)
        {
            foreach (Document wb in OpenWordInstances)
            {
                if (wb.Name.ToString() == name)
                {
                    wb.Save();
                }
            }
        }

        public IEnumerable<(string appName, string instanceName)> GetOpenApplicationNames()
        {
            foreach(Document d in OpenWordInstances)
            {
                yield return (OfficeAppName, d.Name);
            }
        }

    }
}
