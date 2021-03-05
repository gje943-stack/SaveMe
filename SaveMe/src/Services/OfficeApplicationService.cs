using src.Services.Helpers;
using System;
using System.Collections.Generic;

namespace src.Services
{
    public class OfficeApplicationService : IOfficeApplicationService
    {
        private readonly IOfficeHelper<ExcelHelper> _excelHelper;
        private readonly IOfficeHelper<PowerPointHelper> _powerPointHelper;
        private readonly IOfficeHelper<WordHelper> _wordHelper;
        private Dictionary<string, string> OpenOfficeApplications = new Dictionary<string, string>();

        public OfficeApplicationService(IOfficeHelper<ExcelHelper> excelHelper, IOfficeHelper<WordHelper> wordHelper, IOfficeHelper<PowerPointHelper> powerPointHelper)
        {
            _excelHelper = excelHelper;
            _wordHelper = wordHelper;
            _powerPointHelper = powerPointHelper;
            PopulateAllOpenOfficeApplications();
        }

        public void PopulateAllOpenOfficeApplications()
        {
            PopulateOpenExcelApplications();
            PopulateOpenPowerPointApplications();
            PopulateOpenWordApplications();
        }

        public void PopulateOpenExcelApplications()
        {
            foreach (var (appName, instanceName) in _excelHelper.GetOpenApplicationNames())
            {
                OpenOfficeApplications.Add(appName, instanceName);
            }
        }

        public void PopulateOpenWordApplications()
        {
            foreach (var (appName, instanceName) in _wordHelper.GetOpenApplicationNames())
            {
                OpenOfficeApplications.Add(appName, instanceName);
            }
        }

        public void PopulateOpenPowerPointApplications()
        {
            foreach (var (appName, instanceName) in _powerPointHelper.GetOpenApplicationNames())
            {
                OpenOfficeApplications.Add(appName, instanceName);
            }
        }
    }
}