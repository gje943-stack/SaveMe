namespace src.Services
{
    public interface IOfficeApplicationService
    {
        void PopulateAllOpenOfficeApplications();
        void PopulateOpenExcelApplications();
        void PopulateOpenPowerPointApplications();
        void PopulateOpenWordApplications();
    }
}