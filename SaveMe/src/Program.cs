using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using src.Factories;
using src.Presenters;
using src.Services;
using System;
using System.Windows.Forms;

namespace src
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var host = Host.CreateDefaultBuilder()
                .ConfigureServices((services) =>
                {
                    services.AddTransient<IPresenter<MainFormView>, MainFormPresenter<MainFormView>>();
                    services.AddTransient<IOfficeAppProvider, OfficeAppProvider>();
                    services.AddSingleton<IMainFormView, MainFormView>();
                    services.AddSingleton<IProcessWatcher, ProcessWatcher>();
                })
                .Build();

            PresenterFactory.SetHost(host);
            var app = ActivatorUtilities.GetServiceOrCreateInstance<MainFormView>(host.Services);
            Application.Run(app);
        }
    }
}