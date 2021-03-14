using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Prism.Events;
using src.Factories;
using src.Presenters;
using src.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
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
                    services.AddTransient<IOfficeAppService, OfficeAppService>();
                    services.AddTransient<IOfficeProcessManager<ExcelProcessManager>, ExcelProcessManager>();
                    services.AddTransient<IOfficeApplicationSaver, OfficeApplicationSaver>();
                    services.AddSingleton<IMainFormView, MainFormView>();
                    services.AddSingleton<IEventAggregator, EventAggregator>();
                })
                .Build();

            PresenterFactory.SetHost(host);
            var app = ActivatorUtilities.GetServiceOrCreateInstance<MainFormView>(host.Services);
            Application.Run(app);
        }
    }
}