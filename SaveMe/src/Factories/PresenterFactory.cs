using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using src.Presenters;
using System;
using System.Collections.Generic;
using System.Text;

namespace src.Factories
{
    public static class PresenterFactory
    {
        private static IHost _host;
        public static IPresenter<TView> CreateForView<TView>(TView view)
        {
            var presenter = ActivatorUtilities.GetServiceOrCreateInstance<IPresenter<TView>>(_host.Services);
            presenter.View = view;
            presenter.InitialSetup();
            return presenter;
        }
    
        public static void SetHost(IHost host)
        {
            _host = host;
        }
    }
}
