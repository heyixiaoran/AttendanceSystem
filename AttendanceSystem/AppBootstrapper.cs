using System;
using System.Collections.Generic;
using System.Windows;

using AttendanceSystem.ViewModels;
using AttendanceSystem.Views;

using Caliburn.Micro;

namespace AttendanceSystem
{


    public class AppBootstrapper : BootstrapperBase
    {
        SimpleContainer container;

        public AppBootstrapper()
        {
            Initialize();
        }

        protected override void Configure()
        {
            container = new SimpleContainer();

            container.Singleton<IWindowManager, WindowManager>();
            container.Singleton<IEventAggregator, EventAggregator>();
            container.PerRequest<IShell, MainViewModel>();
            container.PerRequest<AddViewModel>();
        }

        protected override object GetInstance(Type service, string key)
        {
            var instance = container.GetInstance(service, key);
            if(instance != null)
            {
                return instance;
            }

            throw new InvalidOperationException("Could not locate any instances.");
        }

        protected override IEnumerable<object> GetAllInstances(Type service)
        {
            return container.GetAllInstances(service);
        }

        protected override void BuildUp(object instance)
        {
            container.BuildUp(instance);
        }

        protected override void OnStartup(object sender, StartupEventArgs e)
        {
            DisplayRootViewFor<IShell>();
        }
    }
}