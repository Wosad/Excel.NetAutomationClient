

using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Wosad.Excel.NetAutomationClient.Demo
{
    public class ExcelNetClientDemoBootstrapper : BootstrapperBase
    {
        public ExcelNetClientDemoBootstrapper()
        {
            Initialize();
        }


        protected override void OnStartup(object sender, StartupEventArgs e)
        {
            DisplayRootViewFor<ExcelNetClientDemoViewModel>();
        }
    }

}
