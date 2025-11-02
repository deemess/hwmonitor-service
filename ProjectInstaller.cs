using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace hwmonitor
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            // ServiceProcessInstaller defines the account under which the service runs
            var processInstaller = new ServiceProcessInstaller
            {
                Account = ServiceAccount.LocalSystem // or NetworkService, etc.
            };

            // ServiceInstaller defines service specific settings
            var serviceInstaller = new ServiceInstaller
            {
                ServiceName = "HWMonitor",
                DisplayName = "HWMonitor Service",
                Description = "HWMonitor Service. Require HWMonitor Device.",
                StartType = ServiceStartMode.Automatic
            };

            Installers.Add(processInstaller);
            Installers.Add(serviceInstaller);
        }
    }
}
