using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace hwmonitor
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            if (!Environment.UserInteractive)
            {
                // Startup as service.
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                    new HWMonService()
                };
                ServiceBase.Run(ServicesToRun);
            }
            else
            {
                // Startup as application
                var service = new HWMonService();
                service.StartAsApp(args);
                Thread.CurrentThread.Join();
            }
        }
    }
}
