using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using LibreHardwareMonitor.Hardware;
using LibreHardwareMonitor.Hardware.Cpu;

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
                if (args.Length == 0)
                {
                    Console.WriteLine("Argument required! Available:\n\t-sensors\n\t-printsensors\n\t-testdevice");
                    return;
                }
                if (args.Length >0 && args[0].Equals("-sensors"))
                {
                    service.UpdateComputerHardware();

                    walkHardwareSensors(service.Computer.Hardware,  null);
                }
                if (args.Length > 0 && args[0].Equals("-printsensors"))
                {
                    service.UpdateComputerHardware();

                    walkHardwareSensors(service.Computer.Hardware, new List<String> { service.SensorIdCpu, service.SensorIdCpuTemp, service.SensorIdMem, service.SensorIdMemValue, service.SensorIdGpu1, service.SensorIdGpu1Temp, service.SensorIdGpu2, service.SensorIdGpu2Temp });
                }
                if (args.Length > 0 && args[0].Equals("-testdevice"))
                {
                    service.StartAsApp(args);
                    service.Thread.Join();
                }
            }
        }

        static void walkHardwareSensors(IList<IHardware> hardlist, IList<String> searchSensors)
        {
            foreach (IHardware hardware in hardlist)
            {
                if(searchSensors == null) Console.WriteLine($"Hardware Name:{hardware.Name} Id:{hardware.Identifier}");
                printSensors(hardware, searchSensors);
                walkHardwareSensors(hardware.SubHardware, searchSensors);
            }
        }

        static void printSensors(IHardware hardware, IList<String> searchSensors)
        {
            foreach (ISensor sensor in hardware.Sensors)
            {
                if((searchSensors == null) || (searchSensors.Contains(sensor.Identifier.ToString())))
                    Console.WriteLine($"Sensor for {hardware.Name} Name:{sensor.Name} Id:{sensor.Identifier} Value:{sensor.Value}");
            }
        }
    }
}
