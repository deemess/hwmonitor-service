using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using LibreHardwareMonitor.Hardware;
using TimeoutException = System.TimeoutException;

namespace hwmonitor
{
    public partial class HWMonService : ServiceBase
    {
        private Thread _workerThread;
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isRunning = false;
        private SerialPort _serialPort;
        private Computer _computer;
        
        // Configuration properties
        private string ComPort { get; set; }
        private int BaudRate { get; set; }
        private int DataBits { get; set; }
        private Parity Parity { get; set; }
        private StopBits StopBits { get; set; }
        private int MonitoringInterval { get; set; }

        public HWMonService()
        {
            InitializeComponent();
            LoadConfiguration();
        }

        private void LoadConfiguration()
        {
            try
            {
                ComPort = ConfigurationManager.AppSettings["ComPort"] ?? "COM1";
                BaudRate = int.Parse(ConfigurationManager.AppSettings["BaudRate"] ?? "9600");
                DataBits = int.Parse(ConfigurationManager.AppSettings["DataBits"] ?? "8");
                MonitoringInterval = int.Parse(ConfigurationManager.AppSettings["MonitoringInterval"] ?? "1000");
                
                string parityStr = ConfigurationManager.AppSettings["Parity"] ?? "None";
                Parity = (Parity)Enum.Parse(typeof(Parity), parityStr, true);
                
                string stopBitsStr = ConfigurationManager.AppSettings["StopBits"] ?? "One";
                StopBits = (StopBits)Enum.Parse(typeof(StopBits), stopBitsStr, true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading configuration: {ex.Message}");
                // Set default values
                ComPort = "COM1";
                BaudRate = 9600;
                DataBits = 8;
                Parity = Parity.None;
                StopBits = StopBits.One;
                MonitoringInterval = 1000; // 1 second default
            }
        }

        public void StartAsApp(string[] args)
        {
            this.OnStart(args);
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                // Initialize hardware monitoring
                _computer = new Computer
                {
                    IsCpuEnabled = true,
                    IsGpuEnabled = true,
                    IsMemoryEnabled = false,
                    IsMotherboardEnabled = false,
                    IsControllerEnabled = false,
                    IsNetworkEnabled = false,
                    IsStorageEnabled = false
                };
                
                _computer.Open();
                _computer.Accept(new UpdateVisitor());
                
                // Initialize serial port
                _serialPort = new SerialPort(ComPort, BaudRate, Parity, DataBits, StopBits);
                _serialPort.ReadTimeout = 1000;
                _serialPort.WriteTimeout = 1000;
                
                _cancellationTokenSource = new CancellationTokenSource();
                _isRunning = true;
                
                _workerThread = new Thread(WorkerMethod)
                {
                    IsBackground = true,
                    Name = "HwMonitorWorker"
                };
                
                _workerThread.Start();

                Debug.WriteLine("Hardware monitoring service started successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error starting service: {ex.Message}");
                throw;
            }
        }

        protected override void OnStop()
        {
            _isRunning = false;
            _cancellationTokenSource?.Cancel();
            
            if (_workerThread != null && _workerThread.IsAlive)
            {
                _workerThread.Join(5000); // Wait up to 5 seconds for thread to finish
            }
            
            // Close serial port
            if (_serialPort != null && _serialPort.IsOpen)
            {
                try
                {
                    _serialPort.Close();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error closing serial port: {ex.Message}");
                }
            }
            
            // Close hardware monitoring
            try
            {
                _computer?.Close();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error closing hardware monitor: {ex.Message}");
            }
            
            _serialPort?.Dispose();
            _cancellationTokenSource?.Dispose();

            Debug.WriteLine("Hardware monitoring service stopped");
        }

        private void WorkerMethod()
        {
            while (_isRunning && !_cancellationTokenSource.Token.IsCancellationRequested)
            {
                try
                {
                    // Open serial port if not already open
                    if (_serialPort != null && !_serialPort.IsOpen)
                    {
                        _serialPort.Open();
                        Debug.WriteLine($"Serial port {ComPort} opened successfully");
                    }
                    
                    // Perform COM port communication
                    if (_serialPort != null && _serialPort.IsOpen)
                    {
                        // Application only sends data, no reading from COM port
                        MonitorHardware();
                    }
                    
                    // Sleep for a short period to avoid excessive CPU usage
                    Thread.Sleep(100); // Sleep for 100ms
                }
                catch (TimeoutException)
                {
                    // This is normal when no data is available
                    // Continue the loop
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Worker thread error: {ex.Message}");
                    
                    // If there's a serial port error, try to close and reopen
                    if (_serialPort != null && _serialPort.IsOpen)
                    {
                        try
                        {
                            _serialPort.Close();
                        }
                        catch { }
                    }
                    
                    // Wait a bit before retrying
                    Thread.Sleep(MonitoringInterval);
                }
            }
        }

        private void MonitorHardware()
        {
            if (!_isRunning || _computer == null)
                return;
                
            try
            {
                // Update hardware sensors
                _computer.Accept(new UpdateVisitor());
                
                var hardwareData = new StringBuilder();
                hardwareData.AppendLine($"HWMON:{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                
                // Collect CPU data
                foreach (var hardware in _computer.Hardware)
                {
                    if (hardware.HardwareType == HardwareType.Cpu)
                    {
                        foreach (var sensor in hardware.Sensors)
                        {
                            if (sensor.SensorType == SensorType.Temperature && sensor.Value.HasValue)
                            {
                                hardwareData.AppendLine($"CPU_TEMP:{sensor.Name}:{sensor.Value.Value:F1}C");
                            }
                            else if (sensor.SensorType == SensorType.Load && sensor.Value.HasValue)
                            {
                                hardwareData.AppendLine($"CPU_LOAD:{sensor.Name}:{sensor.Value.Value:F1}%");
                            }
                        }
                    }
                    else if (hardware.HardwareType == HardwareType.GpuNvidia || 
                             hardware.HardwareType == HardwareType.GpuAmd || 
                             hardware.HardwareType == HardwareType.GpuIntel)
                    {
                        foreach (var sensor in hardware.Sensors)
                        {
                            if (sensor.SensorType == SensorType.Temperature && sensor.Value.HasValue)
                            {
                                hardwareData.AppendLine($"GPU_TEMP:{sensor.Name}:{sensor.Value.Value:F1}C");
                            }
                            else if (sensor.SensorType == SensorType.Load && sensor.Value.HasValue)
                            {
                                hardwareData.AppendLine($"GPU_LOAD:{sensor.Name}:{sensor.Value.Value:F1}%");
                            }
                        }
                    }
                }
                
                hardwareData.AppendLine("END");
                
                // Send data to COM port
                SendToComPort(hardwareData.ToString());

                Debug.WriteLine($"Hardware monitoring data collected and sent to {ComPort}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error monitoring hardware: {ex.Message}");
            }
        }
        
        private void SendToComPort(string data)
        {
            try
            {
                if (_serialPort != null && _serialPort.IsOpen)
                {
                    _serialPort.Write(data);
                    Debug.WriteLine($"Sent to {ComPort}: {data.Replace("\r\n", " | ")}");
                }
                else
                {
                    Debug.WriteLine($"Serial port {ComPort} is not open, cannot send data");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error sending data to COM port: {ex.Message}");
            }
        }
    }
    
    public class UpdateVisitor : IVisitor
    {
        public void VisitComputer(IComputer computer)
        {
            computer.Traverse(this);
        }
        
        public void VisitHardware(IHardware hardware)
        {
            hardware.Update();
            foreach (IHardware subHardware in hardware.SubHardware)
                subHardware.Accept(this);
        }
        
        public void VisitSensor(ISensor sensor) { }
        
        public void VisitParameter(IParameter parameter) { }
    }
}
