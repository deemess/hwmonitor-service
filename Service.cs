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
        
        public Thread Thread { get { return _workerThread; } }
        public Computer Computer { get { return _computer; } }

        // Configuration properties
        private string ComPort { get; set; }
        private int BaudRate { get; set; }
        private int DataBits { get; set; }
        private Parity Parity { get; set; }
        private StopBits StopBits { get; set; }
        private int MonitoringInterval { get; set; }

        public String CpuTxt { get; set; }
        public String CpuTxtEnd { get; set; }
        public String SensorIdCpu { get; set; }
        public String SensorIdCpuTemp { get; set; }
        public String MemTxt { get; set; }
        public String MemTxtEnd { get; set; }
        public String SensorIdMem { get; set; }
        public String SensorIdMemValue { get; set; }
        public String Gpu1Txt { get; set; }
        public String Gpu1TxtEnd { get; set; }
        public String SensorIdGpu1 { get; set; }
        public String SensorIdGpu1Temp { get; set; }
        public String Gpu2Txt { get; set; }
        public String Gpu2TxtEnd { get; set; }
        public String SensorIdGpu2 { get; set; }
        public String SensorIdGpu2Temp { get; set; }


        private Dictionary<String,String> sensorNameToId = new Dictionary<String,String>();
        private Dictionary<String, String> sensorIdToValue = new Dictionary<String, String>();
        private List<String> sensorIds = new List<String>();


        public HWMonService()
        {
            InitializeComponent();
            LoadConfiguration();
            InitComputerHardware();
        }

        private void LoadConfiguration()
        {
            try
            {
                ComPort = ConfigurationManager.AppSettings["ComPort"] ?? "COM1";
                BaudRate = int.Parse(ConfigurationManager.AppSettings["BaudRate"] ?? "115200");
                DataBits = int.Parse(ConfigurationManager.AppSettings["DataBits"] ?? "8");
                MonitoringInterval = int.Parse(ConfigurationManager.AppSettings["MonitoringInterval"] ?? "500");
                
                string parityStr = ConfigurationManager.AppSettings["Parity"] ?? "None";
                Parity = (Parity)Enum.Parse(typeof(Parity), parityStr, true);
                
                string stopBitsStr = ConfigurationManager.AppSettings["StopBits"] ?? "One";
                StopBits = (StopBits)Enum.Parse(typeof(StopBits), stopBitsStr, true);

                SensorIdCpu = ConfigurationManager.AppSettings["SensorIdCpu"];
                SensorIdCpuTemp = ConfigurationManager.AppSettings["SensorIdCpuTemp"];
                SensorIdMem = ConfigurationManager.AppSettings["SensorIdMem"];
                SensorIdMemValue = ConfigurationManager.AppSettings["SensorIdMemValue"];
                SensorIdGpu1 = ConfigurationManager.AppSettings["SensorIdGpu1"];
                SensorIdGpu1Temp = ConfigurationManager.AppSettings["SensorIdGpu1Temp"];
                SensorIdGpu2 = ConfigurationManager.AppSettings["SensorIdGpu2"];
                SensorIdGpu2Temp = ConfigurationManager.AppSettings["SensorIdGpu2Temp"];

                CpuTxt = ConfigurationManager.AppSettings["CpuTxt"];
                MemTxt = ConfigurationManager.AppSettings["MemTxt"];
                Gpu1Txt = ConfigurationManager.AppSettings["Gpu1Txt"];
                Gpu2Txt = ConfigurationManager.AppSettings["Gpu2Txt"];

                CpuTxtEnd = ConfigurationManager.AppSettings["CpuTxtEnd"];
                MemTxtEnd = ConfigurationManager.AppSettings["MemTxtEnd"];
                Gpu1TxtEnd = ConfigurationManager.AppSettings["Gpu1TxtEnd"];
                Gpu2TxtEnd = ConfigurationManager.AppSettings["Gpu2TxtEnd"];

                sensorNameToId.Add("SensorIdCpu", SensorIdCpu);
                sensorNameToId.Add("SensorIdCpuTemp", SensorIdCpuTemp);
                sensorNameToId.Add("SensorIdMem", SensorIdMem);
                sensorNameToId.Add("SensorIdMemValue", SensorIdMemValue);
                sensorNameToId.Add("SensorIdGpu1", SensorIdGpu1);
                sensorNameToId.Add("SensorIdGpu1Temp", SensorIdGpu1Temp);
                sensorNameToId.Add("SensorIdGpu2", SensorIdGpu2);
                sensorNameToId.Add("SensorIdGpu2Temp", SensorIdGpu2Temp);

                sensorIds.AddRange(sensorNameToId.Values);

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
                    Thread.Sleep(MonitoringInterval); // Sleep
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
                    Thread.Sleep(3000);
                }
            }
        }

        public void InitComputerHardware()
        {
            // Initialize hardware monitoring
            _computer = new Computer
            {
                IsCpuEnabled = true,
                IsGpuEnabled = true,
                IsMemoryEnabled = true,
                IsMotherboardEnabled = false,
                IsControllerEnabled = false,
                IsNetworkEnabled = false,
                IsStorageEnabled = false
            };

            _computer.Open();
            _computer.Accept(new UpdateVisitor());
        }

        public void UpdateComputerHardware()
        {
            // Update hardware sensors
            _computer.Accept(new UpdateVisitor());
        }

        private void walkHardwareSensors(IList<IHardware> hardlist, IList<String> searchSensors, IDictionary<String,String> result)
        {
            foreach (IHardware hardware in hardlist)
            {
                checkSensors(hardware, searchSensors, result);
                walkHardwareSensors(hardware.SubHardware, searchSensors, result);
            }
        }

        private void checkSensors(IHardware hardware, IList<String> searchSensors, IDictionary<String, String> result)
        {
            foreach (ISensor sensor in hardware.Sensors)
            {
                if (searchSensors.Contains(sensor.Identifier.ToString()))
                {
                    result.Add(sensor.Identifier.ToString(), $"{sensor.Value:000}");
                }
            }
        }

        private void MonitorHardware()
        {
            if (!_isRunning || _computer == null)
                return;
                
            try
            {
                UpdateComputerHardware();                
                sensorIdToValue.Clear();
                walkHardwareSensors(_computer.Hardware, sensorIds, sensorIdToValue);

                var hardwareData = new StringBuilder();

                hardwareData.AppendLine($"CPU:{sensorIdToValue[sensorNameToId["SensorIdCpu"]]}");
                hardwareData.AppendLine($"MEM:{sensorIdToValue[sensorNameToId["SensorIdMem"]]}");
                hardwareData.AppendLine($"GPU1:{sensorIdToValue[sensorNameToId["SensorIdGpu1"]]}");
                hardwareData.AppendLine($"GPU2:{sensorIdToValue[sensorNameToId["SensorIdGpu2"]]}");

                hardwareData.AppendLine($"CPUTXT:{CpuTxt}{sensorIdToValue[sensorNameToId["SensorIdCpuTemp"]].Substring(1)}{CpuTxtEnd}");
                hardwareData.AppendLine($"MEMTXT:{MemTxt}{sensorIdToValue[sensorNameToId["SensorIdMemValue"]].Substring(1)}{MemTxtEnd}");
                hardwareData.AppendLine($"GPU1TXT:{Gpu1Txt}{sensorIdToValue[sensorNameToId["SensorIdGpu1Temp"]].Substring(1)}{Gpu1TxtEnd}");
                hardwareData.AppendLine($"GPU2TXT:{Gpu2Txt}{sensorIdToValue[sensorNameToId["SensorIdGpu2Temp"]].Substring(1)}{Gpu2TxtEnd}");

                Console.WriteLine(hardwareData.ToString());

                // Send data to COM port
                SendToComPort(hardwareData.ToString());
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
