using System;
using System.Threading.Tasks;
using NationalInstruments.Visa;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Win32;

namespace AGXCalibrationUI.Services
{
    public enum CommunicationType
    {
        USB,
        LAN,
        GPIB,
        Unknown
    }

    public class VisaCommunicationService : IDisposable
    {
        private MessageBasedSession? session;
        private ResourceManager? resourceManager;
        private bool isConnected;
        private readonly int defaultTimeout = 5000; // 5 seconds timeout
        private const string TerminationCharacter = "\n";
        private CommunicationType currentInterfaceType = CommunicationType.Unknown;
        private const int MaxRetries = 3;
        private const int RetryDelayMs = 1000;

        // Default port for LAN communication as per manual
        private const int DefaultLanPort = 5025;
        // Default GPIB address as per manual
        private const int DefaultGpibAddress = 1;
        // USB-LAN emulation IP as per manual
        private const string UsbLanEmulationIp = "192.168.123.1";

        public bool IsConnected => isConnected;
        public CommunicationType CurrentInterface => currentInterfaceType;

        private string GetDriverInstallationMessage()
        {
            var message = new StringBuilder();
            message.AppendLine("AGX USB drivers not found. Please follow these steps to install:");
            message.AppendLine("1. Connect the AGX USB port to the PC with a USB cable");
            message.AppendLine("2. Power on the AGX device");
            message.AppendLine("3. Wait for Windows to detect the device");
            message.AppendLine("4. If drivers are not installed automatically:");
            message.AppendLine("   a. Check if a new USB drive appears in File Explorer");
            message.AppendLine("   b. Navigate to the 'drivers/Windows' directory on that drive");
            message.AppendLine("   c. Run Driver_Installer.exe");
            message.AppendLine("5. After installation, verify in Device Manager:");
            message.AppendLine("   - Look for 'AGX Virtual COM Port' under Ports (COM & LPT)");
            message.AppendLine("   - Look for 'AGX Network Adapter' under Network adapters");
            return message.ToString();
        }

        private bool VerifyDriverInstallation()
        {
            try
            {
                Debug.WriteLine("Verifying driver installation...");

                // Check for required VISA DLLs
                var driversPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "drivers");
                var requiredDlls = new[] {
                    "Ivi.Visa.dll",
                    "NationalInstruments.Visa.dll",
                    "Microsoft.VisualStudio.Tools.Applications.Runtime.dll"
                };

                var missingDlls = requiredDlls.Where(dll => !File.Exists(Path.Combine(driversPath, dll))).ToList();
                if (missingDlls.Any())
                {
                    throw new Exception($"Missing required VISA DLLs: {string.Join(", ", missingDlls)}");
                }

                // Check for USB drivers in Device Manager
                bool hasUsbDrivers = false;
                using (var networkKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}"))
                using (var comKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Class\{4D36E978-E325-11CE-BFC1-08002BE10318}"))
                {
                    if (networkKey != null && comKey != null)
                    {
                        // Look for AGX USB drivers
                        var hasNetworkDriver = networkKey.GetSubKeyNames()
                            .Any(name => {
                                using var subKey = networkKey.OpenSubKey(name);
                                var driverDesc = subKey?.GetValue("DriverDesc") as string;
                                return driverDesc?.Contains("AGX", StringComparison.OrdinalIgnoreCase) ?? false;
                            });

                        var hasComDriver = comKey.GetSubKeyNames()
                            .Any(name => {
                                using var subKey = comKey.OpenSubKey(name);
                                var driverDesc = subKey?.GetValue("DriverDesc") as string;
                                return driverDesc?.Contains("AGX", StringComparison.OrdinalIgnoreCase) ?? false;
                            });

                        hasUsbDrivers = hasNetworkDriver && hasComDriver;
                        Debug.WriteLine($"USB drivers found: Network={hasNetworkDriver}, COM={hasComDriver}");

                        if (!hasUsbDrivers)
                        {
                            throw new Exception(GetDriverInstallationMessage());
                        }
                    }
                    else
                    {
                        throw new Exception("Unable to access Device Manager. Please run the application as Administrator.");
                    }
                }

                Debug.WriteLine("Driver verification completed successfully");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Driver verification failed: {ex.Message}");
                throw;
            }
        }

        private async Task<bool> DetectActiveInterface()
        {
            try
            {
                // Check USB first
                try
                {
                    var usbResponse = await QueryAsync("SYSTem:COMMunicate:USB:VIRTualport:ENABle?");
                    if (usbResponse == "1")
                    {
                        currentInterfaceType = CommunicationType.USB;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"USB detection failed: {ex.Message}");
                }

                // Check LAN
                try
                {
                    var lanResponse = await QueryAsync("SYSTem:COMMunicate:LAN:STATus?");
                    if (!string.IsNullOrEmpty(lanResponse))
                    {
                        currentInterfaceType = CommunicationType.LAN;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"LAN detection failed: {ex.Message}");
                }

                // Check GPIB
                try
                {
                    var gpibResponse = await QueryAsync("SYSTem:COMMunicate:GPIB:ENABle?");
                    if (gpibResponse == "1")
                    {
                        currentInterfaceType = CommunicationType.GPIB;
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"GPIB detection failed: {ex.Message}");
                }

                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Interface detection failed: {ex.Message}");
                return false;
            }
        }

        private async Task EnableAllInterfaces()
        {
            try
            {
                // Enable all interfaces as per AGX_Connection_Details.txt
                await WriteAsync("SYSTem:COMMunicate:USB:VIRTualport:ENABle 1");
                await WriteAsync("SYSTem:COMMunicate:LAN:ENABle 1"); // Fixed command
                await WriteAsync("SYSTem:COMMunicate:GPIB:ENABle 1");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to enable interfaces: {ex.Message}");
                throw;
            }
        }

        public async Task<bool> ConnectAsync(string resourceName)
        {
            try
            {
                if (isConnected)
                {
                    await DisconnectAsync();
                }

                Debug.WriteLine($"Attempting to connect to resource: {resourceName}");
                
                // Verify driver installation first
                if (!VerifyDriverInstallation())
                {
                    throw new Exception("Required drivers not properly installed");
                }

                // Create new resource manager if not exists
                if (resourceManager == null)
                {
                    for (int i = 0; i < MaxRetries; i++)
                    {
                        try
                        {
                            resourceManager = new ResourceManager();
                            Debug.WriteLine("ResourceManager created successfully");
                            break;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Attempt {i + 1} to create ResourceManager failed: {ex.Message}");
                            if (i < MaxRetries - 1)
                            {
                                await Task.Delay(RetryDelayMs);
                            }
                            else throw;
                        }
                    }
                }
                
                try
                {
                    Debug.WriteLine("Attempting to open session...");
                    session = (MessageBasedSession)await Task.Run(() => resourceManager.Open(resourceName));
                    session.TimeoutMilliseconds = defaultTimeout;
                    Debug.WriteLine("Session opened successfully");

                    // Enable all interfaces first
                    await EnableAllInterfaces();
                    
                    // Detect which interface is actually active
                    if (!await DetectActiveInterface())
                    {
                        throw new Exception("No active interface detected");
                    }

                    Debug.WriteLine($"Detected active interface: {currentInterfaceType}");
                    
                    // Initialize IEEE-488.2 status reporting system
                    await InitializeIEEE4882Async();
                    
                    // Configure the detected interface
                    await ConfigureInterface(currentInterfaceType);

                    // Verify connection with device identification
                    var idResponse = await QueryAsync("*IDN?");
                    if (string.IsNullOrEmpty(idResponse))
                    {
                        throw new Exception("Device identification failed");
                    }

                    isConnected = true;
                    Debug.WriteLine("Connection established successfully");
                    return true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to open session: {ex.Message}");
                    Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    if (resourceManager != null)
                    {
                        resourceManager.Dispose();
                        resourceManager = null;
                    }
                    throw;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Connection failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw new Exception($"Failed to connect to instrument: {ex.Message}", ex);
            }
        }

        private async Task InitializeIEEE4882Async()
        {
            try
            {
                Debug.WriteLine("Initializing IEEE-488.2 status reporting system...");
                
                // Clear device
                await WriteAsync("*CLS");
                
                // Reset to default settings
                await WriteAsync("*RST");
                
                // Enable Operation Complete bit in Standard Event Status Register
                await WriteAsync("*ESE 1");
                
                // Enable Status Byte Register bits for Standard Event Status and Message Available
                await WriteAsync("*SRE 48");
                
                // Get device ID to verify connection
                string idResponse = await QueryAsync("*IDN?");
                if (string.IsNullOrEmpty(idResponse))
                {
                    throw new Exception("Device identification failed");
                }
                Debug.WriteLine($"Device identified: {idResponse}");

                // Perform self-test
                string tstResponse = await QueryAsync("*TST?");
                if (tstResponse.Trim() != "0")
                {
                    throw new Exception("Device self-test failed");
                }
                Debug.WriteLine("Self-test passed");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"IEEE-488.2 initialization failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw new Exception($"IEEE-488.2 initialization failed: {ex.Message}");
            }
        }

        private async Task ConfigureInterface(CommunicationType interfaceType)
        {
            Debug.WriteLine($"Configuring interface: {interfaceType}");
            switch (interfaceType)
            {
                case CommunicationType.USB:
                    await WriteAsync("SYSTem:COMMunicate:USB:VIRTualport:ENABle 1");
                    break;

                case CommunicationType.LAN:
                    // Enable LAN interface
                    await WriteAsync("SYSTem:COMMunicate:LAN:ENABle 1");
                    // Configure default port
                    await WriteAsync($"SYSTem:COMMunicate:LAN:PORT {DefaultLanPort}");
                    // Enable DHCP by default
                    await WriteAsync("SYSTem:COMMunicate:LAN:DHCP:ENABle 1");
                    await WriteAsync("SYSTem:COMMunicate:LAN:DHCP:RENEW");
                    break;

                case CommunicationType.GPIB:
                    await WriteAsync("SYSTem:COMMunicate:GPIB:ENABle 1");
                    await WriteAsync($"SYSTem:COMMunicate:GPIB:ADDress {DefaultGpibAddress}");
                    break;
            }
            Debug.WriteLine("Interface configuration completed");
        }

        public async Task<Dictionary<string, string>> GetInterfaceDetails()
        {
            var details = new Dictionary<string, string>();

            try
            {
                // Get common IEEE-488.2 status
                var esr = await QueryAsync("*ESR?");
                var stb = await QueryAsync("*STB?");
                details.Add("Event Status Register", esr);
                details.Add("Status Byte", stb);

                switch (currentInterfaceType)
                {
                    case CommunicationType.USB:
                        var usbStatus = await QueryAsync("SYSTem:COMMunicate:USB:VIRTualport:ENABle?");
                        details.Add("USB Virtual Port", usbStatus == "1" ? "Enabled" : "Disabled");
                        break;

                    case CommunicationType.LAN:
                        var lanStatus = await QueryAsync("SYSTem:COMMunicate:LAN:STATus?");
                        details.Add("LAN Status", lanStatus);
                        var dhcpStatus = await QueryAsync("SYSTem:COMMunicate:LAN:DHCP:ENABle?");
                        details.Add("DHCP", dhcpStatus == "1" ? "Enabled" : "Disabled");
                        var port = await QueryAsync("SYSTem:COMMunicate:LAN:PORT?");
                        details.Add("Port", port);
                        break;

                    case CommunicationType.GPIB:
                        var gpibAddress = await QueryAsync("SYSTem:COMMunicate:GPIB:ADDress?");
                        details.Add("GPIB Address", gpibAddress);
                        break;
                }
            }
            catch (Exception ex)
            {
                details.Add("Error", ex.Message);
                Debug.WriteLine($"Error getting interface details: {ex.Message}");
            }

            return details;
        }

        public async Task DisconnectAsync()
        {
            if (session != null)
            {
                try
                {
                    // Return to local control
                    await WriteAsync("&GTL");
                }
                catch (Exception ex) 
                {
                    Debug.WriteLine($"Warning during disconnect: {ex.Message}");
                }
                finally
                {
                    await Task.Run(() =>
                    {
                        session.Dispose();
                        session = null;
                    });
                    
                    if (resourceManager != null)
                    {
                        resourceManager.Dispose();
                        resourceManager = null;
                    }
                    
                    isConnected = false;
                    currentInterfaceType = CommunicationType.Unknown;
                    Debug.WriteLine("Disconnected successfully");
                }
            }
        }

        public async Task WriteAsync(string command)
        {
            if (!isConnected || session == null)
            {
                throw new Exception("Not connected to any instrument");
            }

            try
            {
                Debug.WriteLine($"Writing command: {command}");
                var message = command + TerminationCharacter;
                await Task.Run(() => session.RawIO.Write(message));
                
                // Wait for operation complete if it's an IEEE-488.2 command
                if (command.StartsWith("*"))
                {
                    await WaitForOperationComplete();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to write command: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw new Exception($"Failed to write command: {ex.Message}");
            }
        }

        public async Task<string> QueryAsync(string query)
        {
            if (!isConnected || session == null)
            {
                throw new Exception("Not connected to any instrument");
            }

            try
            {
                Debug.WriteLine($"Querying: {query}");
                await WriteAsync(query);
                
                var response = await Task.Run(() =>
                {
                    var result = session.RawIO.ReadString();
                    return result.TrimEnd();
                });
                Debug.WriteLine($"Response received: {response}");
                return response;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to query: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                throw new Exception($"Failed to query: {ex.Message}");
            }
        }

        public async Task<List<string>> GetAvailableResourcesAsync()
        {
            List<string> resources = new List<string>();
            Exception? lastException = null;
            
            Debug.WriteLine("Searching for available resources...");

            // Verify driver installation first
            try
            {
                if (!VerifyDriverInstallation())
                {
                    throw new Exception("Required drivers not properly installed");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Driver verification failed: {ex.Message}");
                throw;
            }
            
            // Try to initialize ResourceManager if not already initialized
            if (resourceManager == null)
            {
                for (int i = 0; i < MaxRetries; i++)
                {
                    try
                    {
                        resourceManager = new ResourceManager();
                        Debug.WriteLine("ResourceManager created successfully");
                        break;
                    }
                    catch (Exception ex)
                    {
                        lastException = ex;
                        Debug.WriteLine($"Attempt {i + 1} to create ResourceManager failed: {ex.Message}");
                        if (i < MaxRetries - 1)
                        {
                            await Task.Delay(RetryDelayMs);
                        }
                    }
                }

                if (resourceManager == null)
                {
                    throw new Exception($"Failed to initialize VISA system after {MaxRetries} attempts", lastException);
                }
            }

            // Search for resources with retry mechanism
            for (int i = 0; i < MaxRetries; i++)
            {
                try
                {
                    var foundResources = await Task.Run(() => resourceManager.Find("?*"));
                    Debug.WriteLine($"Found {foundResources.Count()} resources");

                    // Validate each resource
                    foreach (var resource in foundResources)
                    {
                        try
                        {
                            // Try to open a temporary session to verify the resource
                            using (var tempSession = (MessageBasedSession)resourceManager.Open(resource))
                            {
                                // Resource is valid
                                resources.Add(resource);
                                Debug.WriteLine($"Validated resource: {resource}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Resource validation failed for {resource}: {ex.Message}");
                        }
                    }

                    if (resources.Count > 0)
                    {
                        Debug.WriteLine($"Successfully found and validated {resources.Count} resources");
                        return resources;
                    }

                    if (i < MaxRetries - 1)
                    {
                        Debug.WriteLine($"No valid resources found, retrying... (Attempt {i + 1})");
                        await Task.Delay(RetryDelayMs);
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    Debug.WriteLine($"Attempt {i + 1} to find resources failed: {ex.Message}");
                    if (i < MaxRetries - 1)
                    {
                        await Task.Delay(RetryDelayMs);
                    }
                }
            }

            if (resources.Count == 0)
            {
                Debug.WriteLine("No resources found or all resources failed validation");
                if (lastException != null)
                {
                    throw new Exception("Failed to find any valid resources", lastException);
                }
            }

            return resources;
        }

        private async Task WaitForOperationComplete()
        {
            try
            {
                await QueryAsync("*OPC?");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Warning during operation complete check: {ex.Message}");
            }
        }

        public async Task EnableDhcp(bool enable)
        {
            if (currentInterfaceType == CommunicationType.LAN)
            {
                await WriteAsync($"SYSTem:COMMunicate:LAN:DHCP:ENABle {(enable ? 1 : 0)}");
                if (enable)
                {
                    await WriteAsync("SYSTem:COMMunicate:LAN:DHCP:RENEW");
                }
            }
        }

        public async Task<string> GetDeviceId() => await QueryAsync("*IDN?");

        public void Dispose()
        {
            DisconnectAsync().Wait();
        }
    }
}
