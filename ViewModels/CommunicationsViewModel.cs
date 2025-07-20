using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Controls;
using AGXCalibrationUI.Services;
using AGXCalibrationUI.Models;
using AGXCalibrationUI.Views;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace AGXCalibrationUI.ViewModels
{
    public class CommunicationsViewModel : INotifyPropertyChanged
    {
        private readonly VisaCommunicationService _visaService;
        private readonly CalibrationModel _calibrationModel;
        private string? _selectedResource;
        private string _commandInput = string.Empty;
        private string _responseOutput = string.Empty;
        private bool _isConnected;
        private string _connectionDetails = string.Empty;
        private string _calibrationStatus = string.Empty;
        private bool _isManualEntry;
        private bool _isReadback;
        private ObservableCollection<string> _availableResources;
        private CommunicationType _currentInterfaceType;
        private Dictionary<string, string> _interfaceDetails;
        private UserControl? _currentView;
        private bool _isConnectViewVisible = true;
        private string _connectionStatus = "Not Connected";
        private bool _isRefreshing;
        private bool _isInitializing = true;

        // Commands
        private ICommand? _connectCommand;
        private ICommand? _disconnectCommand;
        private ICommand? _sendCommand;
        private ICommand? _refreshResourcesCommand;
        private ICommand? _startCalibrationCommand;
        private ICommand? _loadExcelCommand;
        private ICommand? _navigateCommand;
        private ICommand? _configureInterfaceCommand;

        public event PropertyChangedEventHandler? PropertyChanged;

        public CommunicationsViewModel()
        {
            Debug.WriteLine("Initializing CommunicationsViewModel");
            
            try
            {
                _visaService = new VisaCommunicationService();
                Debug.WriteLine("VisaCommunicationService created successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error creating VisaCommunicationService: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                ResponseOutput = $"Failed to initialize VISA service: {ex.Message}";
                throw;
            }
            
            _calibrationModel = new CalibrationModel();
            _availableResources = new ObservableCollection<string>();
            _interfaceDetails = new Dictionary<string, string>();
            InitializeCommands();
            
            // Start initial resource refresh
            _ = InitialResourceRefreshAsync();
        }

        private async Task InitialResourceRefreshAsync()
        {
            try
            {
                ConnectionStatus = "Initializing...";
                await RefreshResourcesAsync();
                
                if (_availableResources.Count == 0)
                {
                    ResponseOutput = "No VISA resources found. Please check:\n" +
                                   "1. Device is powered on and connected\n" +
                                   "2. USB/GPIB/LAN cable is properly connected\n" +
                                   "3. Required drivers are installed\n" +
                                   "4. Run application as Administrator";
                    ConnectionStatus = "No Devices Found";
                }
                else
                {
                    ResponseOutput = $"Found {_availableResources.Count} VISA resources";
                    ConnectionStatus = "Ready";
                }
            }
            catch (Exception ex)
            {
                ResponseOutput = $"Failed to detect resources:\n{ex.Message}\n\nPlease check:\n" +
                               "1. NI-VISA is properly installed\n" +
                               "2. Required drivers are installed\n" +
                               "3. Run application as Administrator";
                ConnectionStatus = "Initialization Failed";
                Debug.WriteLine($"Initial resource refresh failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                _isInitializing = false;
            }
        }

        public ObservableCollection<string> AvailableResources
        {
            get => _availableResources;
            set
            {
                _availableResources = value;
                OnPropertyChanged(nameof(AvailableResources));
            }
        }

        public Dictionary<string, string> InterfaceDetails
        {
            get => _interfaceDetails;
            set
            {
                _interfaceDetails = value;
                OnPropertyChanged(nameof(InterfaceDetails));
            }
        }

        public CommunicationType CurrentInterfaceType
        {
            get => _currentInterfaceType;
            set
            {
                _currentInterfaceType = value;
                OnPropertyChanged(nameof(CurrentInterfaceType));
                OnPropertyChanged(nameof(IsLanInterface));
                OnPropertyChanged(nameof(IsGpibInterface));
            }
        }

        public bool IsLanInterface => CurrentInterfaceType == CommunicationType.LAN;
        public bool IsGpibInterface => CurrentInterfaceType == CommunicationType.GPIB;

        public string? SelectedResource
        {
            get => _selectedResource;
            set
            {
                _selectedResource = value;
                OnPropertyChanged(nameof(SelectedResource));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string CommandInput
        {
            get => _commandInput;
            set
            {
                _commandInput = value;
                OnPropertyChanged(nameof(CommandInput));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string ResponseOutput
        {
            get => _responseOutput;
            set
            {
                _responseOutput = value;
                OnPropertyChanged(nameof(ResponseOutput));
            }
        }

        public bool IsConnected
        {
            get => _isConnected;
            set
            {
                _isConnected = value;
                ConnectionStatus = value ? "Connected" : "Not Connected";
                OnPropertyChanged(nameof(IsConnected));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string ConnectionStatus
        {
            get => _connectionStatus;
            set
            {
                _connectionStatus = value;
                OnPropertyChanged(nameof(ConnectionStatus));
            }
        }

        public string ConnectionDetails
        {
            get => _connectionDetails;
            set
            {
                _connectionDetails = value;
                OnPropertyChanged(nameof(ConnectionDetails));
            }
        }

        public string CalibrationStatus
        {
            get => _calibrationStatus;
            set
            {
                _calibrationStatus = value;
                OnPropertyChanged(nameof(CalibrationStatus));
            }
        }

        public bool IsManualEntry
        {
            get => _isManualEntry;
            set
            {
                _isManualEntry = value;
                _isReadback = !value;
                OnPropertyChanged(nameof(IsManualEntry));
                OnPropertyChanged(nameof(IsReadback));
            }
        }

        public bool IsReadback
        {
            get => _isReadback;
            set
            {
                _isReadback = value;
                _isManualEntry = !value;
                OnPropertyChanged(nameof(IsReadback));
                OnPropertyChanged(nameof(IsManualEntry));
            }
        }

        public UserControl? CurrentView
        {
            get => _currentView;
            set
            {
                _currentView = value;
                OnPropertyChanged(nameof(CurrentView));
            }
        }

        public bool IsConnectViewVisible
        {
            get => _isConnectViewVisible;
            set
            {
                _isConnectViewVisible = value;
                OnPropertyChanged(nameof(IsConnectViewVisible));
            }
        }

        public bool IsRefreshing
        {
            get => _isRefreshing;
            set
            {
                _isRefreshing = value;
                OnPropertyChanged(nameof(IsRefreshing));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public ICommand ConnectCommand => _connectCommand ??= new RelayCommand(async _ => await ConnectAsync(), _ => !IsConnected && !string.IsNullOrEmpty(SelectedResource) && !IsRefreshing && !_isInitializing);
        public ICommand DisconnectCommand => _disconnectCommand ??= new RelayCommand(async _ => await DisconnectAsync(), _ => IsConnected && !IsRefreshing);
        public ICommand SendCommand => _sendCommand ??= new RelayCommand(async _ => await SendCommandAsync(), _ => IsConnected && !string.IsNullOrEmpty(CommandInput) && !IsRefreshing);
        public ICommand RefreshResourcesCommand => _refreshResourcesCommand ??= new RelayCommand(async _ => await RefreshResourcesAsync(), _ => !IsRefreshing && !_isInitializing);
        public ICommand StartCalibrationCommand => _startCalibrationCommand ??= new RelayCommand(async _ => await StartCalibrationAsync(), _ => IsConnected && !IsRefreshing);
        public ICommand LoadExcelCommand => _loadExcelCommand ??= new RelayCommand(_ => LoadExcelFile());
        public ICommand NavigateCommand => _navigateCommand ??= new RelayCommand(Navigate);
        public ICommand ConfigureInterfaceCommand => _configureInterfaceCommand ??= new RelayCommand(async param => await ConfigureInterfaceAsync(param), _ => IsConnected && !IsRefreshing);

        private void InitializeCommands()
        {
            // Commands are initialized using the null-coalescing operator in the properties
        }

        private async Task ConnectAsync()
        {
            if (string.IsNullOrEmpty(SelectedResource))
            {
                ResponseOutput = "No resource selected";
                return;
            }

            try
            {
                Debug.WriteLine($"Attempting to connect to resource: {SelectedResource}");
                ConnectionStatus = "Connecting...";
                IsConnected = await _visaService.ConnectAsync(SelectedResource);
                
                if (IsConnected)
                {
                    var idResponse = await _visaService.GetDeviceId();
                    Debug.WriteLine($"Device identified: {idResponse}");
                    
                    CurrentInterfaceType = _visaService.CurrentInterface;
                    Debug.WriteLine($"Interface type: {CurrentInterfaceType}");
                    
                    InterfaceDetails = await _visaService.GetInterfaceDetails();
                    Debug.WriteLine("Interface details retrieved");
                    
                    var details = new List<string>
                    {
                        $"Device ID: {idResponse}",
                        $"Interface: {CurrentInterfaceType}"
                    };
                    
                    foreach (var detail in InterfaceDetails)
                    {
                        details.Add($"{detail.Key}: {detail.Value}");
                        Debug.WriteLine($"Interface detail - {detail.Key}: {detail.Value}");
                    }
                    
                    ConnectionDetails = string.Join("\n", details);
                    CalibrationStatus = "Ready for calibration";
                    ResponseOutput = "Connection established successfully";
                    Debug.WriteLine("Connection established successfully");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Connection failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                ResponseOutput = $"Connection failed: {ex.Message}";
                IsConnected = false;
                ConnectionStatus = "Connection Failed";
            }
        }

        private async Task DisconnectAsync()
        {
            try
            {
                Debug.WriteLine("Attempting to disconnect");
                ConnectionStatus = "Disconnecting...";
                await _visaService.DisconnectAsync();
                IsConnected = false;
                ConnectionDetails = string.Empty;
                CalibrationStatus = string.Empty;
                CurrentInterfaceType = CommunicationType.Unknown;
                InterfaceDetails.Clear();
                ResponseOutput = "Disconnected successfully";
                Debug.WriteLine("Disconnected successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Disconnect failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                ResponseOutput = $"Disconnect failed: {ex.Message}";
            }
            finally
            {
                ConnectionStatus = "Not Connected";
            }
        }

        private async Task SendCommandAsync()
        {
            try
            {
                Debug.WriteLine($"Sending command: {CommandInput}");
                if (CommandInput.EndsWith("?"))
                {
                    var response = await _visaService.QueryAsync(CommandInput);
                    ResponseOutput = response;
                    Debug.WriteLine($"Query response: {response}");
                }
                else
                {
                    await _visaService.WriteAsync(CommandInput);
                    ResponseOutput = "Command sent successfully";
                    Debug.WriteLine("Command sent successfully");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Command failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                ResponseOutput = $"Command failed: {ex.Message}";
            }
        }

        private async Task RefreshResourcesAsync()
        {
            if (IsRefreshing) return;

            try
            {
                IsRefreshing = true;
                ConnectionStatus = "Searching for devices...";
                Debug.WriteLine("Starting resource refresh");
                
                var resources = await _visaService.GetAvailableResourcesAsync();
                AvailableResources.Clear();
                foreach (var resource in resources)
                {
                    Debug.WriteLine($"Found resource: {resource}");
                    AvailableResources.Add(resource);
                }
                
                if (resources.Count == 0)
                {
                    ResponseOutput = "No VISA resources found. Please check:\n" +
                                   "1. Device is powered on and connected\n" +
                                   "2. USB/GPIB/LAN cable is properly connected\n" +
                                   "3. Required drivers are installed\n" +
                                   "4. Run application as Administrator";
                    Debug.WriteLine("No resources found");
                }
                else
                {
                    ResponseOutput = $"Found {resources.Count} VISA resources";
                    Debug.WriteLine($"Resource refresh completed. Found {resources.Count} resources");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Resource refresh failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                ResponseOutput = $"Failed to refresh resources:\n{ex.Message}\n\nPlease check:\n" +
                               "1. NI-VISA is properly installed\n" +
                               "2. Required drivers are installed\n" +
                               "3. Run application as Administrator";
            }
            finally
            {
                IsRefreshing = false;
                ConnectionStatus = IsConnected ? "Connected" : "Not Connected";
            }
        }

        private async Task ConfigureInterfaceAsync(object? parameter)
        {
            if (parameter is string config)
            {
                try
                {
                    Debug.WriteLine($"Configuring interface: {config}");
                    switch (config)
                    {
                        case "EnableDHCP":
                            await _visaService.EnableDhcp(true);
                            break;
                        case "DisableDHCP":
                            await _visaService.EnableDhcp(false);
                            break;
                    }

                    InterfaceDetails = await _visaService.GetInterfaceDetails();
                    ResponseOutput = "Interface configuration updated successfully";
                    Debug.WriteLine("Interface configuration completed successfully");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Interface configuration failed: {ex.Message}");
                    Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    ResponseOutput = $"Failed to configure interface: {ex.Message}";
                }
            }
        }

        private async Task StartCalibrationAsync()
        {
            try
            {
                Debug.WriteLine("Starting calibration");
                CalibrationStatus = "Calibration in progress...";
                await Task.Run(() => _calibrationModel.StartCalibration());
                CalibrationStatus = "Calibration completed successfully";
                Debug.WriteLine("Calibration completed successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Calibration failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                CalibrationStatus = $"Calibration failed: {ex.Message}";
            }
        }

        private void LoadExcelFile()
        {
            try
            {
                Debug.WriteLine("Opening file dialog for Excel file selection");
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel Files|*.xlsx;*.xls",
                    Title = "Select Excel File"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    Debug.WriteLine($"Selected Excel file: {openFileDialog.FileName}");
                    ResponseOutput = $"Selected Excel file: {openFileDialog.FileName}";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Excel file selection failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                ResponseOutput = $"Failed to load Excel file: {ex.Message}";
            }
        }

        private void Navigate(object? parameter)
        {
            try
            {
                if (parameter is string view)
                {
                    Debug.WriteLine($"Navigating to view: {view}");
                    switch (view.ToLower())
                    {
                        case "connect":
                            IsConnectViewVisible = true;
                            CurrentView = null;
                            break;
                        case "pattests":
                            IsConnectViewVisible = false;
                            CurrentView = new PatTests();
                            break;
                        case "ppstests":
                            IsConnectViewVisible = false;
                            CurrentView = new PpsTests();
                            break;
                        case "meters":
                            IsConnectViewVisible = false;
                            CurrentView = new Meters();
                            break;
                        case "automatedtests":
                            IsConnectViewVisible = false;
                            CurrentView = new AutomatedTests();
                            break;
                        default:
                            IsConnectViewVisible = true;
                            CurrentView = null;
                            break;
                    }
                    Debug.WriteLine("Navigation completed successfully");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Navigation failed: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                ResponseOutput = $"Navigation failed: {ex.Message}";
            }
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action<object?> _execute;
        private readonly Predicate<object?>? _canExecute;

        public RelayCommand(Action<object?> execute, Predicate<object?>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public RelayCommand(Func<object?, Task> execute, Predicate<object?>? canExecute = null)
        {
            _execute = async parameter =>
            {
                await execute(parameter);
            };
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter)
        {
            return _canExecute == null || _canExecute(parameter);
        }

        public void Execute(object? parameter)
        {
            _execute(parameter);
        }
    }
}
