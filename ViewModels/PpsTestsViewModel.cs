using System;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Diagnostics;

namespace AGXCalibrationUI.ViewModels
{
    public class PpsTestsViewModel : INotifyPropertyChanged
    {
        private string? _selectedPpsModel;
        private string _voltage = string.Empty;
        private string _frequency = string.Empty;
        private string _duration = string.Empty;
        private string _initializationStatus = string.Empty;
        private string _testResults = string.Empty;
        private bool _isTestRunning;
        private bool _isLoading;
        private string _loadingMessage = string.Empty;
        private double _testProgress;
        private bool _hasTestResults;
        private readonly Dictionary<string, (double min, double max)> _voltageRanges;
        private readonly Dictionary<string, (double min, double max)> _frequencyRanges;

        public event PropertyChangedEventHandler? PropertyChanged;

        public PpsTestsViewModel()
        {
            // Initialize voltage ranges for each PPS model
            _voltageRanges = new Dictionary<string, (double min, double max)>
            {
                { "115 ACx UPC-1", (0, 300) },
                { "118 ACx UPC1", (0, 300) },
                { "360 AMx UPC32", (0, 400) },
                { "360 ASx UPC3", (0, 400) },
                { "3150 AFx", (0, 1000) }
            };

            // Initialize frequency ranges
            _frequencyRanges = new Dictionary<string, (double min, double max)>
            {
                { "115 ACx UPC-1", (45, 500) },
                { "118 ACx UPC1", (45, 500) },
                { "360 AMx UPC32", (45, 500) },
                { "360 ASx UPC3", (45, 500) },
                { "3150 AFx", (45, 1000) }
            };

            InitializeCommands();
        }

        #region Properties

        public string? SelectedPpsModel
        {
            get => _selectedPpsModel;
            set
            {
                if (_selectedPpsModel != value)
                {
                    _selectedPpsModel = value;
                    OnPropertyChanged(nameof(SelectedPpsModel));
                    OnPropertyChanged(nameof(VoltageRange));
                    OnPropertyChanged(nameof(FrequencyRange));
                    ValidateInputs();
                }
            }
        }

        public string Voltage
        {
            get => _voltage;
            set
            {
                if (_voltage != value)
                {
                    _voltage = value;
                    OnPropertyChanged(nameof(Voltage));
                    ValidateInputs();
                }
            }
        }

        public string Frequency
        {
            get => _frequency;
            set
            {
                if (_frequency != value)
                {
                    _frequency = value;
                    OnPropertyChanged(nameof(Frequency));
                    ValidateInputs();
                }
            }
        }

        public string Duration
        {
            get => _duration;
            set
            {
                if (_duration != value)
                {
                    _duration = value;
                    OnPropertyChanged(nameof(Duration));
                    ValidateInputs();
                }
            }
        }

        public string InitializationStatus
        {
            get => _initializationStatus;
            set
            {
                if (_initializationStatus != value)
                {
                    _initializationStatus = value;
                    OnPropertyChanged(nameof(InitializationStatus));
                }
            }
        }

        public string TestResults
        {
            get => _testResults;
            set
            {
                if (_testResults != value)
                {
                    _testResults = value;
                    OnPropertyChanged(nameof(TestResults));
                    HasTestResults = !string.IsNullOrWhiteSpace(value);
                }
            }
        }

        public bool IsTestRunning
        {
            get => _isTestRunning;
            set
            {
                if (_isTestRunning != value)
                {
                    _isTestRunning = value;
                    OnPropertyChanged(nameof(IsTestRunning));
                    OnPropertyChanged(nameof(CanStartTest));
                }
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                if (_isLoading != value)
                {
                    _isLoading = value;
                    OnPropertyChanged(nameof(IsLoading));
                }
            }
        }

        public string LoadingMessage
        {
            get => _loadingMessage;
            set
            {
                if (_loadingMessage != value)
                {
                    _loadingMessage = value;
                    OnPropertyChanged(nameof(LoadingMessage));
                }
            }
        }

        public double TestProgress
        {
            get => _testProgress;
            set
            {
                if (_testProgress != value)
                {
                    _testProgress = value;
                    OnPropertyChanged(nameof(TestProgress));
                }
            }
        }

        public bool HasTestResults
        {
            get => _hasTestResults;
            set
            {
                if (_hasTestResults != value)
                {
                    _hasTestResults = value;
                    OnPropertyChanged(nameof(HasTestResults));
                }
            }
        }

        public string VoltageRange
        {
            get
            {
                if (SelectedPpsModel != null && _voltageRanges.TryGetValue(SelectedPpsModel, out var range))
                {
                    return $"({range.min}-{range.max}V)";
                }
                return string.Empty;
            }
        }

        public string FrequencyRange
        {
            get
            {
                if (SelectedPpsModel != null && _frequencyRanges.TryGetValue(SelectedPpsModel, out var range))
                {
                    return $"({range.min}-{range.max}Hz)";
                }
                return string.Empty;
            }
        }

        public bool CanStartTest => !IsTestRunning && ValidateInputs();

        #endregion

        #region Commands

        public ICommand InitializePpsCommand { get; private set; }
        public ICommand StartTestCommand { get; private set; }
        public ICommand StopTestCommand { get; private set; }

        private void InitializeCommands()
        {
            InitializePpsCommand = new RelayCommand(async _ => await InitializePpsAsync());
            StartTestCommand = new RelayCommand(async _ => await StartTestAsync(), _ => CanStartTest);
            StopTestCommand = new RelayCommand(_ => StopTest(), _ => IsTestRunning);
        }

        #endregion

        #region Command Methods

        private async Task InitializePpsAsync()
        {
            if (string.IsNullOrEmpty(SelectedPpsModel))
            {
                InitializationStatus = "Please select a PPS model";
                return;
            }

            try
            {
                IsLoading = true;
                LoadingMessage = "Initializing PPS...";
                InitializationStatus = "Initializing...";

                // Simulate initialization (replace with actual initialization code)
                await Task.Delay(2000);

                InitializationStatus = "Initialized successfully";
                Debug.WriteLine($"PPS {SelectedPpsModel} initialized");
            }
            catch (Exception ex)
            {
                InitializationStatus = $"Initialization failed: {ex.Message}";
                Debug.WriteLine($"PPS initialization failed: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                LoadingMessage = string.Empty;
            }
        }

        private async Task StartTestAsync()
        {
            if (!ValidateInputs())
            {
                TestResults = "Invalid test parameters. Please check inputs.";
                return;
            }

            try
            {
                IsTestRunning = true;
                TestProgress = 0;
                TestResults = "Test started...\n";

                // Parse validated inputs
                double voltage = double.Parse(Voltage);
                double frequency = double.Parse(Frequency);
                int duration = int.Parse(Duration);

                // Simulate test progress
                for (int i = 0; i <= duration && IsTestRunning; i++)
                {
                    TestProgress = (i * 100.0) / duration;
                    TestResults += $"Time: {i}s - Voltage: {voltage}V, Frequency: {frequency}Hz\n";
                    await Task.Delay(1000);
                }

                if (IsTestRunning) // If not stopped
                {
                    TestResults += "Test completed successfully";
                    TestProgress = 100;
                }
            }
            catch (Exception ex)
            {
                TestResults += $"\nTest failed: {ex.Message}";
                Debug.WriteLine($"Test failed: {ex.Message}");
            }
            finally
            {
                IsTestRunning = false;
            }
        }

        private void StopTest()
        {
            if (IsTestRunning)
            {
                IsTestRunning = false;
                TestResults += "\nTest stopped by user";
            }
        }

        #endregion

        #region Validation

        private bool ValidateInputs()
        {
            if (string.IsNullOrEmpty(SelectedPpsModel))
                return false;

            bool isValid = true;

            // Validate voltage
            if (!string.IsNullOrEmpty(Voltage) && 
                double.TryParse(Voltage, out double voltageValue) &&
                _voltageRanges.TryGetValue(SelectedPpsModel, out var voltageRange))
            {
                isValid &= voltageValue >= voltageRange.min && voltageValue <= voltageRange.max;
            }
            else
            {
                isValid = false;
            }

            // Validate frequency
            if (!string.IsNullOrEmpty(Frequency) && 
                double.TryParse(Frequency, out double frequencyValue) &&
                _frequencyRanges.TryGetValue(SelectedPpsModel, out var frequencyRange))
            {
                isValid &= frequencyValue >= frequencyRange.min && frequencyValue <= frequencyRange.max;
            }
            else
            {
                isValid = false;
            }

            // Validate duration
            if (!string.IsNullOrEmpty(Duration) && 
                int.TryParse(Duration, out int durationValue))
            {
                isValid &= durationValue >= 1 && durationValue <= 3600;
            }
            else
            {
                isValid = false;
            }

            return isValid;
        }

        #endregion

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
