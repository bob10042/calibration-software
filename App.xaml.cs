using System;
using System.Windows;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Reflection;

namespace AGXCalibrationUI
{
    public partial class App : Application
    {
        private static Mutex? _mutex;

        public App()
        {
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        private Assembly? CurrentDomain_AssemblyResolve(object? sender, ResolveEventArgs args)
        {
            // Get the assembly name
            var assemblyName = new AssemblyName(args.Name).Name;
            if (string.IsNullOrEmpty(assemblyName)) return null;

            // First check the executable directory
            string assemblyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, assemblyName + ".dll");
            if (File.Exists(assemblyPath))
            {
                return Assembly.LoadFrom(assemblyPath);
            }

            // Then check the drivers subdirectory
            assemblyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "drivers", assemblyName + ".dll");
            if (File.Exists(assemblyPath))
            {
                return Assembly.LoadFrom(assemblyPath);
            }

            return null;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show($"An error occurred: {e.Exception.Message}\n\nStack trace:\n{e.Exception.StackTrace}", 
                          "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception exception)
            {
                MessageBox.Show($"A fatal error occurred: {exception.Message}\n\nStack trace:\n{exception.StackTrace}",
                              "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            const string appName = "AGXCalibrationUI";
            bool createdNew;

            _mutex = new Mutex(true, appName, out createdNew);

            if (!createdNew)
            {
                MessageBox.Show("Application is already running!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                Application.Current.Shutdown();
                return;
            }

            base.OnStartup(e);
            Debug.WriteLine("Application starting...");
            try
            {
                // Log the current directory and available DLLs for debugging
                Debug.WriteLine($"Base Directory: {AppDomain.CurrentDomain.BaseDirectory}");
                foreach (var file in Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.dll"))
                {
                    Debug.WriteLine($"Found DLL: {Path.GetFileName(file)}");
                }
                if (Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "drivers")))
                {
                    foreach (var file in Directory.GetFiles(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "drivers"), "*.dll"))
                    {
                        Debug.WriteLine($"Found Driver DLL: {Path.GetFileName(file)}");
                    }
                }

                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                Debug.WriteLine("MainWindow created and shown");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error during startup: {ex}");
                MessageBox.Show($"Error during startup: {ex.Message}\n\nStack trace:\n{ex.StackTrace}",
                              "Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Shutdown();
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            if (_mutex != null)
            {
                _mutex.ReleaseMutex();
                _mutex.Dispose();
            }
        }
    }
}
