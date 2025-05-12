using System;
using System.IO;
using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;

namespace Showdown
{
    // Shutdown Configuration Class
    class ShutdownConfig
    {
        public int Hour { get; set; } = 17;
        public int Minute { get; set; } = 30;
        public int Second { get; set; } = 0;
        public bool ForceShutdown { get; set; } = false;

        public override string ToString()
        {
            return $"Shutdown time set to: {Hour}:{Minute}:{Second}" +
                   $"{(ForceShutdown ? " (Force Shutdown)" : "")}";
        }
    }

    class Program
    {
        private static readonly string ConfigPath = "shutdown_config.json";
        
        // Import Windows API functions for window minimization
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // Show window commands
        private const int SW_HIDE = 0;
        private const int SW_MINIMIZE = 6;

        static void Main(string[] args)
        {
            // Check if program should start minimized
            bool startMinimized = args.Length > 0 && (args[0] == "--minimized" || args[0] == "-m");
            
            if (startMinimized)
            {
                // Get handle to console window and minimize it
                IntPtr handle = GetConsoleWindow();
                ShowWindow(handle, SW_MINIMIZE);
            }
            
            Console.WriteLine("Auto Shutdown Tool - v1.0");
            
            var config = LoadOrCreateConfig();
            DisplayMenu(config);
        }

        static ShutdownConfig LoadOrCreateConfig()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    var config = JsonSerializer.Deserialize<ShutdownConfig>(json);
                    Console.WriteLine($"Configuration loaded: {config}");
                    return config;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading configuration: {ex.Message}");
            }

            // Create default configuration
            var defaultConfig = new ShutdownConfig();
            SaveConfig(defaultConfig);
            Console.WriteLine($"Default configuration created: {defaultConfig}");
            return defaultConfig;
        }

        static void SaveConfig(ShutdownConfig config)
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(config, options);
                File.WriteAllText(ConfigPath, json);
                Console.WriteLine("Configuration saved");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving configuration: {ex.Message}");
            }
        }

        static void DisplayMenu(ShutdownConfig config)
        {
            while (true)
            {
                Console.WriteLine("\n===== Auto Shutdown Tool Menu =====");
                Console.WriteLine("1. Show Current Configuration");
                Console.WriteLine("2. Modify Shutdown Time");
                Console.WriteLine("3. Toggle Force Shutdown");
                Console.WriteLine("4. Shutdown Now");
                Console.WriteLine("5. Exit Program");
                Console.Write("Choose an option (1-5): ");

                if (int.TryParse(Console.ReadLine(), out int choice))
                {
                    switch (choice)
                    {
                        case 1:
                            Console.WriteLine($"Current configuration: {config}");
                            break;
                        case 2:
                            ModifyShutdownTime(config);
                            SaveConfig(config);
                            break;
                        case 3:
                            config.ForceShutdown = !config.ForceShutdown;
                            Console.WriteLine($"Force shutdown set to: {(config.ForceShutdown ? "Enabled" : "Disabled")}");
                            SaveConfig(config);
                            break;
                        case 4:
                            ExecuteShutdown(config);
                            break;
                        case 5:
                            Console.WriteLine("Program exited");
                            return;
                        default:
                            Console.WriteLine("Invalid choice, please try again");
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Please enter a valid number");
                }
            }
        }

        static void ModifyShutdownTime(ShutdownConfig config)
        {
            Console.WriteLine("Modify shutdown time (press Enter to keep current value)");
            
            Console.Write($"Hour (0-23) [{config.Hour}]: ");
            string hourInput = Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(hourInput) && int.TryParse(hourInput, out int hour) && hour >= 0 && hour <= 23)
            {
                config.Hour = hour;
            }

            Console.Write($"Minute (0-59) [{config.Minute}]: ");
            string minuteInput = Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(minuteInput) && int.TryParse(minuteInput, out int minute) && minute >= 0 && minute <= 59)
            {
                config.Minute = minute;
            }

            Console.Write($"Second (0-59) [{config.Second}]: ");
            string secondInput = Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(secondInput) && int.TryParse(secondInput, out int second) && second >= 0 && second <= 59)
            {
                config.Second = second;
            }

            Console.WriteLine($"New shutdown time set to: {config.Hour}:{config.Minute}:{config.Second}");
        }

        static void ExecuteShutdown(ShutdownConfig config)
        {
            try
            {
                string command;
                
                if (config.ForceShutdown)
                {
                    command = "shutdown /s /f /t 0"; // Force immediate shutdown
                }
                else
                {
                    // Include system update handling (parameter /g indicates shutdown after update)
                    command = "shutdown /g /t 0";
                }

                Console.WriteLine($"Executing shutdown command: {command}");
                
                var processInfo = new ProcessStartInfo("cmd.exe", $"/c {command}")
                {
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                
                Process.Start(processInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing shutdown: {ex.Message}");
            }
        }
    }
}