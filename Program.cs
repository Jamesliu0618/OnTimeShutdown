using System;
using System.IO;
using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using System.Drawing;

namespace Showdown
{
    // Shutdown Configuration Class
    class ShutdownConfig
    {
        public int Hour { get; set; } = 17;
        public int Minute { get; set; } = 30;
        public int Second { get; set; } = 0;
        public bool ForceShutdown { get; set; } = false;
        public bool EnableAutoShutdown { get; set; } = true;

        public override string ToString()
        {
            return $"Shutdown time set to: {Hour}:{Minute}:{Second}" +
                   $"{(ForceShutdown ? " (Force Shutdown)" : "")}" +
                   $", Auto shutdown: {(EnableAutoShutdown ? "Enabled" : "Disabled")}";
        }
    }

    class Program
    {
        private static readonly string ConfigPath = "shutdown_config.json";
        private static System.Timers.Timer shutdownTimer;
        private static NotifyIcon trayIcon;
        private static ShutdownConfig config;
        private static readonly Object configLock = new Object();
        
        // Import Windows API functions for window minimization
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // Show window commands
        private const int SW_HIDE = 0;
        private const int SW_MINIMIZE = 6;
        private const int SW_SHOW = 5;

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Initialize console for logging
            AllocConsole();
            
            Console.WriteLine("=================================");
            Console.WriteLine("Auto Shutdown Tool - v1.0");
            Console.WriteLine("=================================");
            
            // Load or create config
            config = LoadOrCreateConfig();
            
            // Create system tray icon and menu
            InitializeSystemTrayIcon();
            
            // Always hide the console window on startup
            HideConsoleWindow();
            
            // Immediately check current time against shutdown time
            Console.WriteLine("Checking current time against shutdown settings...");
            ShowNextShutdownTime(config);
            
            // Start the shutdown timer
            StartShutdownTimer(config);
            
            // Run the application
            Application.Run();
        }
        
        static void InitializeSystemTrayIcon()
        {
            trayIcon = new NotifyIcon()
            {
                Icon = SystemIcons.Application,
                Text = "Auto Shutdown Tool",
                Visible = true
            };
            
            // Create context menu
            ContextMenuStrip menu = new ContextMenuStrip();
            
            menu.Items.Add("Show Console", null, OnShowConsole);
            menu.Items.Add("Hide Console", null, OnHideConsole);
            menu.Items.Add("-"); // Separator
            
            menu.Items.Add("Show Configuration", null, OnShowConfig);
            menu.Items.Add("Modify Shutdown Time", null, OnModifyTime);
            
            ToolStripMenuItem autoShutdownItem = new ToolStripMenuItem("Enable Auto Shutdown");
            autoShutdownItem.Checked = config.EnableAutoShutdown;
            autoShutdownItem.Click += OnToggleAutoShutdown;
            menu.Items.Add(autoShutdownItem);
            
            ToolStripMenuItem forceShutdownItem = new ToolStripMenuItem("Force Shutdown");
            forceShutdownItem.Checked = config.ForceShutdown;
            forceShutdownItem.Click += OnToggleForceShutdown;
            menu.Items.Add(forceShutdownItem);
            
            menu.Items.Add("-"); // Separator
            menu.Items.Add("Shutdown Now", null, OnShutdownNow);
            menu.Items.Add("Exit", null, OnExit);
            
            trayIcon.ContextMenuStrip = menu;
            
            // Double-click to toggle console visibility
            trayIcon.DoubleClick += OnTrayIconDoubleClick;
        }
        
        static void OnTrayIconDoubleClick(object sender, EventArgs e)
        {
            ToggleConsoleVisibility();
        }
        
        static void OnShowConsole(object sender, EventArgs e)
        {
            ShowConsoleWindow();
        }
        
        static void OnHideConsole(object sender, EventArgs e)
        {
            HideConsoleWindow();
        }
        
        static void OnShowConfig(object sender, EventArgs e)
        {
            ShowConsoleWindow();
            Console.WriteLine($"Current configuration: {config}");
            ShowNextShutdownTime(config);
        }
        
        static void OnModifyTime(object sender, EventArgs e)
        {
            ShowConsoleWindow();
            ModifyShutdownTime(config);
            SaveConfig(config);
            
            // Update checked state of menu items
            UpdateTrayMenuItems();
        }
        
        static void OnToggleAutoShutdown(object sender, EventArgs e)
        {
            lock (configLock)
            {
                config.EnableAutoShutdown = !config.EnableAutoShutdown;
                Console.WriteLine($"Auto shutdown set to: {(config.EnableAutoShutdown ? "Enabled" : "Disabled")}");
                SaveConfig(config);
                
                // Update menu item checked state
                if (sender is ToolStripMenuItem item)
                {
                    item.Checked = config.EnableAutoShutdown;
                }
            }
        }
        
        static void OnToggleForceShutdown(object sender, EventArgs e)
        {
            lock (configLock)
            {
                config.ForceShutdown = !config.ForceShutdown;
                Console.WriteLine($"Force shutdown set to: {(config.ForceShutdown ? "Enabled" : "Disabled")}");
                SaveConfig(config);
                
                // Update menu item checked state
                if (sender is ToolStripMenuItem item)
                {
                    item.Checked = config.ForceShutdown;
                }
            }
        }
        
        static void OnShutdownNow(object sender, EventArgs e)
        {
            ExecuteShutdown(config);
        }
        
        static void OnExit(object sender, EventArgs e)
        {
            trayIcon.Visible = false;
            trayIcon.Dispose();
            
            // Stop the shutdown timer
            if (shutdownTimer != null)
            {
                shutdownTimer.Stop();
                shutdownTimer.Dispose();
            }
            
            Application.Exit();
        }
        
        static void UpdateTrayMenuItems()
        {
            if (trayIcon?.ContextMenuStrip != null)
            {
                foreach (var item in trayIcon.ContextMenuStrip.Items)
                {
                    if (item is ToolStripMenuItem menuItem)
                    {
                        if (menuItem.Text == "Enable Auto Shutdown")
                        {
                            menuItem.Checked = config.EnableAutoShutdown;
                        }
                        else if (menuItem.Text == "Force Shutdown")
                        {
                            menuItem.Checked = config.ForceShutdown;
                        }
                    }
                }
            }
        }
        
        static void ToggleConsoleVisibility()
        {
            IntPtr handle = GetConsoleWindow();
            if (handle != IntPtr.Zero)
            {
                if (IsConsoleVisible())
                {
                    ShowWindow(handle, SW_HIDE);
                }
                else
                {
                    ShowWindow(handle, SW_SHOW);
                }
            }
        }
        
        static bool IsConsoleVisible()
        {
            return IsWindowVisible(GetConsoleWindow());
        }
        
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindowVisible(IntPtr hWnd);
        
        static void ShowConsoleWindow()
        {
            IntPtr handle = GetConsoleWindow();
            if (handle != IntPtr.Zero)
            {
                ShowWindow(handle, SW_SHOW);
            }
        }
        
        static void HideConsoleWindow()
        {
            IntPtr handle = GetConsoleWindow();
            if (handle != IntPtr.Zero)
            {
                ShowWindow(handle, SW_HIDE);
            }
        }
        
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

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
                
                // Show updated shutdown time
                ShowNextShutdownTime(config);
                
                // Restart the shutdown timer with updated settings
                RestartShutdownTimer(config);
                
                // Update tray icon tooltip
                UpdateTrayIconTooltip();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving configuration: {ex.Message}");
            }
        }
        
        static void UpdateTrayIconTooltip()
        {
            if (trayIcon != null)
            {
                DateTime shutdownTime = GetNextShutdownTime(config);
                string status = config.EnableAutoShutdown 
                    ? $"Next shutdown: {shutdownTime.ToShortTimeString()}" 
                    : "Auto shutdown disabled";
                
                trayIcon.Text = $"Auto Shutdown Tool - {status}";
            }
        }
        
        static DateTime GetNextShutdownTime(ShutdownConfig config)
        {
            DateTime now = DateTime.Now;
            DateTime shutdownTime = new DateTime(
                now.Year, now.Month, now.Day, 
                config.Hour, config.Minute, config.Second);
                
            // If the shutdown time is in the past for today, set it for tomorrow
            if (now > shutdownTime)
            {
                shutdownTime = shutdownTime.AddDays(1);
            }
            
            return shutdownTime;
        }

        static void ShowNextShutdownTime(ShutdownConfig config)
        {
            if (!config.EnableAutoShutdown)
            {
                Console.WriteLine("Auto shutdown is currently DISABLED");
                return;
            }
            
            DateTime now = DateTime.Now;
            DateTime shutdownTime = GetNextShutdownTime(config);
            
            // Update tray icon tooltip with shutdown time
            UpdateTrayIconTooltip();
            
            // If shutdown time is tomorrow
            if (shutdownTime.Date > now.Date)
            {
                Console.WriteLine($"Next scheduled shutdown: TOMORROW at {shutdownTime.ToShortTimeString()}");
            }
            else
            {
                TimeSpan timeLeft = shutdownTime - now;
                Console.WriteLine($"Next scheduled shutdown: TODAY at {shutdownTime.ToShortTimeString()} (in {timeLeft.Hours}h {timeLeft.Minutes}m {timeLeft.Seconds}s)");
            }
        }

        static void StartShutdownTimer(ShutdownConfig config)
        {
            // Stop existing timer if running
            if (shutdownTimer != null)
            {
                shutdownTimer.Stop();
                shutdownTimer.Dispose();
            }
            
            // Create a timer that checks every 30 seconds
            shutdownTimer = new System.Timers.Timer(30000); // 30 seconds
            shutdownTimer.Elapsed += (sender, e) => CheckShutdownTime(config);
            shutdownTimer.AutoReset = true;
            shutdownTimer.Start();
            
            Console.WriteLine("Shutdown timer started - checking every 30 seconds");
            
            // Do initial check immediately
            CheckShutdownTime(config);
        }
        
        static void CheckShutdownTime(ShutdownConfig config)
        {
            try
            {
                lock (configLock)
                {
                    if (!config.EnableAutoShutdown)
                        return;
                    
                    DateTime now = DateTime.Now;
                    DateTime shutdownTime = GetNextShutdownTime(config);
                    TimeSpan timeUntilShutdown = shutdownTime - now;
                    
                    // If it's within 1 minute of shutdown time, prepare shutdown
                    if (timeUntilShutdown.TotalMinutes <= 1 && timeUntilShutdown.TotalSeconds > 0)
                    {
                        Console.WriteLine($"Scheduled shutdown approaching. Shutdown will occur at {shutdownTime.ToShortTimeString()} (in {timeUntilShutdown.Seconds} seconds)");
                        
                        // Show balloon tip for upcoming shutdown
                        if (trayIcon != null && timeUntilShutdown.TotalSeconds <= 60 && timeUntilShutdown.TotalSeconds > 10)
                        {
                            trayIcon.ShowBalloonTip(
                                3000, 
                                "Scheduled Shutdown",
                                $"Computer will shut down in {(int)timeUntilShutdown.TotalSeconds} seconds.",
                                ToolTipIcon.Warning
                            );
                        }
                        
                        // If we're within 5 seconds of shutdown time, execute shutdown
                        if (timeUntilShutdown.TotalSeconds <= 5)
                        {
                            Console.WriteLine("EXECUTING SCHEDULED SHUTDOWN NOW");
                            
                            // Final balloon notification
                            if (trayIcon != null)
                            {
                                trayIcon.ShowBalloonTip(
                                    3000, 
                                    "Shutting Down",
                                    "Computer is shutting down now.",
                                    ToolTipIcon.Info
                                );
                            }
                            
                            ExecuteShutdown(config);
                            
                            // Stop the timer since we're shutting down
                            shutdownTimer.Stop();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking shutdown time: {ex.Message}");
            }
        }
        
        static void RestartShutdownTimer(ShutdownConfig config)
        {
            // Simply restart the timer with new config
            StartShutdownTimer(config);
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
                    // 正常關機指令 (使用 /s 參數表示關機)
                    command = "shutdown /s /t 0"; 
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