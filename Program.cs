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
            // 設定控制台編碼為UTF-8，解決中文亂碼問題
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // 初始化控制台
            AllocConsole();
            
            Console.WriteLine("=================================");
            Console.WriteLine("自動關機工具 - v1.1");
            Console.WriteLine("=================================");
            
            // 允許使用命令列參數立即測試關機功能
            bool testShutdown = args.Length > 0 && args[0] == "--test-shutdown";
            
            // 載入配置
            config = LoadOrCreateConfig();
            
            // 創建系統通知區域圖示
            InitializeSystemTrayIcon();
            
            // 如果不是測試關機，則隱藏控制台
            if (!testShutdown)
            {
                HideConsoleWindow();
            }
            
            // 檢查命令列參數
            if (testShutdown)
            {
                Console.WriteLine("正在測試關機功能...");
                Console.WriteLine("將在5秒後嘗試執行關機命令");
                Thread.Sleep(5000);
                ExecuteDirectShutdown();
                return;
            }
            
            // 立即檢查時間，並顯示下次關機時間
            Console.WriteLine("檢查目前時間與關機設定...");
            ShowNextShutdownTime(config);
            
            // 啟動關機計時器
            StartShutdownTimer(config);
            
            // 運行應用
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
            
            // Create a timer that checks every 10 seconds
            shutdownTimer = new System.Timers.Timer(10000); // 10 seconds
            shutdownTimer.Elapsed += (sender, e) => CheckShutdownTime(config);
            shutdownTimer.AutoReset = true;
            shutdownTimer.Start();
            
            Console.WriteLine("關機計時器已啟動 - 每10秒檢查一次");
            
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
                    {
                        return;
                    }
                    
                    DateTime now = DateTime.Now;
                    DateTime shutdownTime = GetNextShutdownTime(config);
                    TimeSpan timeUntilShutdown = shutdownTime - now;
                    
                    // 添加更多日誌信息，使用固定格式避免亂碼
                    Console.WriteLine($"[DEBUG] 當前時間: {now.ToString("HH:mm:ss")}");
                    Console.WriteLine($"[DEBUG] 計劃關機時間: {shutdownTime.ToString("HH:mm:ss")}");
                    
                    // 格式化剩餘時間顯示
                    int remainHours = (int)timeUntilShutdown.TotalHours;
                    int remainMinutes = timeUntilShutdown.Minutes;
                    int remainSeconds = timeUntilShutdown.Seconds;
                    Console.WriteLine($"[DEBUG] 剩餘時間: {remainHours}小時 {remainMinutes}分鐘 {remainSeconds}秒 (共{(int)timeUntilShutdown.TotalSeconds}秒)");
                    
                    // 直接測試關機功能 - 如果需要測試，取消這段註釋
                    // if (true) {
                    //     Console.WriteLine("正在測試關機功能...");
                    //     ShutdownNow();
                    //     return;
                    // }
                    
                    // 如果時間小於10秒，直接執行關機
                    if (timeUntilShutdown.TotalSeconds <= 10 && timeUntilShutdown.TotalSeconds > 0)
                    {
                        Console.WriteLine($"關機時間即將到來！將在 {shutdownTime.ToString("HH:mm:ss")} 關機 (還有 {remainSeconds} 秒)");
                        
                        // 顯示氣球提示
                        if (trayIcon != null)
                        {
                            trayIcon.ShowBalloonTip(
                                5000,
                                "即將關機",
                                $"電腦將在 {remainSeconds} 秒後關機。",
                                ToolTipIcon.Warning
                            );
                        }
                        
                        // 如果時間非常接近（2秒內），執行關機
                        if (timeUntilShutdown.TotalSeconds <= 2)
                        {
                            Console.WriteLine("正在執行關機指令...");
                            
                            // 最終通知
                            if (trayIcon != null)
                            {
                                trayIcon.ShowBalloonTip(
                                    3000, 
                                    "正在關機",
                                    "電腦正在關機...",
                                    ToolTipIcon.Info
                                );
                            }
                            
                            // 立即執行關機，不等待
                            ExecuteDirectShutdown();
                            
                            // 停止計時器
                            shutdownTimer.Stop();
                        }
                    }
                    // 如果在關機時間的1分鐘內
                    else if (timeUntilShutdown.TotalMinutes <=
1 && timeUntilShutdown.TotalSeconds > 0)
                    {
                        Console.WriteLine($"關機時間接近！將在 {remainMinutes}分鐘{remainSeconds}秒後關機");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"檢查關機時間時發生錯誤: {ex.Message}");
                Console.WriteLine($"錯誤詳情: {ex.ToString()}");
            }
        }
        
        // 直接執行關機，不使用cmd中間層
        static void ExecuteDirectShutdown()
        {
            try
            {
                // 使用ProcessStartInfo直接執行shutdown命令
                Console.WriteLine("正在執行關機命令...");
                
                // 使用進程啟動關機命令
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "shutdown",
                    Arguments = "/s /t 0", // 立即關機
                    CreateNoWindow = true,
                    UseShellExecute = false
                };
                
                Process.Start(psi);
                Console.WriteLine("關機命令已發送！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"執行關機時發生錯誤: {ex.Message}");
                
                // 嘗試使用備用方法
                try
                {
                    Console.WriteLine("嘗試使用備用方法關機...");
                    Process.Start("shutdown", "/s /t 0");
                    Console.WriteLine("備用關機命令已發送！");
                }
                catch (Exception ex2)
                {
                    Console.WriteLine($"備用關機方法也失敗: {ex2.Message}");
                }
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