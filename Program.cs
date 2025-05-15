using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using System.Security.Principal; // 添加權限相關命名空間
using System.Reflection; // 添加反射命名空間，用於獲取執行檔路徑

namespace Showdown
{
    // Shutdown Configuration Class
    internal class ShutdownConfig
    {
        public int Hour { get; set; } = 17;
        public int Minute { get; set; } = 30;
        public int Second { get; set; }
        public bool ForceShutdown { get; set; }
        public bool EnableAutoShutdown { get; set; } = true;

        public override string ToString()
        {
            return $"Shutdown time set to: {Hour}:{Minute}:{Second}" + $"{(ForceShutdown ? " (Force Shutdown)" : "")}" + $", Auto shutdown: {(EnableAutoShutdown ? "Enabled" : "Disabled")}";
        }
    }

    internal class Program
    {
        // Show window commands
        private const int SW_HIDE = 0;
        private const int SW_MINIMIZE = 6;
        private const int SW_SHOW = 5;
        // 修改設定檔和日誌檔案路徑，使用執行檔所在目錄
        private static readonly string ExecutablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
        private static readonly string ConfigPath = Path.Combine(ExecutablePath, "shutdown_config.json");
        private static readonly string LogPath = Path.Combine(ExecutablePath, "shutdown_log.txt");
        private static System.Timers.Timer? shutdownTimer; // 允許為 null
        private static NotifyIcon? trayIcon; // 允許為 null
        private static ShutdownConfig? config; // 允許為 null
        private static readonly object configLock = new object();
        private static int shutdownAttempts = 0; // 關機嘗試次數

        // Import Windows API functions for window minimization
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Initialize console
            AllocConsole();

            Console.WriteLine("=================================");
            Console.WriteLine("Auto Shutdown Tool - v1.1");
            Console.WriteLine("=================================");

            // 初始化日誌檔案
            InitializeLogFile();

            // Allow using command line arguments to test shutdown functionality immediately
            bool testShutdown = (args.Length > 0) && (args[0] == "--test-shutdown");

            // Load configuration
            config = LoadOrCreateConfig();

            // Create system tray icon
            InitializeSystemTrayIcon();

            // Hide console window if not testing shutdown
            if (!testShutdown)
            {
                HideConsoleWindow();
            }

            // Check command line arguments
            if (testShutdown)
            {
                Console.WriteLine("Testing shutdown functionality...");
                Console.WriteLine("Will attempt to execute shutdown command in 5 seconds");
                LogMessage("測試關機功能...");
                Thread.Sleep(5000);
                ExecuteDirectShutdown();
                return;
            }

            // Immediately check time and display next shutdown time
            Console.WriteLine("Checking current time and shutdown settings...");
            if (config != null)
            {
                ShowNextShutdownTime(config);
            }

            // Start shutdown timer
            if (config != null)
            {
                //記錄訊息
                StartShutdownTimer(config);
            }

            // Run application
            Application.Run();
        }

        private static void InitializeSystemTrayIcon()
        {
            trayIcon = new NotifyIcon { Icon = SystemIcons.Application, Text = "Auto Shutdown Tool", Visible = true };

            // Create context menu
            ContextMenuStrip menu = new ContextMenuStrip();

            menu.Items.Add("Show Console", null, OnShowConsole);
            menu.Items.Add("Hide Console", null, OnHideConsole);
            menu.Items.Add("-"); // Separator

            menu.Items.Add("Show Configuration", null, OnShowConfig);
            menu.Items.Add("Modify Shutdown Time", null, OnModifyTime);

            if (config != null)
            {
                ToolStripMenuItem autoShutdownItem = new ToolStripMenuItem("Enable Auto Shutdown");
                autoShutdownItem.Checked = config.EnableAutoShutdown;
                autoShutdownItem.Click += OnToggleAutoShutdown;
                menu.Items.Add(autoShutdownItem);

                ToolStripMenuItem forceShutdownItem = new ToolStripMenuItem("Force Shutdown");
                forceShutdownItem.Checked = config.ForceShutdown;
                forceShutdownItem.Click += OnToggleForceShutdown;
                menu.Items.Add(forceShutdownItem);
            }

            menu.Items.Add("-"); // Separator
            menu.Items.Add("Shutdown Now", null, OnShutdownNow);
            menu.Items.Add("Exit", null, OnExit);

            trayIcon.ContextMenuStrip = menu;

            // Double-click to toggle console visibility
            trayIcon.DoubleClick += OnTrayIconDoubleClick;
        }

        private static void OnTrayIconDoubleClick(object? sender, EventArgs e)
        {
            ToggleConsoleVisibility();
        }

        private static void OnShowConsole(object? sender, EventArgs e)
        {
            ShowConsoleWindow();
        }

        private static void OnHideConsole(object? sender, EventArgs e)
        {
            HideConsoleWindow();
        }

        private static void OnShowConfig(object? sender, EventArgs e)
        {
            ShowConsoleWindow();
            if (config != null)
            {
                Console.WriteLine($"Current configuration: {config}");
                ShowNextShutdownTime(config);
            }
            else
            {
                Console.WriteLine("Configuration is not loaded yet.");
            }
        }

        private static void OnModifyTime(object? sender, EventArgs e)
        {
            ShowConsoleWindow();
            if (config != null)
            {
                ModifyShutdownTime(config);
                SaveConfig(config);

                // Update checked state of menu items
                UpdateTrayMenuItems();
            }
            else
            {
                Console.WriteLine("Configuration is not loaded yet.");
            }
        }

        private static void OnToggleAutoShutdown(object? sender, EventArgs e)
        {
            lock (configLock)
            {
                if (config != null)
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
        }

        private static void OnToggleForceShutdown(object? sender, EventArgs e)
        {
            lock (configLock)
            {
                if (config != null)
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
        }

        private static void OnShutdownNow(object? sender, EventArgs e)
        {
            if (config != null)
            {
                ExecuteShutdown(config);
            }
        }

        private static void OnExit(object? sender, EventArgs e)
        {
            if (trayIcon != null)
            {
                trayIcon.Visible = false;
                trayIcon.Dispose();
            }

            // Stop the shutdown timer
            if (shutdownTimer != null)
            {
                shutdownTimer.Stop();
                shutdownTimer.Dispose();
            }

            Application.Exit();
        }

        private static void UpdateTrayMenuItems()
        {
            if (trayIcon?.ContextMenuStrip != null && config != null)
            {
                foreach (object? item in trayIcon.ContextMenuStrip.Items)
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

        private static void ToggleConsoleVisibility()
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

        private static bool IsConsoleVisible()
        {
            return IsWindowVisible(GetConsoleWindow());
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private static void ShowConsoleWindow()
        {
            IntPtr handle = GetConsoleWindow();

            if (handle != IntPtr.Zero)
            {
                ShowWindow(handle, SW_SHOW);
            }
        }

        private static void HideConsoleWindow()
        {
            IntPtr handle = GetConsoleWindow();

            if (handle != IntPtr.Zero)
            {
                ShowWindow(handle, SW_HIDE);
            }
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AllocConsole();

        private static ShutdownConfig LoadOrCreateConfig()
        {
            try
            {
                Console.WriteLine($"嘗試從 {ConfigPath} 讀取設定檔");
                if (File.Exists(ConfigPath))
                {
                    string json = File.ReadAllText(ConfigPath);
                    ShutdownConfig? loadedConfig = JsonSerializer.Deserialize<ShutdownConfig>(json);
                    if (loadedConfig != null)
                    {
                        Console.WriteLine($"設定檔載入成功: {loadedConfig}");
                        return loadedConfig;
                    }
                    Console.WriteLine("反序列化設定檔為 null，建立預設設定");
                }
                else
                {
                    Console.WriteLine($"設定檔不存在: {ConfigPath}，將建立預設設定");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"設定檔讀取錯誤: {ex.Message}");
            }

            // Create default configuration
            ShutdownConfig defaultConfig = new ShutdownConfig();
            SaveConfig(defaultConfig);
            Console.WriteLine($"已建立預設設定: {defaultConfig}");
            return defaultConfig;
        }

        private static void SaveConfig(ShutdownConfig config)
        {
            try
            {
                JsonSerializerOptions options = new JsonSerializerOptions { WriteIndented = true };
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

        private static void UpdateTrayIconTooltip()
        {
            if (trayIcon != null && config != null)
            {
                DateTime shutdownTime = GetNextShutdownTime(config);
                string status = config.EnableAutoShutdown ? $"Next shutdown: {shutdownTime.ToShortTimeString()}" : "Auto shutdown disabled";

                trayIcon.Text = $"Auto Shutdown Tool - {status}";
            }
        }

        private static DateTime GetNextShutdownTime(ShutdownConfig config)
        {
            // 確保 config 不為 null，若為 null 則使用預設值
            if (config == null)
            {
                Console.WriteLine("警告: config 為 null，使用預設值");
                config = new ShutdownConfig();
            }

            DateTime now = DateTime.Now;
            DateTime shutdownTime = new DateTime(now.Year, now.Month, now.Day, config.Hour, config.Minute, config.Second);

            // If the shutdown time is in the past for today, set it for tomorrow
            if (now > shutdownTime)
            {
                shutdownTime = shutdownTime.AddDays(1);
            }

            return shutdownTime;
        }

        private static void ShowNextShutdownTime(ShutdownConfig config)
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

        private static void StartShutdownTimer(ShutdownConfig config)
        {
            // Stop existing timer if running
            if (shutdownTimer != null)
            {
                shutdownTimer.Stop();
                shutdownTimer.Dispose();
            }
            
            // 建立計時器，每1秒檢查一次，提高準確性
            shutdownTimer = new System.Timers.Timer(1000); // 1秒
            shutdownTimer.Elapsed += (sender, e) =>
            {
                try
                {
                    CheckShutdownTime(config);
                }
                catch (Exception ex)
                {
                    LogMessage($"[TIMER] Exception in Elapsed handler: {ex}");
                }
            };
            shutdownTimer.AutoReset = true;
            shutdownTimer.Start();

            Console.WriteLine("Shutdown timer started - checking every second");

            // Do initial check immediately
            CheckShutdownTime(config);
        }

        private static void CheckShutdownTime(ShutdownConfig config)
        {
            try
            {
                // 取得一次當前時間以及關機時間，避免重複計算
                DateTime currentTime = DateTime.Now;
                DateTime scheduledShutdownTime = GetNextShutdownTime(config);
                TimeSpan remainingTime = scheduledShutdownTime - currentTime;

                // 根據時間點決定日誌記錄等級，避免產生過大的日誌檔案
                // 只在以下情況記錄詳細日誌：
                // 1. 接近關機時間（5分鐘內）
                // 2. 每10分鐘的整點記錄一次
                // 3. 每小時的整點記錄一次
                bool isNearShutdown = remainingTime.TotalMinutes <= 5; // 接近關機時間時記錄詳細日誌

                // 使用目前的系統時間來判斷是否為整點時間，而不是用剩餘時間
                bool isHourMark = currentTime.Minute == 0 && currentTime.Second <= 1; // 每小時整點記錄
                bool isTenMinuteMark = currentTime.Minute % 10 == 0 && currentTime.Second <= 1; // 每10分鐘整點記錄

                bool isDetailedLoggingNeeded = isNearShutdown || isHourMark || isTenMinuteMark;

                if (isDetailedLoggingNeeded)
                {
                    LogMessage($"===== 檢查關機時間 =====");
                    LogMessage($"目前系統時間: {currentTime:yyyy-MM-dd HH:mm:ss.fff} ({TimeZoneInfo.Local.DisplayName})");
                    LogMessage($"預定關機時間: {scheduledShutdownTime:yyyy-MM-dd HH:mm:ss}");
                    LogMessage($"距離關機剩餘秒數: {remainingTime.TotalSeconds:F2}");
                }

                lock (configLock)
                {
                    if (!config.EnableAutoShutdown)
                    {
                        if (isDetailedLoggingNeeded)
                        {
                            LogMessage("自動關機功能已停用，跳過時間檢查。");
                        }
                        return;
                    }

                    // 在 lock 區塊中使用已計算的時間變數

                    // 計算剩餘時間的各部分，供後續使用
                    int remainHours = (int)remainingTime.TotalHours;
                    int remainMinutes = remainingTime.Minutes;
                    int remainSeconds = remainingTime.Seconds;

                    // 只在需要詳細記錄時才在控制台輸出時間訊息，避免過多輸出影響效能
                    if (isDetailedLoggingNeeded || remainingTime.TotalMinutes <= 10) // 距離關機10分鐘內也輸出
                    {
                        // Add more log information, using fixed format to avoid encoding issues
                        Console.WriteLine($"[DEBUG] Current time: {currentTime.ToString("HH:mm:ss")}");
                        Console.WriteLine($"[DEBUG] Scheduled shutdown time: {scheduledShutdownTime.ToString("HH:mm:ss")}");

                        // 輸出剩餘時間
                        Console.WriteLine($"[DEBUG] Time remaining: {remainHours} hours {remainMinutes} minutes {remainSeconds} seconds (total {(int)remainingTime.TotalSeconds} seconds)");
                    }

                    // 重要：使用總秒數進行更精確的比較，避免小數點問題
                    double totalSeconds = remainingTime.TotalSeconds;
                    
                    // 如果時間已到或已過（允許0.5秒誤差），立即執行關機
                    if (totalSeconds <= 0.5 && totalSeconds > -60) // 允許0.5秒誤差，且不超過1分鐘的過期時間
                    {
                        LogMessage("============================================================");
                        LogMessage($"預定關機時間已到達 (剩餘秒數: {totalSeconds:F2})，立即執行關機命令");
                        LogMessage($"目前系統時間: {currentTime:yyyy-MM-dd HH:mm:ss.fff}");
                        LogMessage($"預定關機時間: {scheduledShutdownTime:yyyy-MM-dd HH:mm:ss}");
                        LogMessage("============================================================");
                        
                        Console.WriteLine("預定關機時間已到達，準備關機...");

                        // 顯示通知
                        if (trayIcon != null)
                        {
                            trayIcon.ShowBalloonTip(3000, "電腦關機", "電腦正在關機中...", ToolTipIcon.Info);
                        }

                        // 強制設定為使用強制關機參數以確保關機執行
                        config.ForceShutdown = true;

                        // 執行關機前記錄
                        LogMessage("準備呼叫 ExecuteDirectShutdown() 方法");
                        
                        // 執行關機，不等待
                        ExecuteDirectShutdown();
                        
                        // 記錄關機方法已被呼叫
                        LogMessage("已呼叫 ExecuteDirectShutdown() 方法");

                        // 停止計時器
                        if (shutdownTimer != null)
                        {
                            shutdownTimer.Stop();
                            LogMessage("計時器已停止");
                        }
                        return;
                    }
                    
                    // 如果時間低於10秒，準備關機
                    if ((remainingTime.TotalSeconds <= 10) && (remainingTime.TotalSeconds > 0))
                    {
                        Console.WriteLine($"Shutdown time approaching! Will shutdown at {scheduledShutdownTime.ToString("HH:mm:ss")} ({remainSeconds} seconds remaining)");

                        // Show balloon tip
                        if (trayIcon != null)
                        {
                            trayIcon.ShowBalloonTip(5000, "Shutdown Imminent", $"Computer will shutdown in {remainSeconds} seconds.", ToolTipIcon.Warning);
                        }

                        // 如果時間很接近（2秒以內），執行關機
                        if (remainingTime.TotalSeconds <= 2)
                        {
                            LogMessage("============================================================");
                            LogMessage("關機時間到達2秒內，立即執行關機");
                            LogMessage($"目前系統時間: {currentTime:yyyy-MM-dd HH:mm:ss.fff}");
                            LogMessage($"預定關機時間: {scheduledShutdownTime:yyyy-MM-dd HH:mm:ss}");
                            LogMessage($"剩餘時間: {remainingTime.TotalSeconds:F2} 秒");
                            LogMessage("============================================================");
                            Console.WriteLine("Executing shutdown command...");

                            // 最後通知
                            if (trayIcon != null)
                            {
                                trayIcon.ShowBalloonTip(3000, "Shutting Down", "Computer is now shutting down...", ToolTipIcon.Info);
                            }

                            // 立即執行關機，不等待
                            config.ForceShutdown = true; // 確保使用強制關機參數
                            LogMessage("準備呼叫 ExecuteDirectShutdown() 方法");
                            ExecuteDirectShutdown();
                            LogMessage("已呼叫 ExecuteDirectShutdown() 方法");

                            // 停止計時器
                            if (shutdownTimer != null)
                            {
                                shutdownTimer.Stop();
                                LogMessage("計時器已停止");
                            }
                        }
                    }
                    // 如果在1分鐘內關機
                    else if ((remainingTime.TotalMinutes <= 1) && (remainingTime.TotalSeconds > 0))
                    {
                        Console.WriteLine($"Shutdown time approaching! Will shutdown in {remainMinutes} minutes and {remainSeconds} seconds");
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"檢查關機時間發生錯誤: {ex.Message}";
                LogMessage(errorMsg);
                LogMessage($"詳細錯誤: {ex}");
                Console.WriteLine($"Error checking shutdown time: {ex.Message}");
                Console.WriteLine($"Error details: {ex}");
            }
        }

        // 初始化日誌檔案
        private static void InitializeLogFile()
        {
            try
            {
                // 首先在控制台輸出路徑信息，這不依賴於日誌文件
                Console.WriteLine($"===== 路徑信息 =====");
                Console.WriteLine($"執行檔路徑: {ExecutablePath}");
                Console.WriteLine($"設定檔路徑: {ConfigPath}");
                Console.WriteLine($"日誌檔案路徑: {LogPath}");
                
                // 確保日誌檔案所在目錄存在
                string? logDirectory = Path.GetDirectoryName(LogPath);
                if (!string.IsNullOrEmpty(logDirectory) && !Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                    Console.WriteLine($"建立日誌目錄: {logDirectory}");
                }
                
                // 定義日誌檔案大小限制 (10MB)
                const long MaxLogFileSize = 10 * 1024 * 1024;

                bool shouldCreateNewLog = false;
                string rotatedLogPath = string.Empty;

                // 檢查日誌檔案是否存在
                if (File.Exists(LogPath))
                {
                    FileInfo logFileInfo = new FileInfo(LogPath);

                    // 若檔案超過大小限制，進行日誌輪替
                    if (logFileInfo.Length > MaxLogFileSize)
                    {
                        shouldCreateNewLog = true;
                        // 建立以日期時間為檔名的舊日誌檔案
                        rotatedLogPath = Path.Combine(
                            Path.GetDirectoryName(LogPath) ?? "",
                            $"shutdown_log_{DateTime.Now:yyyyMMdd_HHmmss}.old.txt");

                        try
                        {
                            // 重新命名現有日誌檔案
                            File.Move(LogPath, rotatedLogPath);
                            Console.WriteLine($"日誌檔案過大，已重新命名為: {rotatedLogPath}");
                        }
                        catch (Exception renameEx)
                        {
                            // 若重新命名失敗，則直接清空檔案
                            Console.WriteLine($"無法重新命名日誌檔案: {renameEx.Message}，將清空檔案內容");
                            File.WriteAllText(LogPath, string.Empty);
                        }
                    }
                }
                else
                {
                    // 若檔案不存在，需要建立新檔案
                    shouldCreateNewLog = true;
                    Console.WriteLine($"日誌檔案不存在，將創建新檔案");
                }

                // 建立新日誌檔案
                if (shouldCreateNewLog)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string initialLogEntry = $"[{timestamp}] 日誌檔案初始化成功";
                    if (!string.IsNullOrEmpty(rotatedLogPath))
                    {
                        initialLogEntry += $" (舊日誌已保存至 {Path.GetFileName(rotatedLogPath)})";
                    }
                    
                    // 添加路徑信息到日誌
                    initialLogEntry += Environment.NewLine + 
                        $"[{timestamp}] 執行檔路徑: {ExecutablePath}" + Environment.NewLine + 
                        $"[{timestamp}] 設定檔路徑: {ConfigPath}" + Environment.NewLine + 
                        $"[{timestamp}] 日誌檔案路徑: {LogPath}";
                    
                    File.WriteAllText(LogPath, initialLogEntry + Environment.NewLine);
                    Console.WriteLine("日誌檔案初始化成功");
                }

                // 測試寫入權限 - 直接寫入檔案，不使用LogMessage方法避免循環依賴
                string startupMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 程式啟動，日誌系統初始化";
                File.AppendAllText(LogPath, startupMessage + Environment.NewLine);
                Console.WriteLine(startupMessage);
                
                // 寫入路徑資訊
                string pathInfo = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 使用中路徑 - 執行檔: {ExecutablePath}" + Environment.NewLine +
                                 $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 使用中路徑 - 設定檔: {ConfigPath}" + Environment.NewLine +
                                 $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] 使用中路徑 - 日誌檔: {LogPath}";
                File.AppendAllText(LogPath, pathInfo + Environment.NewLine);
                Console.WriteLine(pathInfo.Replace(Environment.NewLine, " | "));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"無法初始化日誌檔案: {ex.Message}");
                Console.WriteLine($"詳細錯誤: {ex}");

                // 嘗試使用備用路徑 (當前目錄)
                try
                {
                    string alternativeLogPath = Path.Combine(Environment.CurrentDirectory, "shutdown_log_alt.txt");
                    File.WriteAllText(alternativeLogPath, $"[{DateTime.Now}] 原始日誌初始化失敗，使用替代路徑" + Environment.NewLine);
                    File.AppendAllText(alternativeLogPath, $"[{DateTime.Now}] 執行檔路徑: {ExecutablePath}" + Environment.NewLine);
                    File.AppendAllText(alternativeLogPath, $"[{DateTime.Now}] 設定檔路徑: {ConfigPath}" + Environment.NewLine);
                    File.AppendAllText(alternativeLogPath, $"[{DateTime.Now}] 日誌檔案路徑: {LogPath}" + Environment.NewLine);
                    File.AppendAllText(alternativeLogPath, $"[{DateTime.Now}] 錯誤訊息: {ex.Message}" + Environment.NewLine);
                    Console.WriteLine($"已使用替代路徑寫入日誌: {alternativeLogPath}");
                }
                catch (Exception altEx)
                {
                    // 忽略備用路徑的錯誤
                    Console.WriteLine($"備用日誌也失敗: {altEx.Message}");
                }
            }
        }

        // 新增記錄方法
        private static void LogMessage(string message)
        {
            try
            {
                // 寫入前檢查日誌檔案大小，若超過10MB則自動切換
                const long MaxLogFileSize = 10 * 1024 * 1024;
                if (File.Exists(LogPath))
                {
                    FileInfo logFileInfo = new FileInfo(LogPath);
                    if (logFileInfo.Length > MaxLogFileSize)
                    {
                        string rotatedLogPath = Path.Combine(
                            Path.GetDirectoryName(LogPath) ?? "",
                            $"shutdown_log_{DateTime.Now:yyyyMMdd_HHmmss}.old.txt");
                        try
                        {
                            File.Move(LogPath, rotatedLogPath);
                            Console.WriteLine($"日誌檔案過大，自動切換為: {rotatedLogPath}");
                            File.WriteAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 日誌檔案自動切換 (舊檔: {Path.GetFileName(rotatedLogPath)})\n");
                        }
                        catch (Exception renameEx)
                        {
                            Console.WriteLine($"無法自動切換日誌檔案: {renameEx.Message}，將清空檔案內容");
                            File.WriteAllText(LogPath, string.Empty);
                        }
                    }
                }
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                File.AppendAllText(LogPath, logEntry + "\r\n");
                // 也輸出到 Console 方便即時觀察
                Console.WriteLine(logEntry);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[LogMessage Exception]: {ex.Message}");
            }
        }

        private static void RestartShutdownTimer(ShutdownConfig config)
        {
            // Simply restart the timer with new config
            StartShutdownTimer(config);
        }

        private static void ModifyShutdownTime(ShutdownConfig config)
        {
            Console.WriteLine("Modify shutdown time (press Enter to keep current value)");

            Console.Write($"Hour (0-23) [{config.Hour}]: ");
            string hourInput = Console.ReadLine() ?? "";

            if (!string.IsNullOrWhiteSpace(hourInput) && int.TryParse(hourInput, out int hour) && (hour >= 0) && (hour <= 23))
            {
                config.Hour = hour;
            }

            Console.Write($"Minute (0-59) [{config.Minute}]: ");
            string minuteInput = Console.ReadLine() ?? "";

            if (!string.IsNullOrWhiteSpace(minuteInput) && int.TryParse(minuteInput, out int minute) && (minute >= 0) && (minute <= 59))
            {
                config.Minute = minute;
            }

            Console.Write($"Second (0-59) [{config.Second}]: ");
            string secondInput = Console.ReadLine() ?? "";

            if (!string.IsNullOrWhiteSpace(secondInput) && int.TryParse(secondInput, out int second) && (second >= 0) && (second <= 59))
            {
                config.Second = second;
            }

            Console.WriteLine($"New shutdown time set to: {config.Hour}:{config.Minute}:{config.Second}");
        }

        private static void ExecuteShutdown(ShutdownConfig config)
        {
            // 確保 config 不為 null
            if (config == null)
            {
                LogMessage("執行關機錯誤: 設定檔為 null");
                return;
            }

            try
            {
                // 增加關機嘗試次數
                shutdownAttempts++;
                LogMessage($"開始第 {shutdownAttempts} 次關機嘗試 (由使用者手動觸發)");

                string command;

                if (config.ForceShutdown)
                {
                    command = "shutdown /s /f /t 0"; // 強制立即關機
                }
                else
                {
                    // 一般關機指令 (使用 /s 參數關機)
                    command = "shutdown /s /t 0";
                }

                LogMessage($"執行關機命令: {command}");
                Console.WriteLine($"Executing shutdown command: {command}");

                // 使用管理員權限執行
                ProcessStartInfo processInfo = new ProcessStartInfo("cmd.exe", $"/c {command}")
                {
                    CreateNoWindow = true,
                    UseShellExecute = true,
                    Verb = "runas"
                };

                Process.Start(processInfo);
                LogMessage("關機命令已發送!");

                // 確保程序有足夠時間執行
                Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                string errorMsg = $"手動關機執行錯誤: {ex.Message}";
                LogMessage(errorMsg);
                Console.WriteLine($"Error executing shutdown: {ex.Message}");

                // 嘗試使用替代方法
                try
                {
                    LogMessage("嘗試使用直接方法關機...");
                    ExecuteDirectShutdown();
                }
                catch (Exception ex2)
                {
                    LogMessage($"所有關機方法都失敗: {ex2.Message}");
                }
            }
        }

        // Directly execute shutdown without using cmd intermediate layer
        private static void ExecuteDirectShutdown()
        {
            try
            {
                // 確保 config 不為 null
                ShutdownConfig localConfig = config ?? new ShutdownConfig();
                
                // 重要：直接寫入日誌檔案而不是通過LogMessage方法，確保即使有例外也能記錄
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string message = $"[{timestamp}] 開始執行關機程序";
                try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                Console.WriteLine(message);

                // 增加關機嘗試次數
                shutdownAttempts++;
                message = $"[{timestamp}] 第 {shutdownAttempts} 次關機嘗試";
                try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                Console.WriteLine(message);

                // 記錄系統當前狀態以便診斷
                message = $"[{timestamp}] 作業系統: {Environment.OSVersion}";
                try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                Console.WriteLine(message);
                
                message = $"[{timestamp}] 是否以管理員身分執行: {IsRunningAsAdministrator()}";
                try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                Console.WriteLine(message);

                message = $"[{timestamp}] 當前資料夾: {Environment.CurrentDirectory}";
                try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                Console.WriteLine(message);

                // 嘗試使用強制關機參數，確保強制關機
                string arguments = "/s /f /t 0"; // 始終使用強制關機參數
                message = $"[{timestamp}] 執行關機命令: shutdown {arguments}";
                try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                Console.WriteLine(message);

                // 執行關機命令
                try
                {
                    // 使用 ProcessStartInfo 直接執行關機命令，使用系統管理員權限
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "shutdown",
                        Arguments = arguments,
                        CreateNoWindow = true,
                        UseShellExecute = true, // true 允許提升權限
                        Verb = "runas" // 以管理員權限執行
                    };

                    // 啟動進程
                    message = $"[{timestamp}] 嘗試啟動關機進程...";
                    try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                    Console.WriteLine(message);
                    
                    Process? process = Process.Start(psi);
                    
                    if (process != null)
                    {
                        message = $"[{timestamp}] 關機進程已啟動，進程ID: {process.Id}";
                        try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                        Console.WriteLine(message);
                    }
                    else
                    {
                        message = $"[{timestamp}] 警告：關機進程啟動返回null";
                        try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                        Console.WriteLine(message);
                    }
                    
                    // 確保有足夠時間執行
                    message = $"[{timestamp}] 等待關機執行...";
                    try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                    Console.WriteLine(message);
                    
                    Thread.Sleep(5000);
                }
                catch (Exception processEx)
                {
                    message = $"[{timestamp}] 關機進程啟動失敗: {processEx.Message}";
                    try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                    Console.WriteLine(message);

                    // 嘗試使用CMD執行
                    message = $"[{timestamp}] 嘗試使用CMD執行關機...";
                    try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                    Console.WriteLine(message);
                    
                    try
                    {
                        ProcessStartInfo cmdPsi = new ProcessStartInfo
                        {
                            FileName = "cmd.exe",
                            Arguments = $"/c shutdown /s /f /t 0",
                            CreateNoWindow = true,
                            UseShellExecute = true,
                            Verb = "runas"
                        };
                        
                        Process? cmdProcess = Process.Start(cmdPsi);
                        
                        if (cmdProcess != null)
                        {
                            message = $"[{timestamp}] CMD關機進程已啟動，進程ID: {cmdProcess.Id}";
                            try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                            Console.WriteLine(message);
                        }
                        
                        Thread.Sleep(5000);
                    }
                    catch (Exception cmdEx)
                    {
                        message = $"[{timestamp}] CMD關機方法失敗: {cmdEx.Message}";
                        try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                        Console.WriteLine(message);
                        
                        // 最後嘗試
                        message = $"[{timestamp}] 嘗試直接執行shutdown命令...";
                        try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                        Console.WriteLine(message);
                        
                        try
                        {
                            Process.Start("shutdown", "-s -f -t 0");
                            Thread.Sleep(5000);
                        }
                        catch (Exception finalEx)
                        {
                            message = $"[{timestamp}] 所有關機方法嘗試失敗: {finalEx.Message}";
                            try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                            Console.WriteLine(message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 最後的例外處理，確保記錄
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string message = $"[{timestamp}] 嚴重錯誤在ExecuteDirectShutdown方法: {ex.Message}";
                try { File.AppendAllText(LogPath, message + Environment.NewLine); } catch { }
                Console.WriteLine(message);

                try
                {
                    // 嘗試使用Windows API直接關機
                    InitiateSystemShutdownEx("", "系統排程關機", 0, true, true, 0x00000005);
                }
                catch
                {
                    // 忽略最後嘗試的錯誤
                }
            }
        }
        
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool InitiateSystemShutdownEx(
            string lpMachineName,
            string lpMessage,
            uint dwTimeout,
            bool bForceAppsClosed,
            bool bRebootAfterShutdown,
            uint dwReason);

        // 檢查是否以系統管理員身分運行
        private static bool IsRunningAsAdministrator()
        {
            try
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (Exception ex)
            {
                LogMessage($"檢查管理員權限時發生錯誤: {ex.Message}");
                return false;
            }
        }
    }
}
