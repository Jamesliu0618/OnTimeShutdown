using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace Showdown
{
	// Shutdown Configuration Class
	internal class ShutdownConfig
	{
		public int  Hour               { get; set; } = 17;
		public int  Minute             { get; set; } = 33;
		public int  Second             { get; set; }
		public bool ForceShutdown      { get; set; }
		public bool EnableAutoShutdown { get; set; } = true;

		public override string ToString()
		{
			return $"Shutdown time set to: {Hour}:{Minute}:{Second}" + $"{(ForceShutdown ? " (Force Shutdown)" : "")}" + $", Auto shutdown: {(EnableAutoShutdown ? "Enabled" : "Disabled")}";
		}
	}

	internal class Program
	{
		// Show window commands
		private const           int                 SW_HIDE     = 0;
		private const           int                 SW_MINIMIZE = 6;
		private const           int                 SW_SHOW     = 5;
		private static readonly string              ConfigPath  = "shutdown_config.json";
		private static readonly string              LogPath     = "shutdown_log.txt"; // 新增日誌檔案路徑
		private static          System.Timers.Timer shutdownTimer;
		private static          NotifyIcon          trayIcon;
		private static          ShutdownConfig      config;
		private static readonly object              configLock = new object();
		private static          int                 shutdownAttempts = 0; // 關機嘗試次數

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

			// Allow using command line arguments to test shutdown functionality immediately
			bool testShutdown = (args.Length > 0) && (args[0] == "--test-shutdown");

			// Load configuration
			config = LoadOrCreateConfig();

			// Create system tray icon
			InitializeSystemTrayIcon();

			// Hide console window if not testing shutdown
			if(!testShutdown)
			{
				HideConsoleWindow();
			}

			// Check command line arguments
			if(testShutdown)
			{
				Console.WriteLine("Testing shutdown functionality...");
				Console.WriteLine("Will attempt to execute shutdown command in 5 seconds");
				Thread.Sleep(5000);
				ExecuteDirectShutdown();
				return;
			}

			// Immediately check time and display next shutdown time
			Console.WriteLine("Checking current time and shutdown settings...");
			ShowNextShutdownTime(config);

			// Start shutdown timer
			StartShutdownTimer(config);

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

			menu.Items.Add("Show Configuration",   null, OnShowConfig);
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
			menu.Items.Add("Exit",         null, OnExit);

			trayIcon.ContextMenuStrip = menu;

			// Double-click to toggle console visibility
			trayIcon.DoubleClick += OnTrayIconDoubleClick;
		}

		private static void OnTrayIconDoubleClick(object sender, EventArgs e)
		{
			ToggleConsoleVisibility();
		}

		private static void OnShowConsole(object sender, EventArgs e)
		{
			ShowConsoleWindow();
		}

		private static void OnHideConsole(object sender, EventArgs e)
		{
			HideConsoleWindow();
		}

		private static void OnShowConfig(object sender, EventArgs e)
		{
			ShowConsoleWindow();
			Console.WriteLine($"Current configuration: {config}");
			ShowNextShutdownTime(config);
		}

		private static void OnModifyTime(object sender, EventArgs e)
		{
			ShowConsoleWindow();
			ModifyShutdownTime(config);
			SaveConfig(config);

			// Update checked state of menu items
			UpdateTrayMenuItems();
		}

		private static void OnToggleAutoShutdown(object sender, EventArgs e)
		{
			lock(configLock)
			{
				config.EnableAutoShutdown = !config.EnableAutoShutdown;
				Console.WriteLine($"Auto shutdown set to: {(config.EnableAutoShutdown ? "Enabled" : "Disabled")}");
				SaveConfig(config);

				// Update menu item checked state
				if(sender is ToolStripMenuItem item)
				{
					item.Checked = config.EnableAutoShutdown;
				}
			}
		}

		private static void OnToggleForceShutdown(object sender, EventArgs e)
		{
			lock(configLock)
			{
				config.ForceShutdown = !config.ForceShutdown;
				Console.WriteLine($"Force shutdown set to: {(config.ForceShutdown ? "Enabled" : "Disabled")}");
				SaveConfig(config);

				// Update menu item checked state
				if(sender is ToolStripMenuItem item)
				{
					item.Checked = config.ForceShutdown;
				}
			}
		}

		private static void OnShutdownNow(object sender, EventArgs e)
		{
			ExecuteShutdown(config);
		}

		private static void OnExit(object sender, EventArgs e)
		{
			trayIcon.Visible = false;
			trayIcon.Dispose();

			// Stop the shutdown timer
			if(shutdownTimer != null)
			{
				shutdownTimer.Stop();
				shutdownTimer.Dispose();
			}

			Application.Exit();
		}

		private static void UpdateTrayMenuItems()
		{
			if(trayIcon?.ContextMenuStrip != null)
			{
				foreach(object? item in trayIcon.ContextMenuStrip.Items)
				{
					if(item is ToolStripMenuItem menuItem)
					{
						if(menuItem.Text == "Enable Auto Shutdown")
						{
							menuItem.Checked = config.EnableAutoShutdown;
						}
						else if(menuItem.Text == "Force Shutdown")
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

			if(handle != IntPtr.Zero)
			{
				if(IsConsoleVisible())
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

		[DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool IsWindowVisible(IntPtr hWnd);

		private static void ShowConsoleWindow()
		{
			IntPtr handle = GetConsoleWindow();

			if(handle != IntPtr.Zero)
			{
				ShowWindow(handle, SW_SHOW);
			}
		}

		private static void HideConsoleWindow()
		{
			IntPtr handle = GetConsoleWindow();

			if(handle != IntPtr.Zero)
			{
				ShowWindow(handle, SW_HIDE);
			}
		}

		[DllImport("kernel32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool AllocConsole();

		private static ShutdownConfig LoadOrCreateConfig()
		{
			try
			{
				if(File.Exists(ConfigPath))
				{
					string          json   = File.ReadAllText(ConfigPath);
					ShutdownConfig? config = JsonSerializer.Deserialize<ShutdownConfig>(json);
					Console.WriteLine($"Configuration loaded: {config}");
					return config;
				}
			}
			catch(Exception ex)
			{
				Console.WriteLine($"Error reading configuration: {ex.Message}");
			}

			// Create default configuration
			ShutdownConfig defaultConfig = new ShutdownConfig();
			SaveConfig(defaultConfig);
			Console.WriteLine($"Default configuration created: {defaultConfig}");
			return defaultConfig;
		}

		private static void SaveConfig(ShutdownConfig config)
		{
			try
			{
				JsonSerializerOptions options = new JsonSerializerOptions { WriteIndented = true };
				string                json    = JsonSerializer.Serialize(config, options);
				File.WriteAllText(ConfigPath, json);
				Console.WriteLine("Configuration saved");

				// Show updated shutdown time
				ShowNextShutdownTime(config);

				// Restart the shutdown timer with updated settings
				RestartShutdownTimer(config);

				// Update tray icon tooltip
				UpdateTrayIconTooltip();
			}
			catch(Exception ex)
			{
				Console.WriteLine($"Error saving configuration: {ex.Message}");
			}
		}

		private static void UpdateTrayIconTooltip()
		{
			if(trayIcon != null)
			{
				DateTime shutdownTime = GetNextShutdownTime(config);
				string   status       = config.EnableAutoShutdown ? $"Next shutdown: {shutdownTime.ToShortTimeString()}" : "Auto shutdown disabled";

				trayIcon.Text = $"Auto Shutdown Tool - {status}";
			}
		}

		private static DateTime GetNextShutdownTime(ShutdownConfig config)
		{
			DateTime now          = DateTime.Now;
			DateTime shutdownTime = new DateTime(now.Year, now.Month, now.Day, config.Hour, config.Minute, config.Second);

			// If the shutdown time is in the past for today, set it for tomorrow
			if(now > shutdownTime)
			{
				shutdownTime = shutdownTime.AddDays(1);
			}

			return shutdownTime;
		}

		private static void ShowNextShutdownTime(ShutdownConfig config)
		{
			if(!config.EnableAutoShutdown)
			{
				Console.WriteLine("Auto shutdown is currently DISABLED");
				return;
			}

			DateTime now          = DateTime.Now;
			DateTime shutdownTime = GetNextShutdownTime(config);

			// Update tray icon tooltip with shutdown time
			UpdateTrayIconTooltip();

			// If shutdown time is tomorrow
			if(shutdownTime.Date > now.Date)
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
			if(shutdownTimer != null)
			{
				shutdownTimer.Stop();
				shutdownTimer.Dispose();
			}

			// Create a timer that checks every 10 seconds
			shutdownTimer = new System.Timers.Timer(10000); // 10 seconds
			shutdownTimer.Elapsed += (sender, e) => CheckShutdownTime(config);
			shutdownTimer.AutoReset = true;
			shutdownTimer.Start();

			Console.WriteLine("Shutdown timer started - checking every 10 seconds");

			// Do initial check immediately
			CheckShutdownTime(config);
		}

		private static void CheckShutdownTime(ShutdownConfig config)
		{
			try
			{
				lock(configLock)
				{
					if(!config.EnableAutoShutdown)
					{
						return;
					}

					DateTime now               = DateTime.Now;
					DateTime shutdownTime      = GetNextShutdownTime(config);
					TimeSpan timeUntilShutdown = shutdownTime - now;

					// Add more log information, using fixed format to avoid encoding issues
					Console.WriteLine($"[DEBUG] Current time: {now.ToString("HH:mm:ss")}");
					Console.WriteLine($"[DEBUG] Scheduled shutdown time: {shutdownTime.ToString("HH:mm:ss")}");

					// Format remaining time display
					int remainHours   = (int)timeUntilShutdown.TotalHours;
					int remainMinutes = timeUntilShutdown.Minutes;
					int remainSeconds = timeUntilShutdown.Seconds;
					Console.WriteLine($"[DEBUG] Time remaining: {remainHours} hours {remainMinutes} minutes {remainSeconds} seconds (total {(int)timeUntilShutdown.TotalSeconds} seconds)");

					// 重要：使用總秒數進行更精確的比較，避免小數點問題
					double totalSeconds = timeUntilShutdown.TotalSeconds;
					
					// 如果時間已到或已過（允許0.5秒誤差），立即執行關機
					if(totalSeconds <= 0.5 && totalSeconds > -60) // 允許0.5秒誤差，且不超過1分鐘的過期時間
					{
						LogMessage("預定關機時間已到達，立即執行關機命令");
						Console.WriteLine("預定關機時間已到達，準備關機...");
						
						// 顯示通知
						if(trayIcon != null)
						{
							trayIcon.ShowBalloonTip(3000, "電腦關機", "電腦正在關機中...", ToolTipIcon.Info);
						}
						
						// 執行關機，不等待
						ExecuteDirectShutdown();
						
						// 停止計時器
						shutdownTimer.Stop();
						return;
					}
					
					// 如果時間低於10秒，準備關機
					if((timeUntilShutdown.TotalSeconds <= 10) && (timeUntilShutdown.TotalSeconds > 0))
					{
						Console.WriteLine($"Shutdown time approaching! Will shutdown at {shutdownTime.ToString("HH:mm:ss")} ({remainSeconds} seconds remaining)");

						// Show balloon tip
						if(trayIcon != null)
						{
							trayIcon.ShowBalloonTip(5000, "Shutdown Imminent", $"Computer will shutdown in {remainSeconds} seconds.", ToolTipIcon.Warning);
						}

						// 如果時間很接近（2秒以內），執行關機
						if(timeUntilShutdown.TotalSeconds <= 2)
						{
							LogMessage("關機時間到達2秒內，立即執行關機");
							Console.WriteLine("Executing shutdown command...");

							// 最後通知
							if(trayIcon != null)
							{
								trayIcon.ShowBalloonTip(3000, "Shutting Down", "Computer is now shutting down...", ToolTipIcon.Info);
							}

							// 立即執行關機，不等待
							ExecuteDirectShutdown();

							// 停止計時器
							shutdownTimer.Stop();
						}
					}

					// 如果在1分鐘內關機
					else if((timeUntilShutdown.TotalMinutes <= 1) && (timeUntilShutdown.TotalSeconds > 0))
					{
						Console.WriteLine($"Shutdown time approaching! Will shutdown in {remainMinutes} minutes and {remainSeconds} seconds");
					}
				}
			}
			catch(Exception ex)
			{
				string errorMsg = $"檢查關機時間發生錯誤: {ex.Message}";
				LogMessage(errorMsg);
				Console.WriteLine($"Error checking shutdown time: {ex.Message}");
				Console.WriteLine($"Error details: {ex}");
			}
		}
		// Directly execute shutdown without using cmd intermediate layer
		private static void ExecuteDirectShutdown()
		{
			try
			{
				// 增加關機嘗試次數
				shutdownAttempts++;
				LogMessage($"開始第 {shutdownAttempts} 次關機嘗試");
				
				// Use ProcessStartInfo to directly execute shutdown command
				Console.WriteLine("Executing shutdown command...");

				// 嘗試使用強制關機參數
				string arguments = config.ForceShutdown ? "/s /f /t 0" : "/s /t 0";
				LogMessage($"執行關機命令: shutdown {arguments}");

				// Use process to launch shutdown command
				ProcessStartInfo psi = new ProcessStartInfo
									   {
										   FileName = "shutdown",
										   Arguments = arguments,
										   CreateNoWindow = true,
										   UseShellExecute = true, // 改為true以允許提升權限
										   Verb = "runas", // 嘗試以管理員權限執行
									   };

				Process? process = Process.Start(psi);
				LogMessage("關機命令已發送!");
				Console.WriteLine("Shutdown command sent!");
				
				// 確保程序有足夠時間執行
				Thread.Sleep(3000);
			}
			catch(Exception ex)
			{
				string errorMsg = $"關機執行錯誤: {ex.Message}";
				LogMessage(errorMsg);
				Console.WriteLine($"Error executing shutdown: {ex.Message}");

				// Try alternate method
				try
				{
					LogMessage("嘗試使用替代方法關機...");
					Console.WriteLine("Attempting to use alternate method for shutdown...");
					
					// 使用cmd執行關機命令
					ProcessStartInfo cmdPsi = new ProcessStartInfo
					{
						FileName = "cmd.exe",
						Arguments = $"/c shutdown /s /f /t 0",  // 強制使用/f參數
						CreateNoWindow = true,
						UseShellExecute = true,
						Verb = "runas", // 嘗試以管理員權限執行
					};
					
					Process.Start(cmdPsi);
					LogMessage("替代關機命令已發送!");
					Console.WriteLine("Alternate shutdown command sent!");
					
					// 確保程序有足夠時間執行
					Thread.Sleep(3000);
				}
				catch(Exception ex2)
				{
					LogMessage($"替代關機方法也失敗: {ex2.Message}");
					Console.WriteLine($"Alternate shutdown method also failed: {ex2.Message}");
					
					// 最後嘗試
					try
					{
						LogMessage("嘗試最終關機方法...");
						Process.Start("shutdown", "-s -f -t 0");
						Thread.Sleep(3000);
					}
					catch (Exception ex3)
					{
						LogMessage($"所有關機方法都失敗: {ex3.Message}");
					}
				}
			}
		}
		
		// 新增記錄方法
		private static void LogMessage(string message)
		{
			try
			{
				string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
				string logEntry = $"[{timestamp}] {message}";
				
				// 寫入檔案
				File.AppendAllText(LogPath, logEntry + Environment.NewLine);
				
				// 同時在控制台顯示
				Console.WriteLine($"[LOG] {message}");
			}
			catch (Exception ex)
			{
				// 不處理記錄過程中的錯誤 - 避免遞迴
				Console.WriteLine($"Failed to write log: {ex.Message}");
			}
		}
		
		private static void RestartShutdownTimer(ShutdownConfig config)
		{
			// Simply restart the timer with new config
			StartShutdownTimer(config);
		}

		private static void ModifyShutdownTime(ShutdownConfig config)
		{
			Console.WriteLine("Modify shutdown time (press Enter to keep current value)");			Console.Write($"Hour (0-23) [{config.Hour}]: ");
			string hourInput = Console.ReadLine() ?? "";

			if(!string.IsNullOrWhiteSpace(hourInput) && int.TryParse(hourInput, out int hour) && (hour >= 0) && (hour <= 23))
			{
				config.Hour = hour;
			}

			Console.Write($"Minute (0-59) [{config.Minute}]: ");
			string minuteInput = Console.ReadLine() ?? "";

			if(!string.IsNullOrWhiteSpace(minuteInput) && int.TryParse(minuteInput, out int minute) && (minute >= 0) && (minute <= 59))
			{
				config.Minute = minute;
			}

			Console.Write($"Second (0-59) [{config.Second}]: ");
			string secondInput = Console.ReadLine() ?? "";

			if(!string.IsNullOrWhiteSpace(secondInput) && int.TryParse(secondInput, out int second) && (second >= 0) && (second <= 59))
			{
				config.Second = second;
			}

			Console.WriteLine($"New shutdown time set to: {config.Hour}:{config.Minute}:{config.Second}");
		}

		private static void ExecuteShutdown(ShutdownConfig config)
		{
			try
			{
				// 增加關機嘗試次數
				shutdownAttempts++;
				LogMessage($"開始第 {shutdownAttempts} 次關機嘗試 (由使用者手動觸發)");
				
				string command;

				if(config.ForceShutdown)
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
				ProcessStartInfo processInfo = new ProcessStartInfo("cmd.exe", $"/c {command}") { 
					CreateNoWindow = true, 
					UseShellExecute = true,
					Verb = "runas" 
				};

				Process.Start(processInfo);
				LogMessage("關機命令已發送!");
				
				// 確保程序有足夠時間執行
				Thread.Sleep(3000);
			}
			catch(Exception ex)
			{
				string errorMsg = $"手動關機執行錯誤: {ex.Message}";
				LogMessage(errorMsg);
				Console.WriteLine($"Error executing shutdown: {ex.Message}");
				
				// 嘗試使用替代方法
				try {
					LogMessage("嘗試使用直接方法關機...");
					ExecuteDirectShutdown();
				}
				catch (Exception ex2) {
					LogMessage($"所有關機方法都失敗: {ex2.Message}");
				}
			}
		}
	}
}