using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text.Json;

// 添加權限相關命名空間
// 添加反射命名空間，用於獲取執行檔路徑

namespace Showdown
{
	// Shutdown Configuration Class
	internal class ShutdownConfig
	{
		public int  Hour               { get; set; } = 17;
		public int  Minute             { get; set; } = 30;
		public int  Second             { get; set; }
		public bool ForceShutdown      { get; set; }
		public bool EnableAutoShutdown { get; set; } = true;

		public override string ToString()
		{
			return $"設定關機時間為: {Hour}:{Minute}:{Second}" + $"{(ForceShutdown ? " (強制關機)" : "")}" + $", 自動關機: {(EnableAutoShutdown ? "已啟用" : "已停用")}";
		}
	}

	internal class Program
	{
		// Show window commands
		private const int SW_HIDE     = 0;
		private const int SW_MINIMIZE = 6;
		private const int SW_SHOW     = 5;

		// 修改設定檔和日誌檔案路徑，使用執行檔所在目錄
		private static readonly string               ExecutablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "";
		private static readonly string               ConfigPath     = Path.Combine(ExecutablePath, "shutdown_config.json");
		private static readonly string               TrayIconPath   = Path.Combine(ExecutablePath, "Resources", "tray.ico");
		private static          System.Timers.Timer? shutdownTimer; // 允許為 null
		private static          NotifyIcon?          trayIcon;      // 允許為 null
		private static          Icon?                trayIconResource;
		private static          ShutdownConfig?      config;        // 允許為 null
		private static readonly object               configLock = new object();
		private static          int                  shutdownAttempts; // 關機嘗試次數
		private static          DateTime?            postponedShutdownTime;
		private static          int                  shutdownPromptActive;

		// 每日一個日誌檔案
		private static string LogPath { get => Path.Combine(ExecutablePath, $"shutdown_log_{DateTime.Now:yyyy-MM-dd}.txt"); }

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

			// 初始化日誌檔案
			InitializeLogFile();
			LogMessage("=================================");
			LogMessage("自動關機工具 - v1.2");
			LogMessage("=================================");

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
				LogMessage("測試關機功能...");
				LogMessage("將在 5 秒後嘗試執行關機命令");
				Thread.Sleep(5000);
				ExecuteDirectShutdown();
				return;
			}

			// Immediately check time and display next shutdown time
			LogMessage("檢查目前時間和關機設定...");

			if(config != null)
			{
				ShowNextShutdownTime(config);
			}

			// Start shutdown timer
			if(config != null)
			{
				//記錄訊息
				StartShutdownTimer(config);
			}

			// Run application
			Application.Run();
		}

		private static void InitializeSystemTrayIcon()
		{
			trayIcon = new NotifyIcon
					   {
						   Icon    = LoadTrayIcon(),
						   Text    = "自動關機工具",
						   Visible = true
					   };

			// Create context menu
			ContextMenuStrip menu = new ContextMenuStrip();

			menu.Items.Add("顯示主控台", null, OnShowConsole);
			menu.Items.Add("隱藏主控台", null, OnHideConsole);
			menu.Items.Add("-"); // Separator

			menu.Items.Add("顯示設定",   null, OnShowConfig);
			menu.Items.Add("修改關機時間", null, OnModifyTime);

			if(config != null)
			{
				ToolStripMenuItem autoShutdownItem = new ToolStripMenuItem("啟用自動關機");
				autoShutdownItem.Checked =  config.EnableAutoShutdown;
				autoShutdownItem.Click   += OnToggleAutoShutdown;
				menu.Items.Add(autoShutdownItem);

				ToolStripMenuItem forceShutdownItem = new ToolStripMenuItem("強制關機");
				forceShutdownItem.Checked =  config.ForceShutdown;
				forceShutdownItem.Click   += OnToggleForceShutdown;
				menu.Items.Add(forceShutdownItem);
			}

			menu.Items.Add("-"); // Separator
			menu.Items.Add("立即關機", null, OnShutdownNow);
			menu.Items.Add("結束",   null, OnExit);

			trayIcon.ContextMenuStrip = menu;

			// Double-click to toggle console visibility
			trayIcon.DoubleClick += OnTrayIconDoubleClick;
		}

		private static Icon LoadTrayIcon()
		{
			if(!File.Exists(TrayIconPath))
			{
				LogMessage($"找不到系統匣圖示，改用預設圖示: {TrayIconPath}");
				return SystemIcons.Application;
			}

			try
			{
				trayIconResource = new Icon(TrayIconPath);
				return trayIconResource;
			}
			catch(IOException ex)
			{
				LogMessage($"載入系統匣圖示失敗，改用預設圖示: {ex.Message}");
			}
			catch(UnauthorizedAccessException ex)
			{
				LogMessage($"載入系統匣圖示失敗，改用預設圖示: {ex.Message}");
			}
			catch(ArgumentException ex)
			{
				LogMessage($"載入系統匣圖示失敗，改用預設圖示: {ex.Message}");
			}

			return SystemIcons.Application;
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

			if(config != null)
			{
				LogMessage($"目前設定: {config}");
				ShowNextShutdownTime(config);
			}
			else
			{
				LogMessage("設定檔尚未載入。");
			}
		}

		private static void OnModifyTime(object? sender, EventArgs e)
		{
			ShowConsoleWindow();

			if(config != null)
			{
				ModifyShutdownTime(config);
				SaveConfig(config);

				// 更新菜單項的勾選狀態
				UpdateTrayMenuItems();
			}
			else
			{
				LogMessage("設定檔尚未載入。");
			}
		}

		private static void OnToggleAutoShutdown(object? sender, EventArgs e)
		{
			lock(configLock)
			{
				if(config != null)
				{
					config.EnableAutoShutdown = !config.EnableAutoShutdown;
					LogMessage($"自動關機設定為: {(config.EnableAutoShutdown ? "已啟用" : "已停用")}");
					SaveConfig(config);

					// Update menu item checked state
					if(sender is ToolStripMenuItem item)
					{
						item.Checked = config.EnableAutoShutdown;
					}
				}
			}
		}

		private static void OnToggleForceShutdown(object? sender, EventArgs e)
		{
			lock(configLock)
			{
				if(config != null)
				{
					config.ForceShutdown = !config.ForceShutdown;
					LogMessage($"強制關機設定為: {(config.ForceShutdown ? "已啟用" : "已停用")}");
					SaveConfig(config);

					// Update menu item checked state
					if(sender is ToolStripMenuItem item)
					{
						item.Checked = config.ForceShutdown;
					}
				}
			}
		}

		private static void OnShutdownNow(object? sender, EventArgs e)
		{
			if(config != null)
			{
				ExecuteShutdown(config);
			}
		}

		private static void OnExit(object? sender, EventArgs e)
		{
			if(trayIcon != null)
			{
				trayIcon.Visible = false;
				trayIcon.Dispose();
			}

			trayIconResource?.Dispose();
			trayIconResource = null;

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
			if((trayIcon?.ContextMenuStrip != null) && (config != null))
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
				LogMessage($"嘗試從 {ConfigPath} 讀取設定檔");

				if(File.Exists(ConfigPath))
				{
					string          json         = File.ReadAllText(ConfigPath);
					ShutdownConfig? loadedConfig = JsonSerializer.Deserialize<ShutdownConfig>(json);

					if(loadedConfig != null)
					{
						LogMessage($"設定檔載入成功: {loadedConfig}");
						return loadedConfig;
					}
					LogMessage("反序列化設定檔為 null，建立預設設定");
				}
				else
				{
					LogMessage($"設定檔不存在: {ConfigPath}，將建立預設設定");
				}
			}
			catch(Exception ex)
			{
				LogMessage($"設定檔讀取錯誤: {ex.Message}");
			}

			// Create default configuration
			ShutdownConfig defaultConfig = new ShutdownConfig();
			SaveConfig(defaultConfig);
			LogMessage($"已建立預設設定: {defaultConfig}");
			return defaultConfig;
		}

		private static void SaveConfig(ShutdownConfig config)
		{
			try
			{
				JsonSerializerOptions options = new JsonSerializerOptions
												{
													WriteIndented = true
												};
				string json = JsonSerializer.Serialize(config, options);
				File.WriteAllText(ConfigPath, json);
				LogMessage("設定已保存");

				// 顯示更新後的關機時間
				ShowNextShutdownTime(config);

				// 用更新的設定重新啟動關機計時器
				RestartShutdownTimer(config);

				// Update tray icon tooltip
				UpdateTrayIconTooltip();
			}
			catch(Exception ex)
			{
				LogMessage($"設定檔保存錯誤: {ex.Message}");
			}
		}

		private static void UpdateTrayIconTooltip()
		{
			if((trayIcon != null) && (config != null))
			{
				DateTime shutdownTime = GetNextShutdownTime(config);
				string   status       = config.EnableAutoShutdown ? $"Next shutdown: {shutdownTime:HH:mm:ss}" : "Auto shutdown disabled";

				trayIcon.Text = $"Auto Shutdown Tool - {status}";
			}
		}

		private static DateTime GetNextShutdownTime(ShutdownConfig config)
		{
			// 確保 config 不為 null，若為 null 則使用預設值
			if(config == null)
			{
				LogMessage("警告: config 為 null，使用預設值");
				config = new ShutdownConfig();
			}

			DateTime now          = DateTime.Now;

			lock(configLock)
			{
				if(postponedShutdownTime.HasValue)
				{
					if(now <= postponedShutdownTime.Value)
					{
						return postponedShutdownTime.Value;
					}

					postponedShutdownTime = null;
				}
			}

			DateTime shutdownTime = new DateTime(now.Year, now.Month, now.Day, config.Hour, config.Minute, config.Second);

			// 如果關機時間已經過了，設為明天
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
				LogMessage("自動關機功能目前已停用");
				return;
			}

			DateTime now          = DateTime.Now;
			DateTime shutdownTime = GetNextShutdownTime(config);

			// 用關機時間更新系統托盤圖示的提示文字
			UpdateTrayIconTooltip();

			// If shutdown time is tomorrow
			if(shutdownTime.Date > now.Date)
			{
				LogMessage($"下次預定關機: 明天 {shutdownTime.ToShortTimeString()}");
			}
			else
			{
				TimeSpan timeLeft = shutdownTime - now;

				lock(configLock)
				{
					if(postponedShutdownTime.HasValue && (shutdownTime == postponedShutdownTime.Value))
					{
						LogMessage($"本次關機已順延至今天 {shutdownTime:HH:mm:ss} (剩餘 {timeLeft.Hours}小時 {timeLeft.Minutes}分鐘 {timeLeft.Seconds}秒)");
						return;
					}
				}

				LogMessage($"下次預定關機: 今天 {shutdownTime.ToShortTimeString()} (剩餘 {timeLeft.Hours}小時 {timeLeft.Minutes}分鐘 {timeLeft.Seconds}秒)");
			}
		}

		private static void PostponeShutdownByFiveMinutes()
		{
			lock(configLock)
			{
				postponedShutdownTime = DateTime.Now.AddMinutes(5);
				LogMessage($"使用者選擇順延關機，本次關機延後至 {postponedShutdownTime.Value:yyyy-MM-dd HH:mm:ss}");
			}

			if(config != null)
			{
				ShowNextShutdownTime(config);
			}

			UpdateTrayIconTooltip();
		}

		private static bool ShowShutdownPrompt(DateTime scheduledShutdownTime)
		{
			bool shouldPostpone = false;
			Exception? promptException = null;

			Thread promptThread = new Thread(() =>
			{
				try
				{
					using Form promptForm = new Form
						  {
							  Text            = "自動關機提醒",
							  Width           = 420,
							  Height          = 220,
							  FormBorderStyle = FormBorderStyle.FixedDialog,
							  MaximizeBox     = false,
							  MinimizeBox     = false,
							  StartPosition   = FormStartPosition.CenterScreen,
							  TopMost         = true,
							  ShowInTaskbar   = true
						  };

					using Label titleLabel = new Label
						  {
							  Left      = 20,
							  Top       = 20,
							  Width     = 360,
							  Height    = 55,
							  Text      = $"現在到達預定關機時間 {scheduledShutdownTime:HH:mm:ss}。\r\n是否要順延 5 分鐘？",
							  Font      = new Font(SystemFonts.DefaultFont.FontFamily, 11, FontStyle.Bold),
							  TextAlign = ContentAlignment.MiddleLeft
						  };

					using Label countdownLabel = new Label
						  {
							  Left      = 20,
							  Top       = 85,
							  Width     = 360,
							  Height    = 30,
							  TextAlign = ContentAlignment.MiddleLeft
						  };

					using Button postponeButton = new Button
						  {
							  Text         = "順延 5 分鐘",
							  Width        = 130,
							  Height       = 35,
							  Left         = 70,
							  Top          = 130,
							  DialogResult = DialogResult.Yes
						  };

					using Button shutdownButton = new Button
						  {
							  Text         = "照常關機",
							  Width        = 130,
							  Height       = 35,
							  Left         = 220,
							  Top          = 130,
							  DialogResult = DialogResult.No
						  };

					int remainingSeconds = 30;

					void UpdateCountdownText()
					{
						countdownLabel.Text = $"若 {remainingSeconds} 秒內未選擇，系統將照常執行關機。";
					}

					using System.Windows.Forms.Timer countdownTimer = new System.Windows.Forms.Timer
						  {
							  Interval = 1000
						  };

					countdownTimer.Tick += (_, _) =>
					{
						remainingSeconds--;

						if(remainingSeconds <= 0)
						{
							countdownTimer.Stop();
							promptForm.DialogResult = DialogResult.No;
							promptForm.Close();
							return;
						}

						UpdateCountdownText();
					};

					promptForm.Controls.Add(titleLabel);
					promptForm.Controls.Add(countdownLabel);
					promptForm.Controls.Add(postponeButton);
					promptForm.Controls.Add(shutdownButton);
					promptForm.AcceptButton = postponeButton;
					promptForm.CancelButton = shutdownButton;
					promptForm.Shown += (_, _) =>
					{
						UpdateCountdownText();
						countdownTimer.Start();
					};
					promptForm.FormClosed += (_, _) => countdownTimer.Stop();

					shouldPostpone = promptForm.ShowDialog() == DialogResult.Yes;
				}
				catch(Exception ex)
				{
					promptException = ex;
				}
			});

			promptThread.SetApartmentState(ApartmentState.STA);
			promptThread.Start();
			promptThread.Join();

			if(promptException != null)
			{
				throw new InvalidOperationException("顯示關機提醒視窗失敗。", promptException);
			}

			return shouldPostpone;
		}

		private static void HandleShutdownDue(ShutdownConfig config, DateTime currentTime, DateTime scheduledShutdownTime)
		{
			if(Interlocked.CompareExchange(ref shutdownPromptActive, 1, 0) != 0)
			{
				return;
			}

			try
			{
				LogMessage("============================================================");
				LogMessage($"預定關機時間已到達 (目前時間: {currentTime:yyyy-MM-dd HH:mm:ss.fff})");
				LogMessage($"預定關機時間: {scheduledShutdownTime:yyyy-MM-dd HH:mm:ss}");
				LogMessage("顯示 30 秒確認視窗，允許使用者順延 5 分鐘");
				LogMessage("============================================================");

				if(shutdownTimer != null)
				{
					shutdownTimer.Stop();
				}

				if(trayIcon != null)
				{
					trayIcon.ShowBalloonTip(5000, "自動關機提醒", "30 秒內可選擇順延 5 分鐘，否則將照常關機。", ToolTipIcon.Warning);
				}

				bool shouldPostpone = ShowShutdownPrompt(scheduledShutdownTime);

				if(shouldPostpone)
				{
					PostponeShutdownByFiveMinutes();

					if(shutdownTimer != null)
					{
						shutdownTimer.Start();
					}

					return;
				}

				LogMessage("使用者未順延關機，將立即執行關機命令。");

				if(trayIcon != null)
				{
					trayIcon.ShowBalloonTip(3000, "電腦關機", "電腦正在關機中...", ToolTipIcon.Info);
				}

				config.ForceShutdown = true;
				LogMessage("準備呼叫 ExecuteDirectShutdown() 方法");
				ExecuteDirectShutdown();
				LogMessage("已呼叫 ExecuteDirectShutdown() 方法");
			}
			finally
			{
				Interlocked.Exchange(ref shutdownPromptActive, 0);
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

			// 建立計時器，每1秒檢查一次，提高準確性
			shutdownTimer = new System.Timers.Timer(1000); // 1秒

			shutdownTimer.Elapsed += (sender, e) =>
			{
				try
				{
					CheckShutdownTime(config);
				}
				catch(Exception ex)
				{
					LogMessage($"[計時器] Elapsed 例外: {ex}");
				}
			};
			shutdownTimer.AutoReset = true;
			shutdownTimer.Start();

			LogMessage("關機計時器已啟動 - 每秒檢查一次");

			// Do initial check immediately
			CheckShutdownTime(config);
		}

		private static void CheckShutdownTime(ShutdownConfig config)
		{
			try
			{
				if(Volatile.Read(ref shutdownPromptActive) == 1)
				{
					return;
				}

				// 取得一次當前時間以及關機時間，避免重複計算
				DateTime currentTime           = DateTime.Now;
				DateTime scheduledShutdownTime = GetNextShutdownTime(config);
				TimeSpan remainingTime         = scheduledShutdownTime - currentTime;

				// 根據時間點決定日誌記錄等級，避免產生過大的日誌檔案
				// 只在以下情況記錄詳細日誌：
				// 1. 接近關機時間（5分鐘內）
				// 2. 每10分鐘的整點記錄一次
				// 3. 每小時的整點記錄一次
				bool isNearShutdown = remainingTime.TotalMinutes <= 5; // 接近關機時間時記錄詳細日誌

				// 使用目前的系統時間來判斷是否為整點時間，而不是用剩餘時間
				bool isHourMark      = (currentTime.Minute      == 0) && (currentTime.Second <= 1); // 每小時整點記錄
				bool isTenMinuteMark = (currentTime.Minute % 10 == 0) && (currentTime.Second <= 1); // 每10分鐘整點記錄

				bool isDetailedLoggingNeeded = isNearShutdown || isHourMark || isTenMinuteMark;

				if(isDetailedLoggingNeeded)
				{
					LogMessage("===== 檢查關機時間 =====");
					LogMessage($"目前系統時間: {currentTime:yyyy-MM-dd HH:mm:ss.fff} ({TimeZoneInfo.Local.DisplayName})");
					LogMessage($"預定關機時間: {scheduledShutdownTime:yyyy-MM-dd HH:mm:ss}");
					LogMessage($"距離關機剩餘秒數: {remainingTime.TotalSeconds:F2}");
				}

				bool shouldHandleShutdownDue = false;

				lock(configLock)
				{
					if(!config.EnableAutoShutdown)
					{
						if(isDetailedLoggingNeeded)
						{
							LogMessage("自動關機功能已停用，跳過時間檢查。");
						}
						return;
					}

					// 在 lock 區塊中使用已計算的時間變數

					// 計算剩餘時間的各部分，供後續使用
					int remainHours   = (int)remainingTime.TotalHours;
					int remainMinutes = remainingTime.Minutes;
					int remainSeconds = remainingTime.Seconds;

					// 只在需要詳細記錄時才記錄時間訊息，避免過多輸出影響效能
					if(isDetailedLoggingNeeded || (remainingTime.TotalMinutes <= 10))
					{
						LogMessage($"[偵錯] 目前時間: {currentTime:HH:mm:ss}");
						LogMessage($"[偵錯] 預定關機時間: {scheduledShutdownTime:HH:mm:ss}");
						LogMessage($"[偵錯] 剩餘時間: {remainHours}小時 {remainMinutes}分鐘 {remainSeconds}秒 (共 {(int)remainingTime.TotalSeconds} 秒)");
					}

					// 重要：使用總秒數進行更精確的比較，避免小數點問題
					double totalSeconds = remainingTime.TotalSeconds;

					// 如果時間已到或已過（允許0.5秒誤差），立即執行關機
					if((totalSeconds <= 0.5) && (totalSeconds > -60)) // 允許0.5秒誤差，且不超過1分鐘的過期時間
					{
						shouldHandleShutdownDue = true;
					}

					// 如果時間低於10秒，準備關機
					if((remainingTime.TotalSeconds <= 10) && (remainingTime.TotalSeconds > 0))
					{
						LogMessage($"關機時間即將到達! 將在 {scheduledShutdownTime:HH:mm:ss} 執行關機 (剩餘 {remainSeconds} 秒)");

						// Show balloon tip
						if(trayIcon != null)
						{
							trayIcon.ShowBalloonTip(5000, "關機即將執行", $"電腦將在 {remainSeconds} 秒後關機。", ToolTipIcon.Warning);
						}

						// 如果時間很接近（2秒以內），執行關機
						if(remainingTime.TotalSeconds <= 2)
						{
							shouldHandleShutdownDue = true;
						}
					}

					// 如果在1分鐘內關機
					else if((remainingTime.TotalMinutes <= 1) && (remainingTime.TotalSeconds > 0))
					{
						LogMessage($"關機時間即將到達! 將在 {remainMinutes} 分鐘 {remainSeconds} 秒後執行關機");
					}
				}

				if(shouldHandleShutdownDue)
				{
					HandleShutdownDue(config, currentTime, scheduledShutdownTime);
				}
			}
			catch(Exception ex)
			{
				LogMessage($"檢查關機時間發生錯誤: {ex.Message}");
				LogMessage($"詳細錯誤: {ex}");
			}
		}

		// 初始化日誌檔案
		private static void InitializeLogFile()
		{
			try
			{
				string? logDirectory = Path.GetDirectoryName(LogPath);

				if(!string.IsNullOrEmpty(logDirectory) && !Directory.Exists(logDirectory))
				{
					Directory.CreateDirectory(logDirectory);
				}
				LogMessage("程式啟動，日誌系統初始化");
				LogMessage($"執行檔路徑: {ExecutablePath}");
				LogMessage($"設定檔路徑: {ConfigPath}");
				LogMessage($"日誌檔案路徑: {LogPath}");
			}
			catch(Exception ex)
			{
				try
				{
					string altPath = Path.Combine(Environment.CurrentDirectory, $"shutdown_log_{DateTime.Now:yyyy-MM-dd}_alt.txt");
					File.AppendAllText(altPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 日誌初始化失敗: {ex.Message}\r\n");
				}
				catch
				{
				}
			}
		}

		// 記錄訊息至每日日誌檔案
		private static void LogMessage(string message)
		{
			try
			{
				string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
				File.AppendAllText(LogPath, logEntry + "\r\n");
			}
			catch
			{
			}
		}

		private static void RestartShutdownTimer(ShutdownConfig config)
		{
			// Simply restart the timer with new config
			StartShutdownTimer(config);
		}

		private static void ModifyShutdownTime(ShutdownConfig config)
		{
			LogMessage("修改關機時間 (按 Enter 保持目前值)");

			Console.Write($"小時 (0-23) [{config.Hour}]: ");
			string hourInput = Console.ReadLine() ?? "";

			if(!string.IsNullOrWhiteSpace(hourInput) && int.TryParse(hourInput, out int hour) && (hour >= 0) && (hour <= 23))
			{
				config.Hour = hour;
			}

			Console.Write($"分鐘 (0-59) [{config.Minute}]: ");
			string minuteInput = Console.ReadLine() ?? "";

			if(!string.IsNullOrWhiteSpace(minuteInput) && int.TryParse(minuteInput, out int minute) && (minute >= 0) && (minute <= 59))
			{
				config.Minute = minute;
			}

			Console.Write($"秒 (0-59) [{config.Second}]: ");
			string secondInput = Console.ReadLine() ?? "";

			if(!string.IsNullOrWhiteSpace(secondInput) && int.TryParse(secondInput, out int second) && (second >= 0) && (second <= 59))
			{
				config.Second = second;
			}

			LogMessage($"新的關機時間已設定為: {config.Hour}:{config.Minute}:{config.Second}");
		}

		private static void ExecuteShutdown(ShutdownConfig config)
		{
			// 確保 config 不為 null
			if(config == null)
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

				// 使用管理員權限執行
				ProcessStartInfo processInfo = new ProcessStartInfo("cmd.exe", $"/c {command}")
											   {
												   CreateNoWindow  = true,
												   UseShellExecute = true,
												   Verb            = "runas"
											   };

				Process.Start(processInfo);
				LogMessage("關機命令已發送!");

				// 確保程序有足夠時間執行
				Thread.Sleep(3000);
			}
			catch(Exception ex)
			{
				LogMessage($"手動關機執行錯誤: {ex.Message}");

				// 嘗試使用替代方法
				try
				{
					LogMessage("嘗試使用直接方法關機...");
					ExecuteDirectShutdown();
				}
				catch(Exception ex2)
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

				LogMessage("開始執行關機程序");

				// 增加關機嘗試次數
				shutdownAttempts++;
				LogMessage($"第 {shutdownAttempts} 次關機嘗試");

				// 記錄系統當前狀態以便診斷
				LogMessage($"作業系統: {Environment.OSVersion}");
				LogMessage($"是否以管理員身分執行: {IsRunningAsAdministrator()}");
				LogMessage($"當前資料夾: {Environment.CurrentDirectory}");

				// 嘗試使用強制關機參數，確保強制關機
				string arguments = "/s /f /t 0"; // 始終使用強制關機參數
				LogMessage($"執行關機命令: shutdown {arguments}");

				// 執行關機命令
				try
				{
					// 使用 ProcessStartInfo 直接執行關機命令，使用系統管理員權限
					ProcessStartInfo psi = new ProcessStartInfo
										   {
											   FileName        = "shutdown",
											   Arguments       = arguments,
											   CreateNoWindow  = true,
											   UseShellExecute = true,   // true 允許提升權限
											   Verb            = "runas" // 以管理員權限執行
										   };

					LogMessage("嘗試啟動關機進程...");
					Process? process = Process.Start(psi);

					if(process != null)
					{
						LogMessage($"關機進程已啟動，進程ID: {process.Id}");
					}
					else
					{
						LogMessage("警告：關機進程啟動返回null");
					}

					LogMessage("等待關機執行...");
					Thread.Sleep(5000);
				}
				catch(Exception processEx)
				{
					LogMessage($"關機進程啟動失敗: {processEx.Message}");
					LogMessage("嘗試使用CMD執行關機...");

					try
					{
						ProcessStartInfo cmdPsi = new ProcessStartInfo
												  {
													  FileName        = "cmd.exe",
													  Arguments       = "/c shutdown /s /f /t 0",
													  CreateNoWindow  = true,
													  UseShellExecute = true,
													  Verb            = "runas"
												  };

						Process? cmdProcess = Process.Start(cmdPsi);

						if(cmdProcess != null)
						{
							LogMessage($"CMD關機進程已啟動，進程ID: {cmdProcess.Id}");
						}

						Thread.Sleep(5000);
					}
					catch(Exception cmdEx)
					{
						LogMessage($"CMD關機方法失敗: {cmdEx.Message}");
						LogMessage("嘗試直接執行shutdown命令...");

						try
						{
							Process.Start("shutdown", "-s -f -t 0");
							Thread.Sleep(5000);
						}
						catch(Exception finalEx)
						{
							LogMessage($"所有關機方法嘗試失敗: {finalEx.Message}");
						}
					}
				}
			}
			catch(Exception ex)
			{
				LogMessage($"嚴重錯誤在ExecuteDirectShutdown方法: {ex.Message}");

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
		private static extern bool InitiateSystemShutdownEx(string lpMachineName,
															string lpMessage,
															uint   dwTimeout,
															bool   bForceAppsClosed,
															bool   bRebootAfterShutdown,
															uint   dwReason);

		// 檢查是否以系統管理員身分運行
		private static bool IsRunningAsAdministrator()
		{
			try
			{
				WindowsIdentity  identity  = WindowsIdentity.GetCurrent();
				WindowsPrincipal principal = new WindowsPrincipal(identity);
				return principal.IsInRole(WindowsBuiltInRole.Administrator);
			}
			catch(Exception ex)
			{
				LogMessage($"檢查管理員權限時發生錯誤: {ex.Message}");
				return false;
			}
		}
	}
}
