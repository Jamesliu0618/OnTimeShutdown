# Project Guidelines

## Architecture

Single-file Windows Forms application (`Program.cs`, ~1000+ lines) targeting `net8.0-windows`.

- **`ShutdownConfig`** – POCO class serialized to/from `shutdown_config.json` (same directory as executable). Fields: `Hour`, `Minute`, `Second`, `ForceShutdown`, `EnableAutoShutdown`.
- **`Program`** – Static entry point class. Responsibilities:
  - `Main`: initialize console (`AllocConsole`), load config, build system tray icon, hide console, start timer.
  - `InitializeSystemTrayIcon` / `UpdateTrayMenuItems`: manage `NotifyIcon` + `ContextMenuStrip`.
  - `StartShutdownTimer` / timer callback: polls every second, calls `ExecuteShutdown` when time matches.
  - `ExecuteShutdown` / `ExecuteDirectShutdown`: runs `shutdown` via `Process.Start`. Logs before/after.
  - Config helpers: `LoadOrCreateConfig`, `SaveConfig`, `ModifyShutdownTime`.
  - Log helpers: `InitializeLogFile`, `LogMessage` – writes to `shutdown_log.txt` beside executable.
- Config/log paths are derived from `Assembly.GetExecutingAssembly().Location`, not `Directory.GetCurrentDirectory()`.

## Code Style

- Language: C# with `<LangVersion>latest</LangVersion>`, `<Nullable>enable</Nullable>`, `<ImplicitUsings>enable</ImplicitUsings>`.
- Output type: `WinExe` with `<UseWindowsForms>true</UseWindowsForms>`.
- Nullable-aware throughout: use `?` for nullable fields (`config?`, `trayIcon?`, `shutdownTimer?`) and null-check before accessing.
- Shared mutable state (`config`) is guarded by `lock (configLock)`.
- Windows API functions are declared via `[DllImport]` P/Invoke inside `Program`.
- Comments may be written in Traditional Chinese (zh-TW); this is intentional.

## Build and Test

```bash
# 建構
dotnet build

# 執行
dotnet run

# 以最小化方式啟動
Showdown.exe --minimized

# 測試關機邏輯（不實際關機，延遲 5 秒後執行）
Showdown.exe --test-shutdown
```

Output: `bin/Debug/net8.0-windows/Showdown.exe`. The `shutdown_config.json` and `shutdown_log.txt` are placed beside the executable.

## Project Conventions

- Always resolve file paths relative to the executable directory using `ExecutablePath` (`Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)`), not the working directory.
- When modifying `config`, acquire `configLock` and call `SaveConfig(config)` afterward; then call `UpdateTrayMenuItems()` to keep UI in sync.
- Timer uses `System.Timers.Timer` (not `System.Windows.Forms.Timer`). Elapsed handler must be thread-safe.
- Shutdown is triggered via `shutdown /s /t <seconds>` (normal) or `shutdown /s /f /t <seconds>` (force). The attempt count (`shutdownAttempts`) is incremented on each call.
- After any config change from a menu item, always call `UpdateTrayMenuItems()` to refresh checked states.

## Integration Points

- **`AutoStartShutdownTool.bat`**: placed in `shell:startup` to launch `Showdown.exe --minimized` at Windows login. Updates required if output path changes.
- **`shutdown_config.json`**: JSON config file. Edited by the app or manually. Schema: `{"Hour":int,"Minute":int,"Second":int,"ForceShutdown":bool,"EnableAutoShutdown":bool}`.
- **`shutdown_log.txt`**: append-only log beside the executable. Written on startup, shutdown attempts, and errors.

## Security

- The app calls the Windows `shutdown` command. Avoid allowing unvalidated input to flow into process arguments.
- Admin privileges may be required for forced shutdown; the code checks `WindowsIdentity`/`WindowsPrincipal` for elevation where needed.
