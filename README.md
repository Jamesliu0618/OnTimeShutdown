# Auto Shutdown Tool

This is a simple auto shutdown tool that allows you to set shutdown time and supports shutdown after system updates.

## Features

1. Set shutdown time (hour, minute, second)
2. Support force shutdown option
3. Handle system updates by ensuring shutdown after updates
4. Configuration stored in an easy-to-edit text file (JSON format)
5. Support for auto-start on Windows boot (minimized)

## How to Use

1. After running the program, you will see the main menu
2. Choose one of the following options:
   - 1: Show current configuration
   - 2: Modify shutdown time
   - 3: Toggle force shutdown
   - 4: Shutdown now
   - 5: Exit program

## Auto-Start on Windows Boot

To make the application start automatically when Windows boots:

1. Build the application using `dotnet build`
2. Copy the `AutoStartShutdownTool.bat` script to the Windows Startup folder:
   - Press `Win + R`, type `shell:startup` and press Enter
   - Copy the `AutoStartShutdownTool.bat` script to this folder
   - The application will now start minimized on Windows boot

You can also start the application minimized by running:
```
Showdown.exe --minimized
```
or
```
Showdown.exe -m
```

## Configuration File

The program will automatically create a `shutdown_config.json` file in the same directory. You can directly edit this file to modify settings:

```json
{
  "Hour": 17,
  "Minute": 30,
  "Second": 0,
  "ForceShutdown": false
}
```

## System Requirements

- Windows operating system
- .NET 6.0 or higher

## Build and Run

```bash
# Build the program
dotnet build

# Run the program
dotnet run

# Run minimized
dotnet run -- --minimized
``` 