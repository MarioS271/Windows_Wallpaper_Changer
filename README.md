# Windows Wallpaper Changer

A lightweight tray application that randomly cycles through wallpapers in a chosen folder on a configurable timer — a replacement for Windows' built-in Slideshow background, which is known to stop working after sleep, lock, or display changes.

This application is fully vibe-coded, mostly because I was lazy but also because I wanted to test Claude Code a little. :)

## Features
- Randomly selects wallpapers from a target folder (`.jpg`, `.jpeg`, `.png`, `.bmp`)
- Re-scans the folder on every cycle — add or remove images without restarting
- Configurable interval: 5 / 10 / 15 / 30 / 60 minutes
- Lives entirely in the system tray — no window, no taskbar entry, no console
- Left or right click on the tray icon opens the context menu
- **Next wallpaper** — skip immediately and reset the timer
- **Change folder** — native Windows folder picker, applies instantly
- **Change interval** — submenu with the current interval ticked
- **Exit** — clean shutdown
- Settings persist across restarts via a hidden INI file (`%USERPROFILE%\windows_wallpaper_changer_config.ini`)
- Single `.exe`, no installer, no runtime dependencies, ~1–2 MB memory footprint

## How it works

On startup the program registers a hidden message-only window (`HWND_MESSAGE`) — this gives it a message loop without appearing in the taskbar or creating a visible window. It then adds a tray icon via `Shell_NotifyIcon` and starts a `WM_TIMER` at the chosen interval.

Each time the timer fires (or "Next wallpaper" is selected), `FindFirstFile` / `FindNextFile` enumerate all supported images in the configured folder. A random image — different from the previous one — is passed to `SystemParametersInfoW(SPI_SETDESKWALLPAPER, ...)`, which updates the desktop background. Settings are written to and read from a standard INI file via `GetPrivateProfileString` / `WritePrivateProfileString`. The folder picker uses `SHBrowseForFolder` (COM-based), so COM is initialised on startup with `CoInitializeEx`.

## Building

Requires **CMake ≥ 4.1** and a Windows C compiler (MSVC or MinGW-w64).

```
cmake -B build -DCMAKE_BUILD_TYPE=Release
cmake --build build --config Release
```

The resulting `Windows_Wallpaper_Changer.exe` is fully self-contained and can be placed anywhere. To run it automatically on login, add a shortcut to `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`.

<br><hr><br>

This project is licensed under the **GNU Affero General Public License v3.0**. See the [LICENSE](LICENSE) file for details.
