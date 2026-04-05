# Project State

## Working

- **Tray icon** — lives in the system tray, no window or taskbar entry (HWND_MESSAGE window)
- **Context menu** — left/right click; Next wallpaper, Change folder, Change interval submenu (5/10/15/30/60 min, current ticked), Exit
- **Folder picker** — uses `IFileOpenDialog` (modern API, replaces broken `SHBrowseForFolderW`)
- **Wallpaper scanning** — `FindFirstFileW` re-scans the folder each cycle; supports `.jpg`, `.jpeg`, `.png`, `.bmp`
- **No-repeat tracking** — `g_current_wallpaper` (in RAM, not persisted) stores the full path of the current wallpaper so it is never picked again until another file is chosen; resets on folder change
- **Config persistence** — hidden INI at `%USERPROFILE%\windows_wallpaper_changer_config.ini`; saves/restores folder and interval across restarts
- **Interval timer** — `SetTimer`/`KillTimer`; changing interval or clicking Next resets the clock
- **Debug logging** — `--debug` flag opens a console window; `WriteConsoleW`-based, wide-string safe (`%ls`)

## Known Bugs

### Virtual desktop wallpaper — accepted limitation

**What works:** `ApplyWallpaperToAllDesktops` correctly writes the new path to every desktop's `HKCU\...\VirtualDesktops\Desktops\{guid}\Wallpaper` registry entry (confirmed via PowerShell), and `SPI_SETDESKWALLPAPER` refreshes the active desktop live.

**What doesn't work:** The other virtual desktops do not update on switch. Explorer caches each desktop's wallpaper bitmap in memory and does not re-read the registry on switch — the registry is where explorer *saves* the value, not where it reads from at runtime. The in-memory bitmap can only be updated via `IDesktopWallpaper::SetWallpaper`.

**`IDesktopWallpaper` exhausted:** Both known IIDs return `E_NOINTERFACE`, CLSID is confirmed in the registry, and there is no further public Win32 path to update inactive virtual desktop wallpapers. The undocumented `IVirtualDesktopManagerInternal` is the only remaining option and is not viable for a stable release.

**Current behaviour:** Wallpaper changes immediately on the active desktop and is written to all desktops' registry entries. If the user restarts explorer or logs off and back on, all desktops will show the last applied wallpaper. During a session, only the active desktop updates live.

## TODOs

- [ ] **Custom tray icon** — currently uses the generic `IDI_APPLICATION` icon. Needs a custom `.ico` file embedded as a resource (`.rc` file + `CMakeLists.txt` updated to compile it with `windres`), then loaded with `LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_ICON1))`.
