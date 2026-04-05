#pragma once
#define WIN32_LEAN_AND_MEAN
#define UNICODE
#define _UNICODE
#include <windows.h>

/* Attempt to update all virtual desktops' wallpaper via the undocumented
 * IVirtualDesktopManagerInternal COM interface (Windows 11 Build 22621+).
 *
 * Returns TRUE  – COM call succeeded; all desktops updated live.
 * Returns FALSE – interface unavailable (old build, Explorer issue, etc.);
 *                 caller should fall back to registry writes + SPI. */
BOOL VDesktop_SetWallpaperAllDesktops(const wchar_t *path);
