#define WIN32_LEAN_AND_MEAN
#define UNICODE
#define _UNICODE

#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <stdlib.h>
#include <string.h>
#include <time.h>
#include <wchar.h>
#include <wctype.h>
#include <stdarg.h>
#include "vdesktop.h"

#ifndef _countof
#define _countof(arr) (sizeof(arr) / sizeof((arr)[0]))
#endif

#define WM_TRAYICON       (WM_USER + 1)
#define IDT_WALLPAPER     1
#define IDM_NEXT          1001
#define IDM_FOLDER        1002
#define IDM_INTERVAL_5    1003
#define IDM_INTERVAL_10   1004
#define IDM_INTERVAL_15   1005
#define IDM_INTERVAL_30   1006
#define IDM_INTERVAL_60   1007
#define IDM_EXIT          1008
#define MAX_FILES         4096
#define IDI_ICON1         101

static HWND             g_hwnd;
static NOTIFYICONDATAW  g_nid;
static wchar_t          g_folder[MAX_PATH];
static int              g_interval_minutes = 10;
static wchar_t          g_files[MAX_FILES][MAX_PATH];
static int              g_file_count  = 0;
static wchar_t          g_current_wallpaper[MAX_PATH]; /* path of last applied wallpaper, "" = none */
static wchar_t          g_config_path[MAX_PATH];
static int              g_debug = 0;

/* ---------- debug log ---------- */

static void Log(const wchar_t *fmt, ...) {
    if (!g_debug) return;
    va_list args;
    va_start(args, fmt);
    wchar_t msg[2048];
    vswprintf(msg, _countof(msg), fmt, args);
    va_end(args);
    HANDLE hOut = GetStdHandle(STD_OUTPUT_HANDLE);
    DWORD written;
    WriteConsoleW(hOut, msg, (DWORD)wcslen(msg), &written, NULL);
    WriteConsoleW(hOut, L"\n", 1, &written, NULL);
}

/* ---------- config ---------- */

static void BuildConfigPath(void) {
    wchar_t home[MAX_PATH];
    if (SHGetFolderPathW(NULL, CSIDL_PROFILE, NULL, 0, home) != S_OK)
        wcscpy(home, L"C:\\Users\\Default");
    wcscpy(g_config_path, home);
    wcscat(g_config_path, L"\\windows_wallpaper_changer_config.ini");
}

static void LoadConfig(void) {
    wchar_t buf[16];
    GetPrivateProfileStringW(L"Settings", L"Folder", L"",
                             g_folder, _countof(g_folder), g_config_path);
    GetPrivateProfileStringW(L"Settings", L"IntervalMinutes", L"10",
                             buf, _countof(buf), g_config_path);
    g_interval_minutes = (int)wcstol(buf, NULL, 10);
    if (g_interval_minutes < 1) g_interval_minutes = 10;
    Log(L"LoadConfig: folder=\"%ls\" interval=%d", g_folder, g_interval_minutes);
}

static void SaveConfig(void) {
    wchar_t buf[16];
    WritePrivateProfileStringW(L"Settings", L"Folder", g_folder, g_config_path);
    swprintf(buf, _countof(buf), L"%d", g_interval_minutes);
    WritePrivateProfileStringW(L"Settings", L"IntervalMinutes", buf, g_config_path);
    SetFileAttributesW(g_config_path, FILE_ATTRIBUTE_HIDDEN);
}

/* ---------- wallpaper ---------- */

/* Apply wallpaper to all virtual desktops.
 *
 * Primary path (Win11 Build 22621+): calls VDesktop_SetWallpaperAllDesktops()
 * which uses IVirtualDesktopManagerInternal::UpdateWallpaperPathForAllDesktops
 * to update Explorer's in-memory bitmap for every desktop live.
 *
 * Fallback path (older builds / COM failure): writes to the registry for each
 * desktop GUID so the wallpaper is restored on next login, then calls
 * SPI_SETDESKWALLPAPER to refresh the active desktop immediately.
 *
 * Registry layout (Windows 11):
 *   HKCU\...\VirtualDesktops\VirtualDesktopIDs  REG_BINARY  (16 bytes per GUID)
 *   HKCU\...\VirtualDesktops\Desktops\{guid}\Wallpaper  REG_SZ */
static void ApplyWallpaperToAllDesktops(const wchar_t *path) {
    /* COM path: updates all desktops live (Win11 22621+). */
    if (VDesktop_SetWallpaperAllDesktops(path)) {
        Log(L"ApplyWallpaper: COM path succeeded (all desktops updated live)");
        /* Belt-and-suspenders: also trigger the standard SPI refresh so the
         * active desktop picks up the change even if COM only queued it. */
        SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, (void *)path,
                              SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
        return;
    }

    Log(L"ApplyWallpaper: COM unavailable, using registry fallback");

    /* Registry fallback: persist wallpaper path for each virtual desktop GUID
     * so the correct wallpaper is shown after Explorer restarts / user logs in. */
    static const wchar_t VD_KEY[] =
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VirtualDesktops";

    HKEY hVD = NULL;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, VD_KEY, 0, KEY_READ, &hVD) == ERROR_SUCCESS) {
        BYTE  blob[16 * 64]; /* up to 64 virtual desktops */
        DWORD blobSize = sizeof(blob);
        DWORD type = 0;

        if (RegQueryValueExW(hVD, L"VirtualDesktopIDs", NULL,
                             &type, blob, &blobSize) == ERROR_SUCCESS
            && type == REG_BINARY && blobSize >= 16)
        {
            DWORD n = blobSize / 16;
            Log(L"ApplyWallpaper: %lu virtual desktop(s) found", (unsigned long)n);

            for (DWORD i = 0; i < n; i++) {
                /* Each 16-byte GUID: Data1(4 LE) Data2(2 LE) Data3(2 LE) Data4(8 BE) */
                BYTE *g = blob + i * 16;
                DWORD d1 = (DWORD)g[0] | ((DWORD)g[1]<<8) | ((DWORD)g[2]<<16) | ((DWORD)g[3]<<24);
                WORD  d2 = (WORD)g[4]  | ((WORD)g[5]<<8);
                WORD  d3 = (WORD)g[6]  | ((WORD)g[7]<<8);

                wchar_t guidStr[64];
                swprintf(guidStr, _countof(guidStr),
                         L"{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
                         (unsigned)d1, (unsigned)d2, (unsigned)d3,
                         g[8],g[9],g[10],g[11],g[12],g[13],g[14],g[15]);

                wchar_t subkey[256];
                wcscpy(subkey, VD_KEY);
                wcscat(subkey, L"\\Desktops\\");
                wcscat(subkey, guidStr);

                HKEY hDesk = NULL;
                if (RegCreateKeyExW(HKEY_CURRENT_USER, subkey, 0, NULL,
                                    REG_OPTION_NON_VOLATILE, KEY_SET_VALUE,
                                    NULL, &hDesk, NULL) == ERROR_SUCCESS) {
                    DWORD len = (DWORD)((wcslen(path) + 1) * sizeof(wchar_t));
                    RegSetValueExW(hDesk, L"Wallpaper", 0, REG_SZ,
                                   (const BYTE *)path, len);
                    RegCloseKey(hDesk);
                    Log(L"ApplyWallpaper: wrote desktop %ls", guidStr);
                }
            }
        }
        RegCloseKey(hVD);
    }

    /* Refresh the active desktop immediately */
    SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, (void *)path,
                          SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
}

static void ScanFolder(void) {
    g_file_count = 0;
    if (g_folder[0] == L'\0') {
        Log(L"ScanFolder: folder is empty, skipping");
        return;
    }

    wchar_t pattern[MAX_PATH];
    wcscpy(pattern, g_folder);
    wcscat(pattern, L"\\*");

    Log(L"ScanFolder: scanning \"%ls\"", pattern);
    WIN32_FIND_DATAW fd;
    HANDLE h = FindFirstFileW(pattern, &fd);
    if (h == INVALID_HANDLE_VALUE) {
        Log(L"ScanFolder: FindFirstFileW failed (error %lu)", GetLastError());
        return;
    }

    do {
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
        if (g_file_count >= MAX_FILES) break;

        wchar_t *dot = wcsrchr(fd.cFileName, L'.');
        if (!dot) continue;

        wchar_t ext[16];
        wcsncpy(ext, dot, _countof(ext) - 1);
        ext[_countof(ext) - 1] = L'\0';
        for (int i = 0; ext[i]; i++)
            ext[i] = (wchar_t)towlower(ext[i]);

        if (wcscmp(ext, L".jpg")  == 0 ||
            wcscmp(ext, L".jpeg") == 0 ||
            wcscmp(ext, L".png")  == 0 ||
            wcscmp(ext, L".bmp")  == 0)
        {
            wcscpy(g_files[g_file_count], g_folder);
            wcscat(g_files[g_file_count], L"\\");
            wcscat(g_files[g_file_count], fd.cFileName);
            g_file_count++;
        }
    } while (FindNextFileW(h, &fd));

    FindClose(h);
    Log(L"ScanFolder: found %d image(s)", g_file_count);
}

static void SetRandomWallpaper(void) {
    ScanFolder();
    if (g_file_count == 0) {
        Log(L"SetRandomWallpaper: no images found, aborting");
        return;
    }

    int idx;
    if (g_file_count == 1) {
        idx = 0;
    } else {
        do {
            idx = rand() % g_file_count;
        } while (wcscmp(g_files[idx], g_current_wallpaper) == 0);
    }
    wcscpy(g_current_wallpaper, g_files[idx]);

    Log(L"SetRandomWallpaper: applying \"%ls\"", g_files[idx]);
    ApplyWallpaperToAllDesktops(g_files[idx]);
}

/* ---------- tray icon ---------- */

static void UpdateTrayTooltip(void) {
    swprintf(g_nid.szTip, _countof(g_nid.szTip),
             L"Wallpaper Changer \u2014 Every %d min", g_interval_minutes);
    Shell_NotifyIconW(NIM_MODIFY, &g_nid);
}

static void AddTrayIcon(void) {
    memset(&g_nid, 0, sizeof(g_nid));
    g_nid.cbSize           = sizeof(g_nid);
    g_nid.hWnd             = g_hwnd;
    g_nid.uID              = 1;
    g_nid.uFlags           = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon            = LoadIconW(GetModuleHandleW(NULL), MAKEINTRESOURCEW(IDI_ICON1));
    if (!g_nid.hIcon) g_nid.hIcon = LoadIconW(NULL, IDI_APPLICATION);
    swprintf(g_nid.szTip, _countof(g_nid.szTip),
             L"Wallpaper Changer \u2014 Every %d min", g_interval_minutes);
    Shell_NotifyIconW(NIM_ADD, &g_nid);
}

/* ---------- context menu ---------- */

static void ShowContextMenu(HWND hwnd) {
    POINT pt;
    GetCursorPos(&pt);

    HMENU hSubMenu = CreatePopupMenu();
    AppendMenuW(hSubMenu, MF_STRING | (g_interval_minutes ==  5 ? MF_CHECKED : 0), IDM_INTERVAL_5,  L"5 minutes");
    AppendMenuW(hSubMenu, MF_STRING | (g_interval_minutes == 10 ? MF_CHECKED : 0), IDM_INTERVAL_10, L"10 minutes");
    AppendMenuW(hSubMenu, MF_STRING | (g_interval_minutes == 15 ? MF_CHECKED : 0), IDM_INTERVAL_15, L"15 minutes");
    AppendMenuW(hSubMenu, MF_STRING | (g_interval_minutes == 30 ? MF_CHECKED : 0), IDM_INTERVAL_30, L"30 minutes");
    AppendMenuW(hSubMenu, MF_STRING | (g_interval_minutes == 60 ? MF_CHECKED : 0), IDM_INTERVAL_60, L"60 minutes");

    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING,            IDM_NEXT,              L"Next wallpaper");
    AppendMenuW(hMenu, MF_STRING,            IDM_FOLDER,            L"Change folder...");
    AppendMenuW(hMenu, MF_POPUP, (UINT_PTR)hSubMenu,                L"Change interval");
    AppendMenuW(hMenu, MF_SEPARATOR,         0,                     NULL);
    AppendMenuW(hMenu, MF_STRING,            IDM_EXIT,              L"Exit");

    // Required so the menu dismisses when clicking elsewhere
    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_BOTTOMALIGN | TPM_RIGHTALIGN,
                   pt.x, pt.y, 0, hwnd, NULL);
    PostMessageW(hwnd, WM_NULL, 0, 0);

    DestroyMenu(hMenu); // also destroys hSubMenu
}

/* ---------- actions ---------- */

static void ChangeFolder(HWND hwnd) {
    // IFileOpenDialog is the modern replacement for SHBrowseForFolderW and
    // correctly returns full filesystem paths on Windows 10/11.
    IFileOpenDialog *pDlg = NULL;
    HRESULT hr = CoCreateInstance(&CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER,
                                  &IID_IFileOpenDialog, (void **)&pDlg);
    Log(L"ChangeFolder: CoCreateInstance hr=0x%08lX", (unsigned long)hr);
    if (FAILED(hr)) return;

    DWORD opts = 0;
    pDlg->lpVtbl->GetOptions(pDlg, &opts);
    pDlg->lpVtbl->SetOptions(pDlg, opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
    pDlg->lpVtbl->SetTitle(pDlg, L"Select wallpaper folder");

    hr = pDlg->lpVtbl->Show(pDlg, NULL);
    Log(L"ChangeFolder: Show hr=0x%08lX", (unsigned long)hr);
    if (SUCCEEDED(hr)) {
        IShellItem *pItem = NULL;
        if (SUCCEEDED(pDlg->lpVtbl->GetResult(pDlg, &pItem))) {
            PWSTR pszPath = NULL;
            hr = pItem->lpVtbl->GetDisplayName(pItem, SIGDN_FILESYSPATH, &pszPath);
            Log(L"ChangeFolder: GetDisplayName hr=0x%08lX len=%zu path=\"%ls\"",
                (unsigned long)hr, pszPath ? wcslen(pszPath) : 0, pszPath ? pszPath : L"<null>");
            if (SUCCEEDED(hr) && pszPath) {
                wcsncpy(g_folder, pszPath, MAX_PATH - 1);
                g_folder[MAX_PATH - 1] = L'\0';
                Log(L"ChangeFolder: g_folder len=%zu val=\"%ls\"", wcslen(g_folder), g_folder);
                CoTaskMemFree(pszPath);
                SaveConfig();
                g_current_wallpaper[0] = L'\0';
                SetRandomWallpaper();
                KillTimer(hwnd, IDT_WALLPAPER);
                SetTimer(hwnd, IDT_WALLPAPER, (UINT)g_interval_minutes * 60u * 1000u, NULL);
            }
            pItem->lpVtbl->Release(pItem);
        }
    }
    pDlg->lpVtbl->Release(pDlg);
}

static void SetInterval(HWND hwnd, int minutes) {
    g_interval_minutes = minutes;
    SaveConfig();
    KillTimer(hwnd, IDT_WALLPAPER);
    SetTimer(hwnd, IDT_WALLPAPER, (UINT)minutes * 60u * 1000u, NULL);
    UpdateTrayTooltip();
}

/* ---------- window procedure ---------- */

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_CREATE:
            g_hwnd = hwnd;
            srand((unsigned)time(NULL));
            BuildConfigPath();
            LoadConfig();
            AddTrayIcon();
            SetTimer(hwnd, IDT_WALLPAPER, (UINT)g_interval_minutes * 60u * 1000u, NULL);
            if (g_folder[0] != L'\0')
                SetRandomWallpaper();
            return 0;

        case WM_TRAYICON:
            if (lp == WM_LBUTTONUP || lp == WM_RBUTTONUP)
                ShowContextMenu(hwnd);
            return 0;

        case WM_TIMER:
            if (wp == IDT_WALLPAPER)
                SetRandomWallpaper();
            return 0;

        case WM_COMMAND:
            switch (LOWORD(wp)) {
                case IDM_NEXT:
                    SetRandomWallpaper();
                    KillTimer(hwnd, IDT_WALLPAPER);
                    SetTimer(hwnd, IDT_WALLPAPER, (UINT)g_interval_minutes * 60u * 1000u, NULL);
                    break;
                case IDM_FOLDER:      ChangeFolder(hwnd);    break;
                case IDM_INTERVAL_5:  SetInterval(hwnd,  5); break;
                case IDM_INTERVAL_10: SetInterval(hwnd, 10); break;
                case IDM_INTERVAL_15: SetInterval(hwnd, 15); break;
                case IDM_INTERVAL_30: SetInterval(hwnd, 30); break;
                case IDM_INTERVAL_60: SetInterval(hwnd, 60); break;
                case IDM_EXIT:        DestroyWindow(hwnd);   break;
            }
            return 0;

        case WM_DESTROY:
            Shell_NotifyIconW(NIM_DELETE, &g_nid);
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

/* ---------- entry point ---------- */

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                   LPSTR lpCmdLine, int nCmdShow) {
    (void)hPrevInstance;
    (void)lpCmdLine;
    (void)nCmdShow;

    if (wcsstr(GetCommandLineW(), L"--debug")) {
        g_debug = 1;
        AllocConsole();
    }

    HANDLE hMutex = CreateMutexW(NULL, TRUE, L"WallpaperChangerSingleInstance");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        CloseHandle(hMutex);
        return 0;
    }

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    WNDCLASSEXW wc   = {0};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInstance;
    wc.lpszClassName = L"WallpaperChangerClass";
    RegisterClassExW(&wc);

    CreateWindowExW(0, L"WallpaperChangerClass", L"Wallpaper Changer",
                    0, 0, 0, 0, 0,
                    HWND_MESSAGE, NULL, hInstance, NULL);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    CoUninitialize();
    return (int)msg.wParam;
}
