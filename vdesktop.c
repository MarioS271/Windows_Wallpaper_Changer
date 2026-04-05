/* vdesktop.c – Live wallpaper update for all Windows 11 virtual desktops.
 *
 * Strategy
 * --------
 * Windows Explorer caches each virtual desktop's wallpaper bitmap in memory.
 * The only way to update inactive desktops at runtime is through the
 * undocumented IVirtualDesktopManagerInternal COM interface, obtained via
 * IServiceProvider::QueryService on the ImmersiveShell.
 *
 * We call UpdateWallpaperPathForAllDesktops(HSTRING path), which tells
 * Explorer to refresh the in-memory bitmap for every desktop at once.
 *
 * The vtable layout of IVirtualDesktopManagerInternal changed across builds:
 *
 *   Build 22621 (22H2) – 22 interface methods before the break point
 *     UpdateWallpaperPathForAllDesktops → vtable slot 17
 *
 *   Build 26100 (24H2) – SwitchDesktopAndMoveForegroundView inserted at
 *     slot 10, shifting all later methods by one:
 *     UpdateWallpaperPathForAllDesktops → vtable slot 18
 *
 * Supported: Windows 11 Build 22621 and later.
 * Unsupported builds return FALSE so the caller falls back to registry writes.
 *
 * References
 * ----------
 *   https://github.com/hwtnb/SylphyHornPlusWin11
 *   https://github.com/Grabacr07/VirtualDesktop
 *   https://github.com/MScholtes/VirtualDesktop
 *   https://github.com/Ciantic/VirtualDesktopAccessor
 */

#define WIN32_LEAN_AND_MEAN
#define UNICODE
#define _UNICODE

#include <windows.h>
#include <objbase.h>
#include "vdesktop.h"

/* ---------------------------------------------------------------------------
 * HSTRING – opaque WinRT string handle.
 * Defined here to avoid requiring WinRT/UWP headers (winstring.h may be
 * absent in older MinGW distributions).
 * --------------------------------------------------------------------------- */
#ifndef HSTRING
typedef struct HSTRING__ { int unused; } *HSTRING;
#endif

typedef HRESULT (WINAPI *PFN_WindowsCreateString)(LPCWSTR src, UINT32 len, HSTRING *out);
typedef HRESULT (WINAPI *PFN_WindowsDeleteString)(HSTRING str);

/* ---------------------------------------------------------------------------
 * GUIDs
 * --------------------------------------------------------------------------- */

/* {C2F03A33-21F5-47FA-B4BB-156362A2F239} – ImmersiveShell (Explorer's COM host) */
static const GUID s_CLSID_ImmersiveShell = {
    0xC2F03A33, 0x21F5, 0x47FA,
    {0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39}
};

/* {C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B} – service ID for QueryService */
static const GUID s_SID_VDesktopManagerInternal = {
    0xC5E0CDCA, 0x7B6E, 0x41B2,
    {0x9F, 0xC4, 0xD9, 0x39, 0x75, 0xCC, 0x46, 0x7B}
};

/* {53F5CA0B-158F-4124-900C-057158060B27} – IVirtualDesktopManagerInternal (Build 22621+) */
static const GUID s_IID_IVDesktopManagerInternal = {
    0x53F5CA0B, 0x158F, 0x4124,
    {0x90, 0x0C, 0x05, 0x71, 0x58, 0x06, 0x0B, 0x27}
};

/* {6D5140C1-7436-11CE-8034-00AA006009FA} – IServiceProvider */
static const GUID s_IID_IServiceProvider = {
    0x6D5140C1, 0x7436, 0x11CE,
    {0x80, 0x34, 0x00, 0xAA, 0x00, 0x60, 0x09, 0xFA}
};

/* ---------------------------------------------------------------------------
 * Helpers
 * --------------------------------------------------------------------------- */

static DWORD GetBuildNumber(void) {
    typedef LONG (WINAPI *PFN_RtlGetVersion)(OSVERSIONINFOEXW *);
    HMODULE hNtdll = GetModuleHandleW(L"ntdll.dll");
    if (!hNtdll) return 0;
    PFN_RtlGetVersion pfn =
        (PFN_RtlGetVersion)(void *)GetProcAddress(hNtdll, "RtlGetVersion");
    if (!pfn) return 0;
    OSVERSIONINFOEXW vi;
    vi.dwOSVersionInfoSize = sizeof(vi);
    if (pfn(&vi) != 0) return 0;
    return vi.dwBuildNumber;
}

/* Call IUnknown::Release via vtable slot 2 without needing full type info. */
static void ComRelease(void *pUnk) {
    if (!pUnk) return;
    typedef ULONG (STDMETHODCALLTYPE *PFN_Release)(void *);
    void **vtbl = *(void ***)pUnk;
    ((PFN_Release)vtbl[2])(pUnk);
}

/* ---------------------------------------------------------------------------
 * Public API
 * --------------------------------------------------------------------------- */

BOOL VDesktop_SetWallpaperAllDesktops(const wchar_t *path) {
    DWORD build = GetBuildNumber();

    /* Only supported on Win11 Build 22621+ where SetDesktopWallpaper exists. */
    if (build < 22621) return FALSE;

    /* combase.dll exports WindowsCreateString/WindowsDeleteString.
     * It is always loaded by the time we get here because CoInitializeEx
     * was already called. */
    HMODULE hComBase = GetModuleHandleW(L"combase.dll");
    if (!hComBase) return FALSE;

    PFN_WindowsCreateString pfnCreate =
        (PFN_WindowsCreateString)(void *)GetProcAddress(hComBase, "WindowsCreateString");
    PFN_WindowsDeleteString pfnDelete =
        (PFN_WindowsDeleteString)(void *)GetProcAddress(hComBase, "WindowsDeleteString");
    if (!pfnCreate || !pfnDelete) return FALSE;

    /* Connect to ImmersiveShell (runs inside Explorer.exe).
     * IServiceProvider is defined in <servprov.h> (via objbase.h). */
    IServiceProvider *pSP = NULL;
    HRESULT hr = CoCreateInstance(
        &s_CLSID_ImmersiveShell, NULL,
        CLSCTX_LOCAL_SERVER,
        &s_IID_IServiceProvider, (void **)&pSP);
    if (FAILED(hr) || !pSP) return FALSE;

    /* Obtain IVirtualDesktopManagerInternal. */
    void *pMgr = NULL;
    hr = pSP->lpVtbl->QueryService(
        pSP,
        &s_SID_VDesktopManagerInternal,
        &s_IID_IVDesktopManagerInternal,
        &pMgr);
    pSP->lpVtbl->Release(pSP);
    if (FAILED(hr) || !pMgr) return FALSE;

    /* Determine vtable slot for UpdateWallpaperPathForAllDesktops.
     *
     * Build 22621 vtable layout (slot numbers, 0-based):
     *   [0]  QueryInterface   [1] AddRef   [2] Release
     *   [3]  GetCount
     *   [4]  MoveViewToDesktop
     *   [5]  CanViewMoveDesktops
     *   [6]  GetCurrentDesktop
     *   [7]  GetDesktops
     *   [8]  GetAdjacentDesktop
     *   [9]  SwitchDesktop
     *   [10] CreateDesktop
     *   [11] MoveDesktop
     *   [12] RemoveDesktop
     *   [13] FindDesktop
     *   [14] GetDesktopSwitchIncludeExcludeViews
     *   [15] SetDesktopName
     *   [16] SetDesktopWallpaper
     *   [17] UpdateWallpaperPathForAllDesktops  ← slot 17
     *   ...
     *
     * Build 26100 (24H2) inserted SwitchDesktopAndMoveForegroundView at [10],
     * pushing every subsequent slot up by one:
     *   [17] SetDesktopWallpaper
     *   [18] UpdateWallpaperPathForAllDesktops  ← slot 18
     */
    int vtblIdx = (build >= 26100) ? 18 : 17;

    /* Wrap the path in an HSTRING as required by the COM ABI. */
    HSTRING hstr = NULL;
    hr = pfnCreate(path, (UINT32)wcslen(path), &hstr);
    if (FAILED(hr)) { ComRelease(pMgr); return FALSE; }

    typedef HRESULT (STDMETHODCALLTYPE *PFN_UpdateWallpaper)(void *, HSTRING);
    void **vtbl = *(void ***)pMgr;
    hr = ((PFN_UpdateWallpaper)vtbl[vtblIdx])(pMgr, hstr);

    pfnDelete(hstr);
    ComRelease(pMgr);

    return SUCCEEDED(hr);
}
