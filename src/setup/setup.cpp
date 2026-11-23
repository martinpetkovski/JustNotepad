// JustNotepad installer and uninstaller (one small exe, no dependencies).
//
// Install (default): for the current user, no admin rights needed.
//   - JustNotepad.exe, carried inside this exe, goes to %LOCALAPPDATA%\Programs\JustNotepad
//   - a copy of this exe goes next to it as uninstall.exe
//   - Start menu shortcut (and a desktop one if ticked)
//   - an entry in Settings > Apps, so it can be removed like any other app
// Uninstall (uninstall.exe /uninstall, which is what Settings runs):
//   removes all of the above. Your settings and session are kept in %APPDATA%.
#ifndef UNICODE
#define UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <commctrl.h>
#include <objbase.h>
#include <shlobj.h>
#include <shobjidl.h>

#include <string>

#include "../win/theme.h"
#include "../win/version.h"
#include "resource.h"

#define SW2(x) L##x
#define SW(x) SW2(x)

namespace {

const wchar_t* const kTitle = L"JustNotepad Setup";
const wchar_t* const kUninstallKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\JustNotepad";
const wchar_t* const kAppFrameClass = L"JustNotepadFrame";  // main window class of the app

bool g_uninstall = false;
Palette g_pal;
HFONT g_titleFont = nullptr;

std::wstring knownFolder(REFKNOWNFOLDERID id) {
    PWSTR p = nullptr;
    std::wstring s;
    if (SUCCEEDED(SHGetKnownFolderPath(id, 0, nullptr, &p))) s = p;
    CoTaskMemFree(p);
    return s;
}

std::wstring installDir() { return knownFolder(FOLDERID_LocalAppData) + L"\\Programs\\JustNotepad"; }
std::wstring startMenuLink() { return knownFolder(FOLDERID_Programs) + L"\\JustNotepad.lnk"; }
std::wstring desktopLink() { return knownFolder(FOLDERID_Desktop) + L"\\JustNotepad.lnk"; }

void fail(HWND owner, const std::wstring& what, DWORD err) {
    wchar_t sys[512] = L"";
    FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, err, 0, sys, 512, nullptr);
    std::wstring m = what + L"\n\n" + sys;
    MessageBoxW(owner, m.c_str(), kTitle, MB_OK | MB_ICONERROR);
}

// Asks running copies of JustNotepad to close (they keep unsaved work in
// their session) and waits for them to go.
bool closeRunningApp(HWND owner) {
    HWND w = FindWindowW(kAppFrameClass, nullptr);
    if (!w) return true;
    if (MessageBoxW(owner, L"JustNotepad is running. Close it to continue?\n\nYour open tabs and unsaved text come back next time.",
                    kTitle, MB_OKCANCEL | MB_ICONINFORMATION) != IDOK)
        return false;
    for (int round = 0; round < 50; round++) {
        w = FindWindowW(kAppFrameClass, nullptr);
        if (!w) return true;
        PostMessageW(w, WM_CLOSE, 0, 0);
        Sleep(100);
    }
    return FindWindowW(kAppFrameClass, nullptr) == nullptr;
}

bool makeShortcut(const std::wstring& link, const std::wstring& target, const std::wstring& dir) {
    IShellLinkW* sl = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&sl)))) return false;
    sl->SetPath(target.c_str());
    sl->SetWorkingDirectory(dir.c_str());
    sl->SetDescription(L"JustNotepad");
    sl->SetIconLocation(target.c_str(), 0);
    IPersistFile* pf = nullptr;
    bool ok = SUCCEEDED(sl->QueryInterface(IID_PPV_ARGS(&pf))) && SUCCEEDED(pf->Save(link.c_str(), TRUE));
    if (pf) pf->Release();
    sl->Release();
    return ok;
}

void setString(HKEY k, const wchar_t* name, const std::wstring& v) {
    RegSetValueExW(k, name, 0, REG_SZ, (const BYTE*)v.c_str(), (DWORD)((v.size() + 1) * sizeof(wchar_t)));
}
void setDword(HKEY k, const wchar_t* name, DWORD v) { RegSetValueExW(k, name, 0, REG_DWORD, (const BYTE*)&v, sizeof(v)); }

// Writes `n` bytes to `path` via a temporary file, retrying while the old file
// is still in use (a copy that is just closing).
bool writeFileReplacing(const std::wstring& path, const void* data, DWORD n, DWORD& err) {
    std::wstring tmp = path + L".new";
    HANDLE h = CreateFileW(tmp.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) {
        err = GetLastError();
        return false;
    }
    DWORD put = 0;
    bool ok = WriteFile(h, data, n, &put, nullptr) && put == n;
    err = ok ? 0 : GetLastError();
    CloseHandle(h);
    if (!ok) {
        DeleteFileW(tmp.c_str());
        return false;
    }
    for (int i = 0; i < 50; i++) {
        if (MoveFileExW(tmp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING)) return true;
        err = GetLastError();
        Sleep(100);
    }
    DeleteFileW(tmp.c_str());
    return false;
}

bool install(HWND owner, bool desktop) {
    if (!closeRunningApp(owner)) return false;
    const std::wstring dir = installDir(), exe = dir + L"\\JustNotepad.exe", unexe = dir + L"\\uninstall.exe";
    int r = SHCreateDirectoryExW(nullptr, dir.c_str(), nullptr);
    if (r != ERROR_SUCCESS && r != ERROR_ALREADY_EXISTS && r != ERROR_FILE_EXISTS) {
        fail(owner, L"Could not create " + dir, (DWORD)r);
        return false;
    }
    // The program, from inside this installer.
    HRSRC res = FindResourceW(nullptr, MAKEINTRESOURCEW(IDR_APPEXE), RT_RCDATA);
    HGLOBAL hg = res ? LoadResource(nullptr, res) : nullptr;
    const void* bytes = hg ? LockResource(hg) : nullptr;
    DWORD size = res ? SizeofResource(nullptr, res) : 0;
    DWORD err = 0;
    if (!bytes || !size) {
        fail(owner, L"This installer is damaged (the program is missing from it).", ERROR_FILE_CORRUPT);
        return false;
    }
    if (!writeFileReplacing(exe, bytes, size, err)) {
        fail(owner, L"Could not write " + exe, err);
        return false;
    }
    // A copy of this exe is the uninstaller.
    wchar_t self[MAX_PATH];
    GetModuleFileNameW(nullptr, self, MAX_PATH);
    if (lstrcmpiW(self, unexe.c_str()) != 0) CopyFileW(self, unexe.c_str(), FALSE);

    makeShortcut(startMenuLink(), exe, dir);
    if (desktop) makeShortcut(desktopLink(), exe, dir);
    else DeleteFileW(desktopLink().c_str());

    // Settings > Apps entry.
    HKEY k;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, kUninstallKey, 0, nullptr, 0, KEY_WRITE, nullptr, &k, nullptr) == ERROR_SUCCESS) {
        setString(k, L"DisplayName", L"JustNotepad");
        setString(k, L"DisplayVersion", SW(APP_VERSION_STRING));
        setString(k, L"Publisher", L"Martin Petkovski");
        setString(k, L"DisplayIcon", exe);
        setString(k, L"InstallLocation", dir);
        setString(k, L"UninstallString", L"\"" + unexe + L"\" /uninstall");
        setString(k, L"URLInfoAbout", L"https://www.najjak.com/JustNotepad/");
        setDword(k, L"EstimatedSize", (size + 1023) / 1024 * 2);  // KB: program + uninstaller
        setDword(k, L"NoModify", 1);
        setDword(k, L"NoRepair", 1);
        RegCloseKey(k);
    }
    return true;
}

bool uninstall(HWND owner) {
    if (!closeRunningApp(owner)) return false;
    const std::wstring dir = installDir();
    DeleteFileW((dir + L"\\JustNotepad.exe").c_str());
    DeleteFileW(startMenuLink().c_str());
    DeleteFileW(desktopLink().c_str());
    RegDeleteTreeW(HKEY_CURRENT_USER, kUninstallKey);
    // This exe (uninstall.exe) cannot delete itself while running: a hidden
    // cmd waits a moment, then removes it and the folder.
    wchar_t self[MAX_PATH];
    GetModuleFileNameW(nullptr, self, MAX_PATH);
    std::wstring cmd = L"cmd.exe /c ping 127.0.0.1 -n 3 > nul & del /f /q \"" + std::wstring(self) + L"\" & rmdir \"" + dir + L"\"";
    STARTUPINFOW si = {sizeof(si)};
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    PROCESS_INFORMATION pi = {};
    std::wstring buf = cmd;
    if (CreateProcessW(nullptr, &buf[0], nullptr, nullptr, FALSE, CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi)) {
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
    }
    return true;
}

INT_PTR CALLBACK setupProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg >= WM_CTLCOLORMSGBOX && msg <= WM_CTLCOLORSTATIC) return themeCtlColor(msg, wp, g_pal);
    switch (msg) {
    case WM_INITDIALOG: {
        themeDialog(dlg, g_pal);
        LOGFONTW lf = {};
        GetObjectW((HFONT)SendMessageW(dlg, WM_GETFONT, 0, 0), sizeof(lf), &lf);
        lf.lfHeight = lf.lfHeight * 16 / 10;
        lf.lfWeight = FW_BOLD;
        g_titleFont = CreateFontIndirectW(&lf);
        SendDlgItemMessageW(dlg, IDC_TITLE, WM_SETFONT, (WPARAM)g_titleFont, TRUE);
        int sz = MulDiv(32, (int)GetDpiForWindow(dlg), 96);
        HICON icon = nullptr;
        LoadIconWithScaleDown(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDI_APP), sz, sz, &icon);
        if (icon) {
            SendDlgItemMessageW(dlg, IDC_ICON, STM_SETICON, (WPARAM)icon, 0);
            SendMessageW(dlg, WM_SETICON, ICON_BIG, (LPARAM)icon);
        }
        SetDlgItemTextW(dlg, IDC_PATH, installDir().c_str());
        if (g_uninstall) {
            SetWindowTextW(dlg, L"Uninstall JustNotepad");
            SetDlgItemTextW(dlg, IDC_TITLE, L"Uninstall JustNotepad");
            SetDlgItemTextW(dlg, IDC_INFO, L"JustNotepad will be removed from this folder, the Start menu and the desktop. "
                                           L"Your settings are kept.");
            ShowWindow(GetDlgItem(dlg, IDC_DESKTOP), SW_HIDE);
            ShowWindow(GetDlgItem(dlg, IDC_RUN), SW_HIDE);
            SetDlgItemTextW(dlg, IDOK, L"&Uninstall");
        } else {
            std::wstring info = L"JustNotepad " SW(APP_VERSION_STRING) L" will be installed for you (no admin rights needed) in:";
            SetDlgItemTextW(dlg, IDC_INFO, info.c_str());
            const bool upgrading = GetFileAttributesW((installDir() + L"\\JustNotepad.exe").c_str()) != INVALID_FILE_ATTRIBUTES;
            if (upgrading) SetDlgItemTextW(dlg, IDOK, L"&Update");
            CheckDlgButton(dlg, IDC_DESKTOP,
                           GetFileAttributesW(desktopLink().c_str()) != INVALID_FILE_ATTRIBUTES ? BST_CHECKED : BST_UNCHECKED);
            CheckDlgButton(dlg, IDC_RUN, BST_CHECKED);
        }
        return TRUE;
    }
    case WM_DRAWITEM:
        if (themeDrawDialogButton((DRAWITEMSTRUCT*)lp, g_pal)) return TRUE;
        break;
    case WM_COMMAND:
        if (LOWORD(wp) == IDOK) {
            HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
            bool ok;
            if (g_uninstall) {
                ok = uninstall(dlg);
                SetCursor(old);
                if (ok) MessageBoxW(dlg, L"JustNotepad has been removed.", L"Uninstall JustNotepad", MB_OK | MB_ICONINFORMATION);
            } else {
                const bool run = IsDlgButtonChecked(dlg, IDC_RUN) == BST_CHECKED;
                ok = install(dlg, IsDlgButtonChecked(dlg, IDC_DESKTOP) == BST_CHECKED);
                SetCursor(old);
                if (ok && run) ShellExecuteW(nullptr, L"open", (installDir() + L"\\JustNotepad.exe").c_str(), nullptr,
                                             installDir().c_str(), SW_SHOWNORMAL);
                else if (ok) MessageBoxW(dlg, L"JustNotepad is installed. Find it in the Start menu.", kTitle, MB_OK | MB_ICONINFORMATION);
            }
            if (ok) EndDialog(dlg, IDOK);
            return TRUE;
        }
        if (LOWORD(wp) == IDCANCEL) {
            EndDialog(dlg, IDCANCEL);
            return TRUE;
        }
        break;
    }
    return FALSE;
}

}  // namespace

int WINAPI wWinMain(HINSTANCE inst, HINSTANCE, PWSTR cmd, int) {
    g_uninstall = cmd && (wcsstr(cmd, L"/uninstall") || wcsstr(cmd, L"-uninstall"));
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    INITCOMMONCONTROLSEX icc = {sizeof(icc), ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&icc);
    themeInit();
    gdiplusStart();
    g_pal = makePalette(systemUsesDarkMode());
    themeSetAppDark(g_pal.dark);
    DialogBoxParamW(inst, MAKEINTRESOURCEW(IDD_SETUP), nullptr, setupProc, 0);
    gdiplusStop();
    if (g_titleFont) DeleteObject(g_titleFont);
    CoUninitialize();
    return 0;
}
