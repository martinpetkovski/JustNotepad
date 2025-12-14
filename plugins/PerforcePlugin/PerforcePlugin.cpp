#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <string>
#include <vector>
#include <algorithm>
#include <filesystem>
#include <shlwapi.h>
#include "resource.h"

#pragma comment(lib, "shlwapi.lib")

HINSTANCE g_hInst = NULL;
std::wstring g_LastOutput;
WNDPROC g_OldEditorWndProc = NULL;
HWND g_hEditor = NULL;
DWORD g_P4LastCheck = 0;
std::wstring g_P4StatusCache;
std::wstring g_P4LastFile;

struct P4Settings {
    std::wstring server;
    std::wstring user;
    std::wstring client;
    std::wstring password;
} g_P4Settings;

std::wstring g_SettingsPath;

void LoadSettings() {
    if (g_SettingsPath.empty()) return;
    
    TCHAR buf[256];
    GetPrivateProfileString(_T("Perforce"), _T("Server"), _T(""), buf, 256, g_SettingsPath.c_str());
    g_P4Settings.server = buf;
    GetPrivateProfileString(_T("Perforce"), _T("User"), _T(""), buf, 256, g_SettingsPath.c_str());
    g_P4Settings.user = buf;
    GetPrivateProfileString(_T("Perforce"), _T("Client"), _T(""), buf, 256, g_SettingsPath.c_str());
    g_P4Settings.client = buf;
    // Password usually not saved in plain text, but for this demo...
    GetPrivateProfileString(_T("Perforce"), _T("Password"), _T(""), buf, 256, g_SettingsPath.c_str());
    g_P4Settings.password = buf;
}

void SaveSettings() {
    if (g_SettingsPath.empty()) return;
    
    WritePrivateProfileString(_T("Perforce"), _T("Server"), g_P4Settings.server.c_str(), g_SettingsPath.c_str());
    WritePrivateProfileString(_T("Perforce"), _T("User"), g_P4Settings.user.c_str(), g_SettingsPath.c_str());
    WritePrivateProfileString(_T("Perforce"), _T("Client"), g_P4Settings.client.c_str(), g_SettingsPath.c_str());
    WritePrivateProfileString(_T("Perforce"), _T("Password"), g_P4Settings.password.c_str(), g_SettingsPath.c_str());
}

// Helper to get file path from main window title
std::wstring GetCurrentFilePath(HWND hEditor) {
    HWND hMain = GetParent(hEditor);

    // Try to get path from window property first
    HANDLE hProp = GetProp(hMain, L"FullPath");
    if (hProp) {
        TCHAR* path = (TCHAR*)hProp;
        if (path && path[0]) {
            return std::wstring(path);
        }
    }

    int len = GetWindowTextLength(hMain);
    if (len <= 0) return L"";

    std::vector<wchar_t> buf(len + 1);
    GetWindowText(hMain, buf.data(), len + 1);
    std::wstring title = buf.data();

    // Format is: "[* ]Filename - Just Notepad"
    std::wstring suffix = L" - Just Notepad";
    if (title.length() >= suffix.length() && 
        title.compare(title.length() - suffix.length(), suffix.length(), suffix) == 0) {
        title = title.substr(0, title.length() - suffix.length());
    }

    if (title.length() > 0 && title[0] == L'*') {
        title = title.substr(1);
        if (title.length() > 0 && title[0] == L' ') {
            title = title.substr(1);
        }
    }

    if (title == L"Untitled") return L"";
    return title;
}

std::wstring GetPerforcePath() {
    TCHAR szPath[MAX_PATH];
    if (SearchPath(NULL, _T("p4.exe"), NULL, MAX_PATH, szPath, NULL)) {
        return std::wstring(szPath);
    }
    
    // Common locations? P4 is usually in PATH, but let's check Program Files
    const wchar_t* commonPaths[] = {
        L"C:\\Program Files\\Perforce\\p4.exe",
        L"C:\\Program Files (x86)\\Perforce\\p4.exe"
    };

    for (const auto& path : commonPaths) {
        if (std::filesystem::exists(path)) {
            return std::wstring(path);
        }
    }

    return L"p4";
}

std::wstring RunPerforceCommand(const std::wstring& cmd, const std::wstring& dir) {
    SECURITY_ATTRIBUTES sa;
    sa.nLength = sizeof(SECURITY_ATTRIBUTES);
    sa.bInheritHandle = TRUE;
    sa.lpSecurityDescriptor = NULL;

    HANDLE hRead, hWrite;
    if (!CreatePipe(&hRead, &hWrite, &sa, 0)) return L"Error creating pipe";
    SetHandleInformation(hRead, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFO si = { sizeof(STARTUPINFO) };
    si.cb = sizeof(STARTUPINFO);
    si.hStdError = hWrite;
    si.hStdOutput = hWrite;
    si.dwFlags |= STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    PROCESS_INFORMATION pi = { 0 };
    
    std::wstring p4Path = GetPerforcePath();
    std::wstring fullCmd = L"\"" + p4Path + L"\"";
    
    if (!g_P4Settings.server.empty()) fullCmd += L" -p " + g_P4Settings.server;
    if (!g_P4Settings.user.empty()) fullCmd += L" -u " + g_P4Settings.user;
    if (!g_P4Settings.client.empty()) fullCmd += L" -c " + g_P4Settings.client;
    if (!g_P4Settings.password.empty()) fullCmd += L" -P " + g_P4Settings.password;
    
    fullCmd += L" " + cmd;
    
    std::vector<wchar_t> cmdBuf(fullCmd.begin(), fullCmd.end());
    cmdBuf.push_back(0);

    if (!CreateProcess(NULL, cmdBuf.data(), NULL, NULL, TRUE, 0, NULL, dir.empty() ? NULL : dir.c_str(), &si, &pi)) {
        CloseHandle(hRead);
        CloseHandle(hWrite);
        return L"Failed to execute p4 command";
    }

    CloseHandle(hWrite);

    std::string output;
    char buffer[4096];
    DWORD bytesRead;
    while (ReadFile(hRead, buffer, sizeof(buffer), &bytesRead, NULL) && bytesRead > 0) {
        output.append(buffer, bytesRead);
    }

    CloseHandle(hRead);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    // Convert to wstring
    int len = MultiByteToWideChar(CP_UTF8, 0, output.c_str(), -1, NULL, 0);
    if (len == 0) {
        // Try system default code page if UTF-8 fails
        len = MultiByteToWideChar(CP_ACP, 0, output.c_str(), -1, NULL, 0);
        if (len == 0) return L"";
        std::vector<wchar_t> wbuf(len);
        MultiByteToWideChar(CP_ACP, 0, output.c_str(), -1, wbuf.data(), len);
        return std::wstring(wbuf.data());
    }
    
    std::vector<wchar_t> wbuf(len);
    MultiByteToWideChar(CP_UTF8, 0, output.c_str(), -1, wbuf.data(), len);
    return std::wstring(wbuf.data());
}

INT_PTR CALLBACK OutputDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        SetDlgItemText(hDlg, IDC_P4_OUTPUT_TEXT, g_LastOutput.c_str());
        return (INT_PTR)TRUE;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    case WM_SIZE:
        {
            int width = LOWORD(lParam);
            int height = HIWORD(lParam);
            SetWindowPos(GetDlgItem(hDlg, IDC_P4_OUTPUT_TEXT), NULL, 7, 7, width - 14, height - 40, SWP_NOZORDER);
            SetWindowPos(GetDlgItem(hDlg, IDOK), NULL, width - 57, height - 21, 50, 14, SWP_NOZORDER);
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void ShowOutput(HWND hParent, const std::wstring& output) {
    g_LastOutput = output;
    DialogBox(g_hInst, MAKEINTRESOURCE(IDD_P4_OUTPUT), hParent, OutputDlgProc);
}

INT_PTR CALLBACK LoginDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        SetDlgItemText(hDlg, IDC_P4_SERVER, g_P4Settings.server.c_str());
        SetDlgItemText(hDlg, IDC_P4_USER, g_P4Settings.user.c_str());
        SetDlgItemText(hDlg, IDC_P4_CLIENT, g_P4Settings.client.c_str());
        SetDlgItemText(hDlg, IDC_P4_PASSWORD, g_P4Settings.password.c_str());
        return (INT_PTR)TRUE;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            TCHAR buf[256];
            GetDlgItemText(hDlg, IDC_P4_SERVER, buf, 256); g_P4Settings.server = buf;
            GetDlgItemText(hDlg, IDC_P4_USER, buf, 256); g_P4Settings.user = buf;
            GetDlgItemText(hDlg, IDC_P4_CLIENT, buf, 256); g_P4Settings.client = buf;
            GetDlgItemText(hDlg, IDC_P4_PASSWORD, buf, 256); g_P4Settings.password = buf;
            
            SaveSettings();
            
            // Test connection
            std::wstring output = RunPerforceCommand(L"info", L"");
            if (output.find(L"User name:") != std::wstring::npos) {
                MessageBox(hDlg, L"Connection successful!", L"Perforce", MB_OK | MB_ICONINFORMATION);
                EndDialog(hDlg, IDOK);
            } else {
                MessageBox(hDlg, (L"Connection failed:\n" + output).c_str(), L"Perforce", MB_OK | MB_ICONERROR);
            }
            return (INT_PTR)TRUE;
        } else if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, IDCANCEL);
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void SubclassEditor(HWND hEditor);

void P4Login(HWND hEditor) {
    SubclassEditor(hEditor);
    DialogBox(g_hInst, MAKEINTRESOURCE(IDD_P4_LOGIN), GetParent(hEditor), LoginDlgProc);
}

void CheckAndCheckout(HWND hEditor) {
    std::wstring filePath = GetCurrentFilePath(hEditor);
    if (filePath.empty()) return;
    
    bool needsCheckout = false;
    
    // 1. Check ReadOnly attribute
    DWORD attrs = GetFileAttributes(filePath.c_str());
    if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_READONLY)) {
        needsCheckout = true;
    }
    
    // 2. Check P4 Status (if we know it)
    if (!needsCheckout) {
        if (g_P4LastFile == filePath && !g_P4StatusCache.empty()) {
            if (g_P4StatusCache.find(L"synced") != std::wstring::npos) {
                needsCheckout = true;
            }
        } else {
            // Status unknown. If we are here, it means the file is NOT ReadOnly.
            // But it might be "synced" and writable (e.g. user changed attributes manually).
            // Or it might be a new file not in P4.
            // To be safe, we should check, but we don't want to block every keypress.
            // However, if this is called from "Modified" event (which happens once per dirty state),
            // we SHOULD check.
            // But CheckAndCheckout doesn't know if it's called from keypress or event.
            
            // Let's assume if cache is empty, we should check.
            // But we need to be careful about performance.
            // For now, let's rely on the fact that "Modified" event is rare (once per save cycle).
            // But WM_CHAR is frequent.
            
            // If we are in WM_CHAR, we probably shouldn't block if cache is empty.
            // But if we are in OnFileEvent, we should.
            
            // Let's just check if it's in P4.
            // If we don't check, we miss the case where file is writable but needs checkout.
            // We can optimize by checking if we checked recently.
            if (GetTickCount() - g_P4LastCheck > 5000) {
                 // It's been a while, let's check.
                 // But we can't just set needsCheckout=true because that forces fstat.
                 // We can do a quick check? No, fstat is the check.
                 
                 // Let's just force it if cache is empty or stale.
                 needsCheckout = true;
            }
        }
    }

    if (needsCheckout) {
        std::wstring dir = std::filesystem::path(filePath).parent_path().wstring();
        std::wstring output = RunPerforceCommand(L"fstat \"" + filePath + L"\"", dir);
        
        if (output.find(L"clientFile") != std::wstring::npos) {
            // It is in Perforce.
            // Check if already opened
            if (output.find(L"action edit") == std::wstring::npos && 
                output.find(L"action add") == std::wstring::npos) {
                
                // Not opened. Checkout.
                RunPerforceCommand(L"edit \"" + filePath + L"\"", dir);
                
                // Update attributes and UI
                SetFileAttributes(filePath.c_str(), attrs & ~FILE_ATTRIBUTE_READONLY);
                SendMessage(hEditor, EM_SETREADONLY, FALSE, 0);
                
                g_P4LastCheck = 0; // Invalidate cache
                // Update cache immediately to avoid repeated calls
                g_P4StatusCache = L"P4: edit";
                g_P4LastCheck = GetTickCount();
            }
        }
    }
}

LRESULT CALLBACK EditorWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    if (message == WM_CHAR || message == WM_KEYDOWN || message == WM_PASTE || message == WM_CUT) {
        // Check for modification attempts
        // Filter out navigation keys
        bool isMod = true;
        if (message == WM_KEYDOWN) {
            if (wParam >= VK_LEFT && wParam <= VK_DOWN) isMod = false;
            if (wParam == VK_SHIFT || wParam == VK_CONTROL || wParam == VK_MENU) isMod = false;
            if (wParam >= VK_F1 && wParam <= VK_F24) isMod = false;
        }
        
        if (isMod) {
            // Check if ReadOnly
            bool bReadOnly = false;
            DWORD style = GetWindowLong(hWnd, GWL_STYLE);
            if (style & ES_READONLY) bReadOnly = true;
            else {
                // Try EM_GETOPTIONS (WM_USER + 78)
                // ECO_READONLY is 0x0004
                DWORD options = SendMessage(hWnd, WM_USER + 78, 0, 0);
                if (options & 0x0004) bReadOnly = true;
            }
            
            if (bReadOnly) {
                 CheckAndCheckout(hWnd);
                 // If we successfully checked out, the style should be gone.
                 // Re-check
                 bReadOnly = false;
                 style = GetWindowLong(hWnd, GWL_STYLE);
                 if (style & ES_READONLY) bReadOnly = true;
                 else {
                    DWORD options = SendMessage(hWnd, WM_USER + 78, 0, 0);
                    if (options & 0x0004) bReadOnly = true;
                 }
            } else {
                // Even if not read-only, check if we need to checkout (e.g. file is writable but not opened in p4)
                CheckAndCheckout(hWnd);
            }
        }
    }
    return CallWindowProc(g_OldEditorWndProc, hWnd, message, wParam, lParam);
}

PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
    if (wcscmp(eventType, L"Loaded") == 0) {
        SubclassEditor(hEditor);
        // We don't check out immediately on load, only on edit.
        // But we ensure we are subclassed so we can catch the edit.
    }
    else if (wcscmp(eventType, L"Modified") == 0) {
        // File became dirty. Ensure it is checked out.
        CheckAndCheckout(hEditor);
    }
}

void SubclassEditor(HWND hEditor) {
    if (g_hEditor == hEditor) return;
    if (g_OldEditorWndProc) return; // Already subclassed
    
    g_hEditor = hEditor;
    g_OldEditorWndProc = (WNDPROC)SetWindowLongPtr(hEditor, GWLP_WNDPROC, (LONG_PTR)EditorWndProc);
}

void RunP4Action(HWND hEditor, const std::wstring& action) {
    SubclassEditor(hEditor); // Ensure we are hooked

    std::wstring filePath = GetCurrentFilePath(hEditor);
    if (filePath.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Perforce", MB_ICONWARNING);
        return;
    }

    std::wstring dir = std::filesystem::path(filePath).parent_path().wstring();

    std::wstring cmd = action + L" \"" + filePath + L"\"";
    std::wstring output = RunPerforceCommand(cmd, dir);
    
    // Invalidate cache
    g_P4LastCheck = 0;

    ShowOutput(GetParent(hEditor), output);
}

void P4Edit(HWND hEditor) { RunP4Action(hEditor, L"edit"); }
void P4Add(HWND hEditor) { RunP4Action(hEditor, L"add"); }
void P4Revert(HWND hEditor) { 
    if (MessageBox(hEditor, L"Are you sure you want to revert changes?", L"Perforce Revert", MB_YESNO | MB_ICONQUESTION) == IDYES) {
        RunP4Action(hEditor, L"revert"); 
    }
}
void P4Diff(HWND hEditor) { RunP4Action(hEditor, L"diff"); }
void P4History(HWND hEditor) { RunP4Action(hEditor, L"filelog"); }
void P4Info(HWND hEditor) { RunP4Action(hEditor, L"fstat"); }

PluginMenuItem g_MenuItems[] = {
    { _T("Login..."), P4Login },
    { _T("Edit"), P4Edit },
    { _T("Add"), P4Add },
    { _T("Revert"), P4Revert },
    { _T("Diff"), P4Diff },
    { _T("History"), P4History },
    { _T("Info"), P4Info }
};

extern "C" {

    PLUGIN_API void Initialize(const wchar_t* settingsPath) {
        g_SettingsPath = settingsPath;
        LoadSettings();
    }

    PLUGIN_API void Shutdown() {
        SaveSettings();
    }

PLUGIN_API const TCHAR* GetPluginName() {
    return _T("Perforce");
}

PLUGIN_API const TCHAR* GetPluginDescription() {
    return _T("Basic Perforce integration (Edit, Add, Revert, Diff, History).");
}

PLUGIN_API const TCHAR* GetPluginVersion() {
    return _T("1.0");
}

PLUGIN_API const wchar_t* GetPluginLicense() {
    return L"MIT License\n\n"
           L"Copyright (c) 2025 Just Notepad Contributors\n\n"
           L"Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\n"
           L"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\n"
           L"THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO, THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
}
} // End extern "C"

#include <thread>
#include <atomic>
#include <mutex>

std::atomic<bool> g_P4CheckInProgress(false);
std::mutex g_P4CacheMutex;

void CheckStatusWorker(std::wstring filePath) {
    std::wstring dir = std::filesystem::path(filePath).parent_path().wstring();
    std::wstring output = RunPerforceCommand(L"fstat \"" + filePath + L"\"", dir);
    
    std::wstring status = L"";
    if (output.find(L"clientFile") != std::wstring::npos) {
        // Parse action
        size_t actionPos = output.find(L"action ");
        if (actionPos != std::wstring::npos) {
            size_t end = output.find(L"\n", actionPos);
            std::wstring action = output.substr(actionPos + 7, end - (actionPos + 7));
            // Trim
            while (!action.empty() && iswspace(action.back())) action.pop_back();
            status = L"P4: " + action;
        } else {
            status = L"P4: synced";
        }
    }

    {
        std::lock_guard<std::mutex> lock(g_P4CacheMutex);
        g_P4StatusCache = status;
        g_P4LastCheck = GetTickCount();
    }
    g_P4CheckInProgress = false;
}

extern "C" {
PLUGIN_API const wchar_t* GetPluginStatus(const wchar_t* filePath) {
    if (!filePath || !filePath[0]) return NULL;
    
    std::lock_guard<std::mutex> lock(g_P4CacheMutex);
    
    if (g_P4LastFile != filePath) {
        g_P4LastFile = filePath;
        g_P4StatusCache = L"";
        g_P4LastCheck = 0; // Force update
    }

    if (GetTickCount() - g_P4LastCheck > 2000) {
        if (!g_P4CheckInProgress) {
            g_P4CheckInProgress = true;
            std::thread(CheckStatusWorker, std::wstring(filePath)).detach();
        }
    }
    
    return g_P4StatusCache.empty() ? NULL : g_P4StatusCache.c_str();
}

PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
    // Try to find editor window and subclass it early
    // This is a bit hacky but we need to hook before user interaction if possible
    // However, GetPluginMenuItems is called early.
    // We can't easily find the editor HWND here without context.
    // But we can try to find the main window of the current process.
    
    if (!g_hEditor) {
        HWND hMain = NULL;
        DWORD pid = GetCurrentProcessId();
        // We can't use EnumWindows easily in a DLL without a callback, but we can try FindWindow if we knew the class.
        // Instead, let's rely on the first menu action to hook, OR hook when we can.
        // Actually, we can just load settings here.
        // LoadSettings();
        
        // Auto-login if password is saved
        if (!g_P4Settings.password.empty()) {
             // We don't need to do anything special if the password is in the settings,
             // RunPerforceCommand uses it.
             // But we can verify connection silently.
             // RunPerforceCommand(L"login", L""); // This might hang if it asks for password on stdin
             // Since we pass -P password, we usually don't need 'p4 login' unless we are using tickets.
             // If we are using tickets, we might need to run 'echo password | p4 login'.
             // But our RunPerforceCommand passes -P on every command, which works for many setups.
             // If the user wants to use tickets, they should use 'p4 login' manually once.
             // But let's assume the user wants us to try to establish a session.
             
             // Actually, if we pass -P, we don't need to login.
             // So "Connect automatically" is effectively done by loading settings.
        }
    }

    *count = 7;
    return g_MenuItems;
}

} // extern "C"

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
    switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}
