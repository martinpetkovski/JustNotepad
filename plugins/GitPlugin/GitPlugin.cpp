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
std::wstring g_CommitMessage;
DWORD g_LastCheck = 0;

// Helper to get file path from main window title
std::wstring GetCurrentFilePath(HWND hEditor) {
    HWND hMain = GetParent(hEditor);

    // Try to get path from window property first (New method)
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
    
    // Remove suffix
    std::wstring suffix = L" - Just Notepad";
    if (title.length() >= suffix.length() && 
        title.compare(title.length() - suffix.length(), suffix.length(), suffix) == 0) {
        title = title.substr(0, title.length() - suffix.length());
    }

    // Remove prefix
    if (title.length() > 0 && title[0] == L'*') {
        title = title.substr(1);
        if (title.length() > 0 && title[0] == L' ') {
            title = title.substr(1);
        }
    }

    if (title == L"Untitled") return L"";
    return title;
}

std::wstring GetGitPath() {
    // Check PATH first
    TCHAR szPath[MAX_PATH];
    if (SearchPath(NULL, _T("git.exe"), NULL, MAX_PATH, szPath, NULL)) {
        return std::wstring(szPath);
    }

    // Check common locations
    const wchar_t* commonPaths[] = {
        L"C:\\Program Files\\Git\\cmd\\git.exe",
        L"C:\\Program Files\\Git\\bin\\git.exe",
        L"C:\\Program Files (x86)\\Git\\cmd\\git.exe",
        L"C:\\Program Files (x86)\\Git\\bin\\git.exe"
    };

    for (const auto& path : commonPaths) {
        if (PathFileExists(path)) {
            return std::wstring(path);
        }
    }

    // Check Local AppData
    if (GetEnvironmentVariable(_T("LOCALAPPDATA"), szPath, MAX_PATH)) {
        std::wstring localAppData = szPath;
        std::wstring gitPath = localAppData + L"\\Programs\\Git\\cmd\\git.exe";
        if (PathFileExists(gitPath.c_str())) {
            return gitPath;
        }
    }

    return L"git"; // Fallback to just "git" and hope for the best
}

std::wstring RunGitCommand(const std::wstring& cmd, const std::wstring& dir) {
    SECURITY_ATTRIBUTES sa;
    sa.nLength = sizeof(SECURITY_ATTRIBUTES);
    sa.bInheritHandle = TRUE;
    sa.lpSecurityDescriptor = NULL;

    HANDLE hRead, hWrite;
    if (!CreatePipe(&hRead, &hWrite, &sa, 0)) return L"Error creating pipe";
    SetHandleInformation(hRead, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFO si = { sizeof(STARTUPINFO) };
    si.dwFlags = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    si.hStdOutput = hWrite;
    si.hStdError = hWrite;
    si.wShowWindow = SW_HIDE;

    PROCESS_INFORMATION pi = { 0 };
    
    std::wstring gitExe = GetGitPath();
    std::wstring commandLine = L"\"" + gitExe + L"\" " + cmd;
    std::vector<wchar_t> cmdBuf(commandLine.begin(), commandLine.end());
    cmdBuf.push_back(0);

    if (!CreateProcess(NULL, cmdBuf.data(), NULL, NULL, TRUE, 0, NULL, dir.c_str(), &si, &pi)) {
        DWORD err = GetLastError();
        
        // Try fallback to cmd.exe
        std::wstring cmdLineFallback = L"cmd.exe /c \"\"" + gitExe + L"\" " + cmd + L"\"";
        std::vector<wchar_t> cmdBufFallback(cmdLineFallback.begin(), cmdLineFallback.end());
        cmdBufFallback.push_back(0);
        
        if (CreateProcess(NULL, cmdBufFallback.data(), NULL, NULL, TRUE, 0, NULL, dir.c_str(), &si, &pi)) {
            // Fallback succeeded
            goto ReadOutput;
        }
        
        CloseHandle(hRead);
        CloseHandle(hWrite);
        
        std::wstring msg = L"Error starting git.\nCommand: " + commandLine + 
                          L"\nFallback: " + cmdLineFallback +
                          L"\nWorkDir: " + std::wstring(dir) +
                          L"\nError Code: " + std::to_wstring(err);
        return msg;
    }

ReadOutput:
    CloseHandle(hWrite);

    std::string output;
    char buffer[4096];
    DWORD bytesRead;
    while (ReadFile(hRead, buffer, sizeof(buffer), &bytesRead, NULL) && bytesRead > 0) {
        output.append(buffer, bytesRead);
    }

    CloseHandle(hRead);
    WaitForSingleObject(pi.hProcess, INFINITE);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    // Convert to wstring (assume UTF-8 from git)
    int len = MultiByteToWideChar(CP_UTF8, 0, output.c_str(), -1, NULL, 0);
    if (len == 0) return L"";
    std::vector<wchar_t> wbuf(len);
    MultiByteToWideChar(CP_UTF8, 0, output.c_str(), -1, wbuf.data(), len);
    
    // Normalize line endings
    std::wstring woutput = wbuf.data();
    std::wstring normalized;
    for (size_t i = 0; i < woutput.length(); ++i) {
        if (woutput[i] == L'\n' && (i == 0 || woutput[i-1] != L'\r')) {
            normalized += L"\r\n";
        } else {
            normalized += woutput[i];
        }
    }

    return normalized;
}

INT_PTR CALLBACK GitCommitDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        SetFocus(GetDlgItem(hDlg, IDC_GIT_COMMIT_MSG));
        return FALSE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            TCHAR buffer[1024];
            GetDlgItemText(hDlg, IDC_GIT_COMMIT_MSG, buffer, 1024);
            g_CommitMessage = buffer;
            EndDialog(hDlg, IDOK);
            return TRUE;
        } else if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
        }
        break;
    }
    return FALSE;
}


INT_PTR CALLBACK GitOutputDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        SetDlgItemText(hDlg, IDC_GIT_OUTPUT_TEXT, g_LastOutput.c_str());
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
            HWND hEdit = GetDlgItem(hDlg, IDC_GIT_OUTPUT_TEXT);
            HWND hBtn = GetDlgItem(hDlg, IDOK);
            
            if (hEdit && hBtn) {
                MoveWindow(hEdit, 7, 7, width - 14, height - 35, TRUE);
                MoveWindow(hBtn, width - 57, height - 25, 50, 14, TRUE);
            }
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void ShowGitOutput(HWND hParent, const std::wstring& output) {
    g_LastOutput = output;
    DialogBox(g_hInst, MAKEINTRESOURCE(IDD_GIT_OUTPUT), hParent, GitOutputDlgProc);
}

void GitStatus(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    TCHAR dir[MAX_PATH];
    wcscpy_s(dir, path.c_str());
    PathRemoveFileSpec(dir);

    std::wstring output = RunGitCommand(L"status", dir);
    ShowGitOutput(GetParent(hEditor), output);
}

void GitAdd(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    TCHAR dir[MAX_PATH];
    wcscpy_s(dir, path.c_str());
    PathRemoveFileSpec(dir);
    
    std::wstring filename = PathFindFileName(path.c_str());
    std::wstring output = RunGitCommand(L"add \"" + filename + L"\"", dir);
    
    g_LastCheck = 0; // Invalidate cache
    
    if (output.empty()) output = L"Added " + filename;
    ShowGitOutput(GetParent(hEditor), output);
}

void GitDiff(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    TCHAR dir[MAX_PATH];
    wcscpy_s(dir, path.c_str());
    PathRemoveFileSpec(dir);
    
    std::wstring filename = PathFindFileName(path.c_str());
    std::wstring output = RunGitCommand(L"diff \"" + filename + L"\"", dir);
    
    if (output.find(L"did not match any file(s) known to git") != std::wstring::npos) {
        output += L"\r\n\r\n(This file is likely untracked. Use 'Git Add' first.)";
    }
    else if (output.empty()) output = L"No changes.";

    ShowGitOutput(GetParent(hEditor), output);
}

void GitBlame(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    TCHAR dir[MAX_PATH];
    wcscpy_s(dir, path.c_str());
    PathRemoveFileSpec(dir);
    
    std::wstring filename = PathFindFileName(path.c_str());
    std::wstring output = RunGitCommand(L"blame \"" + filename + L"\"", dir);
    ShowGitOutput(GetParent(hEditor), output);
}

void GitLog(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    TCHAR dir[MAX_PATH];
    wcscpy_s(dir, path.c_str());
    PathRemoveFileSpec(dir);
    
    std::wstring filename = PathFindFileName(path.c_str());
    std::wstring output = RunGitCommand(L"log -n 10 \"" + filename + L"\"", dir);
    ShowGitOutput(GetParent(hEditor), output);
}

void GitCommit(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    if (DialogBox(g_hInst, MAKEINTRESOURCE(IDD_GIT_COMMIT), GetParent(hEditor), GitCommitDlgProc) == IDOK) {
        if (g_CommitMessage.empty()) {
            MessageBox(hEditor, L"Commit message cannot be empty.", L"Git Plugin", MB_OK | MB_ICONWARNING);
            return;
        }

        TCHAR dir[MAX_PATH];
        wcscpy_s(dir, path.c_str());
        PathRemoveFileSpec(dir);
        
        std::wstring filename = PathFindFileName(path.c_str());
        
        // Escape quotes in message
        std::wstring msg = g_CommitMessage;
        // Simple escape: replace " with \"
        size_t pos = 0;
        while ((pos = msg.find(L"\"", pos)) != std::wstring::npos) {
            msg.replace(pos, 1, L"\\\"");
            pos += 2;
        }

        std::wstring output = RunGitCommand(L"commit \"" + filename + L"\" -m \"" + msg + L"\"", dir);
        g_LastCheck = 0; // Invalidate cache
        ShowGitOutput(GetParent(hEditor), output);
    }
}

void GitPull(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    TCHAR dir[MAX_PATH];
    wcscpy_s(dir, path.c_str());
    PathRemoveFileSpec(dir);

    std::wstring output = RunGitCommand(L"pull", dir);
    g_LastCheck = 0; // Invalidate cache
    ShowGitOutput(GetParent(hEditor), output);
}

void GitPush(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    TCHAR dir[MAX_PATH];
    wcscpy_s(dir, path.c_str());
    PathRemoveFileSpec(dir);

    std::wstring output = RunGitCommand(L"push", dir);
    g_LastCheck = 0; // Invalidate cache
    ShowGitOutput(GetParent(hEditor), output);
}

void GitDiscard(HWND hEditor) {
    std::wstring path = GetCurrentFilePath(hEditor);
    if (path.empty()) {
        MessageBox(hEditor, L"Please save the file first.", L"Git Plugin", MB_OK);
        return;
    }

    if (MessageBox(hEditor, L"Are you sure you want to discard changes to this file? This cannot be undone.", L"Git Plugin", MB_YESNO | MB_ICONWARNING) == IDYES) {
        TCHAR dir[MAX_PATH];
        wcscpy_s(dir, path.c_str());
        PathRemoveFileSpec(dir);
        
        std::wstring filename = PathFindFileName(path.c_str());
        std::wstring output = RunGitCommand(L"checkout \"" + filename + L"\"", dir);
        
        g_LastCheck = 0; // Invalidate cache
        
        // Reload file if successful (simple way: tell user to reload, or trigger reload if we had API)
        // Since we don't have a "Reload" API exposed to plugins easily, we'll just show output.
        // Ideally, the editor should detect file change on disk.
        ShowGitOutput(GetParent(hEditor), output.empty() ? L"Changes discarded. Please reload the file." : output);
    }
}

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Git";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Simple Git integration.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.1";
    }

    PLUGIN_API const wchar_t* GetPluginLicense() {
        return L"MIT License\n\n"
               L"Copyright (c) 2025 Just Notepad Contributors\n\n"
               L"Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\n"
               L"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\n"
               L"THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO, THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
    }

    std::wstring g_StatusCache;
    std::wstring g_LastFile;

    PLUGIN_API const wchar_t* GetPluginStatus(const wchar_t* filePath) {
        if (!filePath || !filePath[0]) return NULL;
        
        // Simple cache to avoid running git too often
        if (g_LastFile == filePath && GetTickCount() - g_LastCheck < 2000) {
            return g_StatusCache.empty() ? NULL : g_StatusCache.c_str();
        }
        
        g_LastFile = filePath;
        g_LastCheck = GetTickCount();
        g_StatusCache = L"";

        std::filesystem::path p(filePath);
        std::wstring dir = p.parent_path().wstring();
        
        // Check branch
        std::wstring branch = RunGitCommand(L"rev-parse --abbrev-ref HEAD", dir);
        if (branch.empty() || branch.find(L"fatal") != std::wstring::npos) {
            return NULL;
        }
        
        // Trim newline
        if (!branch.empty() && branch.back() == '\n') branch.pop_back();

        // Check file status
        std::wstring filename = p.filename().wstring();
        std::wstring statusCmd = L"status --porcelain \"" + filename + L"\"";
        std::wstring statusOut = RunGitCommand(statusCmd, dir);
        
        std::wstring statusStr = L"";
        if (statusOut.length() >= 1) {
             if (statusOut.find(L"??") == 0) statusStr = L" [Untracked]";
             else if (statusOut.find(L"M") != std::wstring::npos) statusStr = L" [Modified]";
             else if (statusOut.find(L"A") != std::wstring::npos) statusStr = L" [Added]";
             else statusStr = L" [Changed]";
        }

        g_StatusCache = L"Git: " + branch + statusStr;
        return g_StatusCache.c_str();
    }

    PluginMenuItem g_Items[] = {
        { L"Status", GitStatus, L"Ctrl+Shift+G" },
        { L"Add", GitAdd },
        { L"Diff", GitDiff },
        { L"Commit...", GitCommit },
        { L"Pull", GitPull },
        { L"Push", GitPush },
        { L"Discard Changes", GitDiscard },
        { L"Blame", GitBlame, L"Ctrl+Shift+B" },
        { L"Log", GitLog }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_Items) / sizeof(PluginMenuItem);
        return g_Items;
    }
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
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
