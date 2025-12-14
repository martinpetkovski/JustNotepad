#include "..\..\src\PluginInterface.h"
#include <windows.h>
#include <string>
#include <vector>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <thread>
#include <mutex>
#include <atomic>
#include <map>
#include <iostream>
#include <algorithm>
#include <set>
#include <richedit.h>
#include <commctrl.h>

namespace fs = std::filesystem;

// Custom Messages
#define WM_AC_UPDATE        (WM_USER + 100)
#define WM_AC_HIDE          (WM_USER + 101)
#define WM_AC_EXIT          (WM_USER + 102)

HMODULE g_hModule = NULL;
HostFunctions* g_hostFunctions = NULL;
std::wstring g_ctagsPath;
HWND g_hEditor = NULL;
WNDPROC g_OldEditorProc = NULL;
HWND g_hPopup = NULL; // Owned by worker thread
std::wstring g_CurrentFilePath;

struct Tag {
    std::wstring name;
    std::wstring kind;
    int line;
};

std::vector<Tag> g_AllTags;
std::vector<Tag> g_CtagsCache;
std::mutex g_TagsMutex;

int g_StartPos = 0;
std::atomic<bool> g_IsAutocompleteOpen = false;

HANDLE g_hThread = NULL;
DWORD g_dwThreadId = 0;

struct UpdateParams {
    wchar_t prefix[256];
    int x;
    int y;
    RECT rcWork;
};

void Log(const std::wstring& msg) {
    OutputDebugStringW((L"[CtagsPlugin] " + msg + L"\n").c_str());
}

void FindCtags() {
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(g_hModule, path, MAX_PATH);
    fs::path pluginDir = fs::path(path).parent_path();
    
    fs::path localCtags = pluginDir / "ctags" / "ctags.exe";
    if (fs::exists(localCtags)) {
        g_ctagsPath = localCtags.wstring();
        Log(L"Found ctags at: " + g_ctagsPath);
        return;
    }
    g_ctagsPath = L"ctags.exe"; 
}

std::vector<Tag> ParseCtagsOutput(const std::string& output) {
    std::vector<Tag> tags;
    std::istringstream iss(output);
    std::string line;
    while (std::getline(iss, line)) {
        std::istringstream ls(line);
        std::string name, kind, lineStr;
        ls >> name >> kind >> lineStr;
        
        if (!name.empty()) {
            Tag t;
            int wlen = MultiByteToWideChar(CP_UTF8, 0, name.c_str(), -1, NULL, 0);
            std::wstring wname(wlen, 0);
            MultiByteToWideChar(CP_UTF8, 0, name.c_str(), -1, &wname[0], wlen);
            t.name = wname;
            t.kind = L"ctags";
            if (!t.name.empty() && t.name.back() == 0) t.name.pop_back();
            tags.push_back(t);
        }
    }
    return tags;
}

std::vector<Tag> RunCtags(const std::wstring& content, const std::wstring& filename) {
    if (g_ctagsPath.empty()) return {};

    SECURITY_ATTRIBUTES saAttr;
    saAttr.nLength = sizeof(SECURITY_ATTRIBUTES);
    saAttr.bInheritHandle = TRUE;
    saAttr.lpSecurityDescriptor = NULL;

    HANDLE hChildStd_OUT_Rd = NULL;
    HANDLE hChildStd_OUT_Wr = NULL;

    if (!CreatePipe(&hChildStd_OUT_Rd, &hChildStd_OUT_Wr, &saAttr, 0)) return {};
    if (!SetHandleInformation(hChildStd_OUT_Rd, HANDLE_FLAG_INHERIT, 0)) return {};

    fs::path p(filename);
    std::wstring ext = p.extension().wstring();
    if (ext.empty()) ext = L".cpp"; 
    
    wchar_t tempPath[MAX_PATH];
    GetTempPathW(MAX_PATH, tempPath);
    wchar_t tempFile[MAX_PATH];
    GetTempFileNameW(tempPath, L"JNP", 0, tempFile);
    
    std::wstring tempFileWithExt = std::wstring(tempFile) + ext;
    MoveFileW(tempFile, tempFileWithExt.c_str());
    
    std::ofstream out(tempFileWithExt, std::ios::binary);
    std::string utf8Content;
    int len = WideCharToMultiByte(CP_UTF8, 0, content.c_str(), -1, NULL, 0, NULL, NULL);
    utf8Content.resize(len);
    WideCharToMultiByte(CP_UTF8, 0, content.c_str(), -1, &utf8Content[0], len, NULL, NULL);
    if (!utf8Content.empty() && utf8Content.back() == 0) utf8Content.pop_back();
    out.write(utf8Content.c_str(), utf8Content.size());
    out.close();

    std::wstring cmd = L"\"" + g_ctagsPath + L"\" -x --sort=no -f - \"" + tempFileWithExt + L"\"";

    PROCESS_INFORMATION piProcInfo;
    STARTUPINFO siStartInfo;
    ZeroMemory(&piProcInfo, sizeof(PROCESS_INFORMATION));
    ZeroMemory(&siStartInfo, sizeof(STARTUPINFO));
    siStartInfo.cb = sizeof(STARTUPINFO);
    siStartInfo.hStdError = hChildStd_OUT_Wr; 
    siStartInfo.hStdOutput = hChildStd_OUT_Wr;
    siStartInfo.dwFlags |= STARTF_USESTDHANDLES;

    if (!CreateProcessW(NULL, &cmd[0], NULL, NULL, TRUE, CREATE_NO_WINDOW, NULL, NULL, &siStartInfo, &piProcInfo)) {
        DeleteFileW(tempFileWithExt.c_str());
        return {};
    }

    CloseHandle(hChildStd_OUT_Wr);

    std::string output;
    DWORD dwRead;
    CHAR chBuf[4096];
    BOOL bSuccess = FALSE;

    for (;;) {
        bSuccess = ReadFile(hChildStd_OUT_Rd, chBuf, 4096, &dwRead, NULL);
        if (!bSuccess || dwRead == 0) break;
        output.append(chBuf, dwRead);
    }

    CloseHandle(hChildStd_OUT_Rd);
    CloseHandle(piProcInfo.hProcess);
    CloseHandle(piProcInfo.hThread);
    
    DeleteFileW(tempFileWithExt.c_str());

    return ParseCtagsOutput(output);
}

std::vector<Tag> GetBufferWords(const std::wstring& content) {
    std::vector<Tag> tags;
    std::set<std::wstring> seen;
    
    std::wstring current;
    for (wchar_t c : content) {
        if (iswalnum(c) || c == L'_') {
            current += c;
        } else {
            if (current.length() > 0) { 
                if (seen.find(current) == seen.end()) {
                    Tag t;
                    t.name = current;
                    t.kind = L"word";
                    tags.push_back(t);
                    seen.insert(current);
                }
            }
            current.clear();
        }
    }
    if (current.length() > 0 && seen.find(current) == seen.end()) { 
        Tag t;
        t.name = current;
        t.kind = L"word";
        tags.push_back(t);
    }
    return tags;
}

DWORD WINAPI AutocompleteThreadProc(LPVOID lpParam) {
    // Create Window
    g_hPopup = CreateWindowEx(WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE, L"ListBox", L"", 
        WS_POPUP | WS_BORDER | WS_VSCROLL | LBS_NOTIFY | LBS_HASSTRINGS, 
        0, 0, 200, 150, NULL, NULL, g_hModule, NULL);
    
    HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    SendMessage(g_hPopup, WM_SETFONT, (WPARAM)hFont, TRUE);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (msg.message == WM_AC_UPDATE) {
            UpdateParams* params = (UpdateParams*)msg.lParam;
            std::wstring prefix = params->prefix;
            
            std::vector<Tag> filtered;
            {
                std::lock_guard<std::mutex> lock(g_TagsMutex);
                for (const auto& t : g_AllTags) {
                    if (t.name.length() >= prefix.length() && t.name.substr(0, prefix.length()) == prefix) {
                         filtered.push_back(t);
                    }
                }
            }
            
            if (filtered.empty()) {
                ShowWindow(g_hPopup, SW_HIDE);
                g_IsAutocompleteOpen = false;
            } else {
                SendMessage(g_hPopup, LB_RESETCONTENT, 0, 0);
                for (const auto& t : filtered) {
                    SendMessage(g_hPopup, LB_ADDSTRING, 0, (LPARAM)t.name.c_str());
                }
                SendMessage(g_hPopup, LB_SETCURSEL, 0, 0);
                
                int x = params->x;
                int y = params->y;
                int width = 200;
                int height = 150;
                
                if (x + width > params->rcWork.right) x = params->rcWork.right - width;
                if (y + height > params->rcWork.bottom) y = params->y - height - 20; 
                
                SetWindowPos(g_hPopup, HWND_TOPMOST, x, y, width, height, SWP_SHOWWINDOW | SWP_NOACTIVATE);
                g_IsAutocompleteOpen = true;
            }
            
            delete params;
        }
        else if (msg.message == WM_AC_HIDE) {
            ShowWindow(g_hPopup, SW_HIDE);
            g_IsAutocompleteOpen = false;
        }
        else if (msg.message == WM_AC_EXIT) {
            DestroyWindow(g_hPopup);
            g_hPopup = NULL;
            break; 
        }
        else {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    return 0;
}

void DestroyPopup() {
    if (g_dwThreadId) {
        PostThreadMessage(g_dwThreadId, WM_AC_HIDE, 0, 0);
    }
}

void InsertSelection(HWND hEditor) {
    if (!g_hPopup) return;
    
    // Synchronous call to get selection from the thread's window
    int cur = SendMessage(g_hPopup, LB_GETCURSEL, 0, 0);
    if (cur == LB_ERR) {
        DestroyPopup();
        return;
    }
    
    int len = SendMessage(g_hPopup, LB_GETTEXTLEN, cur, 0);
    std::vector<wchar_t> text(len + 1);
    SendMessage(g_hPopup, LB_GETTEXT, cur, (LPARAM)text.data());
    
    std::wstring tag = text.data();
    
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    CHARRANGE rangeToReplace;
    rangeToReplace.cpMin = g_StartPos;
    rangeToReplace.cpMax = cr.cpMax;
    
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&rangeToReplace);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)tag.c_str());
    
    DestroyPopup();
}

void UpdatePopup(HWND hEditor) {
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    if (cr.cpMin < g_StartPos) {
        DestroyPopup();
        return;
    }
    
    TEXTRANGEW tr;
    tr.chrg.cpMin = g_StartPos;
    tr.chrg.cpMax = cr.cpMax;
    int len = tr.chrg.cpMax - tr.chrg.cpMin;
    std::vector<wchar_t> buf(len + 1);
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::wstring prefix = buf.data();
    
    if (prefix.empty()) {
        DestroyPopup();
        return;
    }
    
    POINT pt;
    GetCaretPos(&pt);
    ClientToScreen(hEditor, &pt);
    
    HMONITOR hMonitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi = { sizeof(MONITORINFO) };
    GetMonitorInfo(hMonitor, &mi);

    UpdateParams* params = new UpdateParams();
    wcsncpy_s(params->prefix, prefix.c_str(), 255);
    params->x = pt.x;
    params->y = pt.y + 20;
    params->rcWork = mi.rcWork;

    PostThreadMessage(g_dwThreadId, WM_AC_UPDATE, 0, (LPARAM)params);
}

LRESULT CALLBACK EditorSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    if (g_IsAutocompleteOpen) {
        if (uMsg == WM_KEYDOWN) {
            if (wParam == VK_UP || wParam == VK_DOWN) {
                int count = SendMessage(g_hPopup, LB_GETCOUNT, 0, 0);
                int cur = SendMessage(g_hPopup, LB_GETCURSEL, 0, 0);
                if (wParam == VK_UP) cur = (cur > 0) ? cur - 1 : 0;
                if (wParam == VK_DOWN) cur = (cur < count - 1) ? cur + 1 : count - 1;
                SendMessage(g_hPopup, LB_SETCURSEL, cur, 0);
                return 0;
            }
            if (wParam == VK_RETURN || wParam == VK_TAB) {
                InsertSelection(hWnd);
                return 0;
            }
            if (wParam == VK_ESCAPE) {
                DestroyPopup();
                return 0;
            }
        }
        if (uMsg == WM_CHAR && wParam == VK_TAB) {
            return 0;
        }
        if (uMsg == WM_KILLFOCUS) {
            if ((HWND)wParam != g_hPopup) {
                DestroyPopup();
            }
        }
        if (uMsg == WM_LBUTTONDOWN) {
            DestroyPopup();
        }
        if (uMsg == WM_MOUSEWHEEL) {
             DestroyPopup();
        }
    }

    LRESULT ret = CallWindowProc(g_OldEditorProc, hWnd, uMsg, wParam, lParam);

    if (uMsg == WM_CHAR || uMsg == WM_KEYDOWN) { 
         if (wParam != VK_UP && wParam != VK_DOWN && wParam != VK_RETURN && wParam != VK_TAB && wParam != VK_ESCAPE) {
             UpdatePopup(hWnd);
         }
    }
    return ret;
}

void RefreshCache(HWND hEditor) {
    int len = GetWindowTextLengthW(hEditor);
    std::vector<wchar_t> buf(len + 1);
    GetWindowTextW(hEditor, buf.data(), len + 1);
    std::wstring content = buf.data();

    std::wstring filename = g_CurrentFilePath;
    if (filename.empty()) filename = L"test.cpp";
    
    std::vector<Tag> tags = RunCtags(content, filename);
    std::vector<Tag> bufferTags = GetBufferWords(content);
    
    std::lock_guard<std::mutex> lock(g_TagsMutex);
    g_CtagsCache = tags;
    
    std::set<std::wstring> seen;
    g_AllTags.clear();
    for (const auto& t : g_CtagsCache) {
        g_AllTags.push_back(t);
        seen.insert(t.name);
    }
    for (const auto& t : bufferTags) {
        if (seen.find(t.name) == seen.end()) {
            g_AllTags.push_back(t);
            seen.insert(t.name);
        }
    }
    
    std::sort(g_AllTags.begin(), g_AllTags.end(), [](const Tag& a, const Tag& b) {
        return a.name < b.name;
    });
}

int GetWordStartPos(HWND hEditor, int cp) {
    if (cp == 0) return 0;
    int startRead = max(0, cp - 128);
    TEXTRANGEW tr;
    tr.chrg.cpMin = startRead;
    tr.chrg.cpMax = cp;
    wchar_t buf[130];
    tr.lpstrText = buf;
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::wstring chunk = buf;
    for (int i = chunk.length() - 1; i >= 0; --i) {
        if (!iswalnum(chunk[i]) && chunk[i] != L'_') {
            return startRead + i + 1;
        }
    }
    return startRead; 
}

void ShowSuggestions(HWND hEditor) {
    g_hEditor = hEditor;
    
    if (!g_OldEditorProc) {
        g_OldEditorProc = (WNDPROC)SetWindowLongPtr(hEditor, GWLP_WNDPROC, (LONG_PTR)EditorSubclassProc);
    }

    RefreshCache(hEditor);

    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    g_StartPos = GetWordStartPos(hEditor, cr.cpMin);

    UpdatePopup(hEditor);
}

PluginMenuItem g_items[] = {
    { L"Show Suggestions", ShowSuggestions, L"Ctrl+Space" }
};

extern "C" {

PLUGIN_API const wchar_t* GetPluginName() { return L"Generic Autocomplete (Ctags)"; }
PLUGIN_API const wchar_t* GetPluginDescription() { return L"Provides autocomplete suggestions using Universal Ctags."; }
PLUGIN_API const wchar_t* GetPluginVersion() { return L"1.0"; }

PLUGIN_API const wchar_t* GetPluginLicense() {
    return L"MIT License\n\n"
           L"Copyright (c) 2025 Just Notepad Contributors\n\n"
           L"Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\n"
           L"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\n"
           L"THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO, THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
}
PLUGIN_API const wchar_t* GetPluginStatus(const wchar_t* filePath) { return L"Ctags Ready"; }
PLUGIN_API void Initialize(const wchar_t* settingsPath) { 
    FindCtags(); 
    g_hThread = CreateThread(NULL, 0, AutocompleteThreadProc, NULL, 0, &g_dwThreadId);
}
PLUGIN_API void Shutdown() {
    if (g_hEditor && g_OldEditorProc) {
        SetWindowLongPtr(g_hEditor, GWLP_WNDPROC, (LONG_PTR)g_OldEditorProc);
        g_OldEditorProc = NULL;
    }
    if (g_dwThreadId) {
        PostThreadMessage(g_dwThreadId, WM_AC_EXIT, 0, 0);
        WaitForSingleObject(g_hThread, 1000);
        CloseHandle(g_hThread);
        g_hThread = NULL;
        g_dwThreadId = 0;
    }
}
PLUGIN_API void SetHostFunctions(HostFunctions* functions) { g_hostFunctions = functions; }
PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) { *count = 1; return g_items; }

PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
    if (filePath) g_CurrentFilePath = filePath;
    
    if (wcscmp(eventType, L"Opened") == 0 || wcscmp(eventType, L"Saved") == 0) {
        RefreshCache(hEditor);
    }
}

PLUGIN_API void OnTextModified(HWND hEditor) {
    g_hEditor = hEditor;
    
    if (!g_OldEditorProc) {
        g_OldEditorProc = (WNDPROC)SetWindowLongPtr(hEditor, GWLP_WNDPROC, (LONG_PTR)EditorSubclassProc);
    }

    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    g_StartPos = GetWordStartPos(hEditor, cr.cpMin);

    UpdatePopup(hEditor);
}

PLUGIN_API bool OnSaveFile(const wchar_t* filePath, HWND hEditor) { return false; }
PLUGIN_API long long GetMaxFileSize() { return 1024 * 1024 * 10; }

} // extern "C"

BOOL APIENTRY DllMain( HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved ) {
    switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH: g_hModule = hModule; break;
    case DLL_PROCESS_DETACH: Shutdown(); break;
    }
    return TRUE;
}
