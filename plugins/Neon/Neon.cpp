#include "../../src/PluginInterface.h"
#include "resource.h"
#include <windows.h>
#include <string>
#include <vector>
#include <richedit.h>
#include <richole.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <fstream>
#include <codecvt>
#include <atomic>
#include <thread>
#include <mutex>
#include <algorithm>
#include <tom.h>
#include <ole2.h>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "ole32.lib")

// {8CC497C0-A1DF-11CE-8098-00AA0047BE5D}
static const GUID IID_ITextDocument = 
{ 0x8cc497c0, 0xa1df, 0x11ce, { 0x80, 0x98, 0x00, 0xaa, 0x00, 0x47, 0xbe, 0x5d } };

std::wstring g_CurrentFilePath;
HINSTANCE g_hInst = NULL;
bool g_bAutoHighlight = true;
HWND g_hEditor = NULL;
HostFunctions* g_HostFunctions = nullptr;

// Debounce timer
UINT_PTR g_TimerId = 0;
const UINT TIMER_DELAY = 16; // 16ms delay (approx 60fps)

// Threading and synchronization
std::atomic<int> g_TextVersion(0);
std::atomic<bool> g_IsUpdating(false);
bool g_bIsBinary = false;

// Settings
std::wstring g_Style = L"github";
std::wstring g_Font = L"Consolas";
std::wstring g_FontSize = L"10";

std::wstring g_SettingsPath;

std::wstring GetDllDir() {
    WCHAR path[MAX_PATH];
    GetModuleFileNameW(g_hInst, path, MAX_PATH);
    PathRemoveFileSpecW(path);
    return std::wstring(path);
}

void LoadSettings() {
    if (g_SettingsPath.empty()) return;
    
    WCHAR buf[256];
    GetPrivateProfileStringW(L"Settings", L"Style", L"github", buf, 256, g_SettingsPath.c_str());
    g_Style = buf;
    
    GetPrivateProfileStringW(L"Settings", L"Font", L"Consolas", buf, 256, g_SettingsPath.c_str());
    g_Font = buf;

    GetPrivateProfileStringW(L"Settings", L"FontSize", L"10", buf, 256, g_SettingsPath.c_str());
    g_FontSize = buf;
}

void SaveSettings() {
    if (g_SettingsPath.empty()) return;
    
    WritePrivateProfileStringW(L"Settings", L"Style", g_Style.c_str(), g_SettingsPath.c_str());
    WritePrivateProfileStringW(L"Settings", L"Font", g_Font.c_str(), g_SettingsPath.c_str());
    WritePrivateProfileStringW(L"Settings", L"FontSize", g_FontSize.c_str(), g_SettingsPath.c_str());
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved) {
    switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        break;
    }
    return TRUE;
}

// Callback for streaming text out of RichEdit
DWORD CALLBACK EditStreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb) {
    std::string* pStr = (std::string*)dwCookie;
    pStr->append((char*)pbBuff, cb);
    *pcb = cb;
    return 0;
}

// Callback for streaming text into RichEdit
struct StreamInCookie {
    const std::string* pStr;
    size_t currentPos;
};

DWORD CALLBACK EditStreamInCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb) {
    StreamInCookie* cookie = (StreamInCookie*)dwCookie;
    size_t remaining = cookie->pStr->length() - cookie->currentPos;
    LONG toCopy = (cb < (LONG)remaining) ? cb : (LONG)remaining;
    
    if (toCopy > 0) {
        memcpy(pbBuff, cookie->pStr->c_str() + cookie->currentPos, toCopy);
        cookie->currentPos += toCopy;
    }
    *pcb = toCopy;
    return 0;
}

#define WM_APPLY_HIGHLIGHT (WM_USER + 1)
HWND g_hWorkerWnd = NULL;

LRESULT CALLBACK WorkerWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_APPLY_HIGHLIGHT) {
        std::string* pRtf = (std::string*)lParam;
        int version = (int)wParam;
        if (g_hEditor && pRtf && version == g_TextVersion.load()) {
            g_IsUpdating = true;

            // Preserve selection and scroll position
            CHARRANGE cr;
            SendMessage(g_hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
            POINT pt;
            SendMessage(g_hEditor, EM_GETSCROLLPOS, 0, (LPARAM)&pt);

            // Disable redraw to prevent flickering
            SendMessage(g_hEditor, WM_SETREDRAW, FALSE, 0);

            // Disable events to prevent EN_CHANGE and other notifications
            DWORD eventMask = SendMessage(g_hEditor, EM_GETEVENTMASK, 0, 0);
            SendMessage(g_hEditor, EM_SETEVENTMASK, 0, eventMask & ~ENM_CHANGE);

            IRichEditOle* pOle = NULL;
            SendMessage(g_hEditor, EM_GETOLEINTERFACE, 0, (LPARAM)&pOle);
            ITextDocument* pDoc = NULL;
            if (pOle) {
                pOle->QueryInterface(IID_ITextDocument, (void**)&pDoc);
                pOle->Release();
            }

            if (pDoc) pDoc->Undo(tomSuspend, NULL);

            // Select all to prevent clearing undo stack
            CHARRANGE all = { 0, -1 };
            SendMessage(g_hEditor, EM_EXSETSEL, 0, (LPARAM)&all);

            // Check for trailing newline in original text
            bool originalEndsWithNewline = false;
            GETTEXTLENGTHEX gtl = { GTL_DEFAULT, 1200 }; 
            int textLen = SendMessage(g_hEditor, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 0);
            if (textLen > 0) {
                wchar_t buff[2] = {0};
                TEXTRANGE tr;
                tr.chrg.cpMin = textLen - 1;
                tr.chrg.cpMax = textLen;
                tr.lpstrText = buff;
                SendMessage(g_hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                if (buff[0] == L'\r' || buff[0] == L'\n') originalEndsWithNewline = true;
            }

            StreamInCookie cookie;
            cookie.pStr = pRtf;
            cookie.currentPos = 0;
            
            EDITSTREAM esIn = { 0 };
            esIn.dwCookie = (DWORD_PTR)&cookie;
            esIn.pfnCallback = EditStreamInCallback;

            SendMessage(g_hEditor, EM_STREAMIN, SF_RTF | SFF_SELECTION, (LPARAM)&esIn);

            // Remove extra newline if added
            int newTextLen = SendMessage(g_hEditor, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 0);
            if (newTextLen > 0 && !originalEndsWithNewline) {
                wchar_t buff[2] = {0};
                TEXTRANGE tr;
                tr.chrg.cpMin = newTextLen - 1;
                tr.chrg.cpMax = newTextLen;
                tr.lpstrText = buff;
                SendMessage(g_hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                if (buff[0] == L'\r' || buff[0] == L'\n') {
                    CHARRANGE delRange = { newTextLen - 1, newTextLen };
                    SendMessage(g_hEditor, EM_EXSETSEL, 0, (LPARAM)&delRange);
                    SendMessage(g_hEditor, EM_REPLACESEL, TRUE, (LPARAM)L"");
                }
            }

            if (pDoc) {
                pDoc->Undo(tomResume, NULL);
                pDoc->Release();
            }

            // Restore selection and scroll
            SendMessage(g_hEditor, EM_EXSETSEL, 0, (LPARAM)&cr);
            SendMessage(g_hEditor, EM_SETSCROLLPOS, 0, (LPARAM)&pt);
            
            // Restore event mask
            SendMessage(g_hEditor, EM_SETEVENTMASK, 0, eventMask);

            SendMessage(g_hEditor, WM_SETREDRAW, TRUE, 0);
            InvalidateRect(g_hEditor, NULL, TRUE);

            g_IsUpdating = false;
        }
        delete pRtf;
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

void EnsureWorkerWindow() {
    if (g_hWorkerWnd) return;
    
    WNDCLASSEX wc = {0};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.lpfnWndProc = WorkerWndProc;
    wc.hInstance = g_hInst;
    wc.lpszClassName = L"NeonWorkerWnd";
    RegisterClassEx(&wc);
    
    g_hWorkerWnd = CreateWindowEx(0, L"NeonWorkerWnd", L"NeonWorker", 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, g_hInst, NULL);
}

void HighlightWorker(std::string editorText, HWND hEditor, std::wstring dllDir, int startVersion, std::wstring style, std::wstring font, std::wstring fontSize, std::wstring sourceFilePath, std::wstring fileExtension) {
    CoInitializeEx(NULL, COINIT_MULTITHREADED);

    if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(10);

    // Check version before starting heavy work
    if (g_TextVersion.load() != startVersion) {
        if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(-1);
        CoUninitialize();
        return;
    }

    std::wstring inputFileName;
    bool bUseTemp = false;

    if (!editorText.empty()) {
        // 2. Save to temp file
        WCHAR tempPath[MAX_PATH];
        GetTempPathW(MAX_PATH, tempPath);
        
        // Create a unique temp file name, preserving extension if possible
        inputFileName = std::wstring(tempPath) + L"jnp_highlight_temp_" + std::to_wstring(GetCurrentThreadId());
        if (!fileExtension.empty()) {
            inputFileName += fileExtension;
        } else {
            inputFileName += L".txt";
        }

        // Write text to temp file
        {
            std::ofstream outFile(inputFileName, std::ios::binary);
            outFile.write(editorText.c_str(), editorText.size());
        }
        bUseTemp = true;
    } else if (!sourceFilePath.empty()) {
        inputFileName = sourceFilePath;
    } else {
        if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(-1);
        CoUninitialize();
        return;
    }

    if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(30);

    // 3. Run highlight.exe
    std::wstring exePath = dllDir + L"\\highlight\\highlight.exe";
    
    // Check if exe exists
    if (!PathFileExistsW(exePath.c_str())) {
        if (bUseTemp) DeleteFileW(inputFileName.c_str());
        if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(-1);
        CoUninitialize();
        return;
    }

    // Construct command line: highlight.exe -i <tempFile> --out-format=rtf --style=<style> --font=<font> --font-size=<size> --include-style
    std::wstring cmdLine = L"\"" + exePath + L"\" -i \"" + inputFileName + L"\" --out-format=rtf --style=\"" + style + L"\" --font=\"" + font + L"\" --font-size=" + fontSize + L" --include-style";

    SECURITY_ATTRIBUTES saAttr;
    saAttr.nLength = sizeof(SECURITY_ATTRIBUTES);
    saAttr.bInheritHandle = TRUE;
    saAttr.lpSecurityDescriptor = NULL;

    HANDLE hChildStd_OUT_Rd = NULL;
    HANDLE hChildStd_OUT_Wr = NULL;

    if (!CreatePipe(&hChildStd_OUT_Rd, &hChildStd_OUT_Wr, &saAttr, 0)) {
        if (bUseTemp) DeleteFileW(inputFileName.c_str());
        if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(-1);
        CoUninitialize();
        return;
    }
    SetHandleInformation(hChildStd_OUT_Rd, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW si = { 0 };
    si.cb = sizeof(STARTUPINFOW);
    si.hStdError = hChildStd_OUT_Wr;
    si.hStdOutput = hChildStd_OUT_Wr;
    si.dwFlags |= STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    PROCESS_INFORMATION pi = { 0 };

    if (CreateProcessW(NULL, &cmdLine[0], NULL, NULL, TRUE, 0, NULL, NULL, &si, &pi)) {
        CloseHandle(hChildStd_OUT_Wr); // Close write end in parent

        if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(50);

        // Read output
        std::string rtfOutput;
        DWORD dwRead;
        CHAR chBuf[4096];
        BOOL bSuccess = FALSE;
        
        // Estimate total size (rough guess: 5x input size for RTF)
        size_t estimatedSize = 0;
        if (!editorText.empty()) estimatedSize = editorText.size() * 5;
        else {
            // Get file size
            WIN32_FILE_ATTRIBUTE_DATA fad;
            if (GetFileAttributesExW(inputFileName.c_str(), GetFileExInfoStandard, &fad)) {
                estimatedSize = ((size_t)fad.nFileSizeLow) * 5;
            }
        }
        if (estimatedSize == 0) estimatedSize = 1000;

        for (;;) {
            bSuccess = ReadFile(hChildStd_OUT_Rd, chBuf, 4096, &dwRead, NULL);
            if (!bSuccess || dwRead == 0) break;
            rtfOutput.append(chBuf, dwRead);
            
            // Update progress (50% to 90%)
            if (g_HostFunctions && g_HostFunctions->SetProgress) {
                int p = 50 + (int)((rtfOutput.size() * 40) / estimatedSize);
                if (p > 89) p = 89;
                g_HostFunctions->SetProgress(p);
            }
        }

        WaitForSingleObject(pi.hProcess, INFINITE);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        CloseHandle(hChildStd_OUT_Rd);

        if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(90);

        // Check version again before applying
        if (g_TextVersion.load() == startVersion && !rtfOutput.empty()) {
            std::string* pRtf = new std::string(rtfOutput);
            PostMessage(g_hWorkerWnd, WM_APPLY_HIGHLIGHT, (WPARAM)startVersion, (LPARAM)pRtf);
        }

    } else {
        CloseHandle(hChildStd_OUT_Wr);
        CloseHandle(hChildStd_OUT_Rd);
    }

    if (g_HostFunctions && g_HostFunctions->SetProgress) g_HostFunctions->SetProgress(-1);

    CoUninitialize();
    if (bUseTemp) {
        DeleteFileW(inputFileName.c_str());
    }
}

void HighlightCode(HWND hEditor, const wchar_t* filePath = nullptr) {
    // 1. Get text from editor (UI Thread) if no file path provided
    std::string editorText;
    if (!filePath) {
        EDITSTREAM es = { 0 };
        es.dwCookie = (DWORD_PTR)&editorText;
        es.pfnCallback = EditStreamOutCallback;
        SendMessage(hEditor, EM_STREAMOUT, SF_TEXT, (LPARAM)&es);

        if (editorText.empty()) {
            return;
        }
    }

    int currentVersion = g_TextVersion.load();
    std::wstring dllDir = GetDllDir();

    std::wstring ext;
    if (!g_CurrentFilePath.empty()) {
        LPCWSTR pExt = PathFindExtensionW(g_CurrentFilePath.c_str());
        if (pExt) ext = pExt;
    }

    // Start worker thread
    std::thread worker(HighlightWorker, editorText, hEditor, dllDir, currentVersion, g_Style, g_Font, g_FontSize, filePath ? std::wstring(filePath) : std::wstring(), ext);
    worker.detach();
}

// Settings Dialog
INT_PTR CALLBACK SettingsDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG: {
        // Populate Themes
        HWND hThemeCombo = GetDlgItem(hDlg, IDC_THEME_COMBO);
        std::wstring dllDir = GetDllDir();
        std::wstring themesDir = dllDir + L"\\highlight\\themes\\";
        
        std::vector<std::wstring> themes;
        
        // Root themes
        WIN32_FIND_DATAW ffd;
        HANDLE hFind = FindFirstFileW((themesDir + L"*.theme").c_str(), &ffd);
        if (hFind != INVALID_HANDLE_VALUE) {
            do {
                if (!(ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
                    std::wstring fileName = ffd.cFileName;
                    size_t lastDot = fileName.find_last_of(L".");
                    if (lastDot != std::wstring::npos) {
                        themes.push_back(fileName.substr(0, lastDot));
                    }
                }
            } while (FindNextFileW(hFind, &ffd) != 0);
            FindClose(hFind);
        }

        // Base16 themes
        hFind = FindFirstFileW((themesDir + L"base16\\*.theme").c_str(), &ffd);
        if (hFind != INVALID_HANDLE_VALUE) {
            do {
                if (!(ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
                    std::wstring fileName = ffd.cFileName;
                    size_t lastDot = fileName.find_last_of(L".");
                    if (lastDot != std::wstring::npos) {
                        themes.push_back(L"base16\\" + fileName.substr(0, lastDot));
                    }
                }
            } while (FindNextFileW(hFind, &ffd) != 0);
            FindClose(hFind);
        }

        std::sort(themes.begin(), themes.end());
        for (const auto& theme : themes) {
            SendMessageW(hThemeCombo, CB_ADDSTRING, 0, (LPARAM)theme.c_str());
        }
        SendMessageW(hThemeCombo, CB_SELECTSTRING, -1, (LPARAM)g_Style.c_str());

        // Populate Fonts
        HWND hFontCombo = GetDlgItem(hDlg, IDC_FONT_COMBO);
        const wchar_t* fonts[] = { L"Consolas", L"Courier New", L"Lucida Console", L"Cascadia Code", L"Fira Code", L"JetBrains Mono", L"Source Code Pro" };
        for (const auto& font : fonts) {
            SendMessageW(hFontCombo, CB_ADDSTRING, 0, (LPARAM)font);
        }
        SetWindowTextW(hFontCombo, g_Font.c_str());

        // Populate Sizes
        HWND hSizeCombo = GetDlgItem(hDlg, IDC_FONT_SIZE_COMBO);
        const wchar_t* sizes[] = { L"8", L"9", L"10", L"11", L"12", L"14", L"16", L"18", L"20", L"24" };
        for (const auto& size : sizes) {
            SendMessageW(hSizeCombo, CB_ADDSTRING, 0, (LPARAM)size);
        }
        SetWindowTextW(hSizeCombo, g_FontSize.c_str());

        return (INT_PTR)TRUE;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_BTN_SAVE) {
            WCHAR buf[256];
            GetDlgItemTextW(hDlg, IDC_THEME_COMBO, buf, 256);
            g_Style = buf;
            GetDlgItemTextW(hDlg, IDC_FONT_COMBO, buf, 256);
            g_Font = buf;
            GetDlgItemTextW(hDlg, IDC_FONT_SIZE_COMBO, buf, 256);
            g_FontSize = buf;
            SaveSettings();
            EndDialog(hDlg, LOWORD(wParam));
            
            // Trigger re-highlight
            if (g_hEditor) HighlightCode(g_hEditor);
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDC_BTN_CANCEL || LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void ShowSettingsDialog(HWND hParent) {
    DialogBox(g_hInst, MAKEINTRESOURCE(IDD_SETTINGS_DLG), hParent, SettingsDlgProc);
}

VOID CALLBACK TimerProc(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime) {
    KillTimer(NULL, idEvent);
    g_TimerId = 0;
    if (g_bAutoHighlight && g_hEditor) {
        HighlightCode(g_hEditor);
    }
}

extern "C" {
    PLUGIN_API void Initialize(const wchar_t* settingsPath) {
        g_SettingsPath = settingsPath;
        LoadSettings();
    }
    
    PLUGIN_API void Shutdown() {
        SaveSettings();
    }

    PLUGIN_API const wchar_t* GetPluginName() { return L"Neon"; }
    PLUGIN_API const wchar_t* GetPluginDescription() { return L"Highlights code using Andre Simon's Highlight tool."; }
    PLUGIN_API const wchar_t* GetPluginVersion() { return L"1.4"; }

    PLUGIN_API const wchar_t* GetPluginLicense() {
        return L"MIT License\n\n"
               L"Copyright (c) 2025 Just Notepad Contributors\n\n"
               L"Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\n"
               L"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\n"
               L"THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO, THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
    }

    PLUGIN_API const wchar_t* GetPluginStatus(const wchar_t* filePath) {
        return L"Neon: Active";
    }

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        static PluginMenuItem items[] = {
            { L"Neon Settings...", (PluginCallback)ShowSettingsDialog, NULL }
        };
        *count = 1;
        return items;
    }

    PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
        g_hEditor = hEditor;
        EnsureWorkerWindow();
        if (filePath) {
            g_CurrentFilePath = filePath;
        }
        
        if (wcscmp(eventType, L"Loaded") == 0) {
            g_bIsBinary = false;
            // Check if binary
            HANDLE hFile = CreateFile(filePath, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
            if (hFile != INVALID_HANDLE_VALUE) {
                BYTE buf[1024];
                DWORD read;
                if (ReadFile(hFile, buf, sizeof(buf), &read, NULL) && read > 0) {
                    for (DWORD i = 0; i < read; i++) {
                        BYTE b = buf[i];
                        if (b < 0x20 && b != 0x09 && b != 0x0A && b != 0x0D) {
                            g_bIsBinary = true;
                            break;
                        }
                    }
                }
                CloseHandle(hFile);
            }
        }

        if (g_bIsBinary) return;

        if (wcscmp(eventType, L"Loaded") == 0 || wcscmp(eventType, L"Saved") == 0) {
            if (g_bAutoHighlight) {
                HighlightCode(hEditor, filePath);
            }
        }
    }

    PLUGIN_API void OnTextModified(HWND hEditor) {
        if (g_bIsBinary) return;
        g_hEditor = hEditor;
        if (g_bAutoHighlight) {
            if (g_IsUpdating) return;
            g_TextVersion++;

            // Debounce
            if (g_TimerId) {
                KillTimer(NULL, g_TimerId);
            }
            g_TimerId = SetTimer(NULL, 0, TIMER_DELAY, TimerProc);
        }
    }

    PLUGIN_API void SetHostFunctions(HostFunctions* functions) {
        g_HostFunctions = functions;
    }

    PLUGIN_API long long GetMaxFileSize() {
        return 20 * 1024 * 1024; // 20 MB
    }
}


