#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <richedit.h>
#include <string>
#include <vector>
#include <algorithm>
#include <shlwapi.h>
#include "resource.h"

#pragma comment(lib, "shlwapi.lib")

HINSTANCE g_hInst = NULL;
std::wstring g_Style = L"file";
int g_IndentWidth = 4;
std::wstring g_UseTab = L"Never";
std::wstring g_SettingsPath;

extern "C" {
    PLUGIN_API void Initialize(const wchar_t* settingsPath) {
        g_SettingsPath = settingsPath;
        
        TCHAR buffer[256];
        GetPrivateProfileString(L"Settings", L"Style", L"file", buffer, 256, g_SettingsPath.c_str());
        g_Style = buffer;
        
        g_IndentWidth = GetPrivateProfileInt(L"Settings", L"IndentWidth", 4, g_SettingsPath.c_str());
        
        GetPrivateProfileString(L"Settings", L"UseTab", L"Never", buffer, 256, g_SettingsPath.c_str());
        g_UseTab = buffer;
    }

    PLUGIN_API void Shutdown() {
        if (!g_SettingsPath.empty()) {
            WritePrivateProfileString(L"Settings", L"Style", g_Style.c_str(), g_SettingsPath.c_str());
            WritePrivateProfileString(L"Settings", L"IndentWidth", std::to_wstring(g_IndentWidth).c_str(), g_SettingsPath.c_str());
            WritePrivateProfileString(L"Settings", L"UseTab", g_UseTab.c_str(), g_SettingsPath.c_str());
        }
    }
}


// Helper to ensure CRLF line endings
std::wstring NormalizeLineEndings(const std::wstring& input) {
    std::wstring output;
    output.reserve(input.length() + input.length() / 20); // Reserve some extra space
    
    for (size_t i = 0; i < input.length(); ++i) {
        if (input[i] == L'\n') {
            if (i == 0 || input[i-1] != L'\r') {
                output += L'\r';
            }
        }
        output += input[i];
    }
    return output;
}

std::wstring ExtractClangFormat() {
    TCHAR szTempPath[MAX_PATH];
    GetTempPath(MAX_PATH, szTempPath);
    
    TCHAR szExePath[MAX_PATH];
    PathCombine(szExePath, szTempPath, _T("clang-format-jn.exe"));
    
    // Check if already exists
    if (PathFileExists(szExePath)) {
        return std::wstring(szExePath);
    }
    
    // Extract resource
    HRSRC hRes = FindResource(g_hInst, MAKEINTRESOURCE(IDR_CLANG_FORMAT), _T("BINARY"));
    if (!hRes) return L"";
    
    HGLOBAL hData = LoadResource(g_hInst, hRes);
    if (!hData) return L"";
    
    DWORD dwSize = SizeofResource(g_hInst, hRes);
    LPVOID pData = LockResource(hData);
    
    HANDLE hFile = CreateFile(szExePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return L"";
    
    DWORD dwWritten;
    WriteFile(hFile, pData, dwSize, &dwWritten, NULL);
    CloseHandle(hFile);
    
    return std::wstring(szExePath);
}

std::wstring GetClangFormatPath() {
    // Try to extract from resource first
    std::wstring extracted = ExtractClangFormat();
    if (!extracted.empty()) return extracted;

    // Fallback to local file
    TCHAR szPath[MAX_PATH];
    if (GetModuleFileName(g_hInst, szPath, MAX_PATH)) {
        PathRemoveFileSpec(szPath);
        PathAppend(szPath, _T("clang-format.exe"));
        if (PathFileExists(szPath)) {
            return std::wstring(szPath);
        }
    }
    return L"clang-format"; // Fallback to PATH
}

// Helper to get file path from main window
std::wstring GetCurrentFilePath(HWND hEditor) {
    HWND hMain = GetParent(hEditor);
    HANDLE hProp = GetProp(hMain, L"FullPath");
    if (hProp) {
        TCHAR* path = (TCHAR*)hProp;
        if (path && path[0]) {
            return std::wstring(path);
        }
    }
    return L"test.cpp"; // Default fallback
}

INT_PTR CALLBACK SettingsDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        {
            HWND hCombo = GetDlgItem(hDlg, IDC_STYLE_COMBO);
            SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"file");
            SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"LLVM");
            SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Google");
            SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Chromium");
            SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Mozilla");
            SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"WebKit");
            SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)L"Microsoft");
            
            int index = SendMessage(hCombo, CB_FINDSTRINGEXACT, -1, (LPARAM)g_Style.c_str());
            if (index != CB_ERR) {
                SendMessage(hCombo, CB_SETCURSEL, index, 0);
            } else {
                SendMessage(hCombo, CB_SETCURSEL, 0, 0);
            }

            SetDlgItemInt(hDlg, IDC_INDENT_WIDTH, g_IndentWidth, FALSE);

            HWND hTabCombo = GetDlgItem(hDlg, IDC_USE_TAB_COMBO);
            SendMessage(hTabCombo, CB_ADDSTRING, 0, (LPARAM)L"Never");
            SendMessage(hTabCombo, CB_ADDSTRING, 0, (LPARAM)L"Always");
            SendMessage(hTabCombo, CB_ADDSTRING, 0, (LPARAM)L"ForIndentation");
            SendMessage(hTabCombo, CB_ADDSTRING, 0, (LPARAM)L"ForContinuationAndIndentation");
            
            index = SendMessage(hTabCombo, CB_FINDSTRINGEXACT, -1, (LPARAM)g_UseTab.c_str());
            if (index != CB_ERR) {
                SendMessage(hTabCombo, CB_SETCURSEL, index, 0);
            } else {
                SendMessage(hTabCombo, CB_SETCURSEL, 0, 0);
            }
        }
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            HWND hCombo = GetDlgItem(hDlg, IDC_STYLE_COMBO);
            int index = SendMessage(hCombo, CB_GETCURSEL, 0, 0);
            if (index != CB_ERR) {
                wchar_t buffer[256];
                SendMessage(hCombo, CB_GETLBTEXT, index, (LPARAM)buffer);
                g_Style = buffer;
            }

            g_IndentWidth = GetDlgItemInt(hDlg, IDC_INDENT_WIDTH, NULL, FALSE);

            HWND hTabCombo = GetDlgItem(hDlg, IDC_USE_TAB_COMBO);
            index = SendMessage(hTabCombo, CB_GETCURSEL, 0, 0);
            if (index != CB_ERR) {
                wchar_t buffer[256];
                SendMessage(hTabCombo, CB_GETLBTEXT, index, (LPARAM)buffer);
                g_UseTab = buffer;
            }

            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        } else if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void ShowSettings(HWND hParent) {
    DialogBox(g_hInst, MAKEINTRESOURCE(IDD_SETTINGS_DIALOG), hParent, SettingsDialogProc);
}

// Helper to run clang-format
std::wstring RunClangFormat(const std::wstring& input, const std::wstring& filename) {
    // Ensure we send \n or \r\n, not just \r (RichEdit uses \r internally)
    std::wstring normalizedInput;
    normalizedInput.reserve(input.length());
    for (size_t i = 0; i < input.length(); ++i) {
        if (input[i] == L'\r') {
            if (i + 1 < input.length() && input[i+1] == L'\n') {
                normalizedInput += L"\r\n";
                i++;
            } else {
                normalizedInput += L"\n";
            }
        } else {
            normalizedInput += input[i];
        }
    }

    // Convert input to UTF-8 for clang-format
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, normalizedInput.c_str(), (int)normalizedInput.length(), NULL, 0, NULL, NULL);
    std::string inputUtf8(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, normalizedInput.c_str(), (int)normalizedInput.length(), &inputUtf8[0], size_needed, NULL, NULL);

    HANDLE hStdInPipeRead = NULL;
    HANDLE hStdInPipeWrite = NULL;
    HANDLE hStdOutPipeRead = NULL;
    HANDLE hStdOutPipeWrite = NULL;

    SECURITY_ATTRIBUTES saAttr;
    saAttr.nLength = sizeof(SECURITY_ATTRIBUTES);
    saAttr.bInheritHandle = TRUE;
    saAttr.lpSecurityDescriptor = NULL;

    if (!CreatePipe(&hStdOutPipeRead, &hStdOutPipeWrite, &saAttr, 0)) return input;
    if (!SetHandleInformation(hStdOutPipeRead, HANDLE_FLAG_INHERIT, 0)) return input;

    if (!CreatePipe(&hStdInPipeRead, &hStdInPipeWrite, &saAttr, 0)) return input;
    if (!SetHandleInformation(hStdInPipeWrite, HANDLE_FLAG_INHERIT, 0)) return input;

    STARTUPINFO si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(STARTUPINFO));
    si.cb = sizeof(STARTUPINFO);
    si.hStdError = hStdOutPipeWrite;
    si.hStdOutput = hStdOutPipeWrite;
    si.hStdInput = hStdInPipeRead;
    si.dwFlags |= STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    ZeroMemory(&pi, sizeof(PROCESS_INFORMATION));

    std::wstring exePath = GetClangFormatPath();
    std::wstring cmdLineStr;
    
    if (g_Style == L"file") {
        cmdLineStr = L"\"" + exePath + L"\" -style=file";
    } else {
        cmdLineStr = L"\"" + exePath + L"\" -style=\"{BasedOnStyle: " + g_Style + 
                     L", IndentWidth: " + std::to_wstring(g_IndentWidth) + 
                     L", UseTab: " + g_UseTab + L"}\"";
    }
    
    if (!filename.empty()) {
        cmdLineStr += L" -assume-filename=\"" + filename + L"\"";
    }

    // Create a mutable buffer for CreateProcess
    std::vector<wchar_t> cmdLine(cmdLineStr.begin(), cmdLineStr.end());
    cmdLine.push_back(0);

    if (!CreateProcess(NULL, cmdLine.data(), NULL, NULL, TRUE, 0, NULL, NULL, &si, &pi)) {
        CloseHandle(hStdOutPipeRead);
        CloseHandle(hStdOutPipeWrite);
        CloseHandle(hStdInPipeRead);
        CloseHandle(hStdInPipeWrite);
        
        std::wstring msg = L"Could not start clang-format.\nPlease ensure 'clang-format.exe' is in the plugins folder or in your PATH.\n\nTried: " + exePath;
        MessageBox(NULL, msg.c_str(), L"Clang-Format Plugin", MB_OK | MB_ICONERROR);
        return input;
    }

    CloseHandle(hStdOutPipeWrite);
    CloseHandle(hStdInPipeRead);

    // Write to stdin
    DWORD dwWritten;
    WriteFile(hStdInPipeWrite, inputUtf8.c_str(), (DWORD)inputUtf8.length(), &dwWritten, NULL);
    CloseHandle(hStdInPipeWrite);

    // Read from stdout
    std::string outputUtf8;
    DWORD dwRead;
    CHAR chBuf[4096];
    BOOL bSuccess = FALSE;

    for (;;) {
        bSuccess = ReadFile(hStdOutPipeRead, chBuf, 4096, &dwRead, NULL);
        if (!bSuccess || dwRead == 0) break;
        outputUtf8.append(chBuf, dwRead);
    }

    CloseHandle(hStdOutPipeRead);
    WaitForSingleObject(pi.hProcess, INFINITE);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    // Convert back to WideChar
    if (outputUtf8.empty()) return input; 

    int wsize_needed = MultiByteToWideChar(CP_UTF8, 0, outputUtf8.c_str(), (int)outputUtf8.length(), NULL, 0);
    std::wstring output(wsize_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, outputUtf8.c_str(), (int)outputUtf8.length(), &output[0], wsize_needed);

    return NormalizeLineEndings(output);
}

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Clang-Format";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Formats selected text using clang-format.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.0";
    }

    bool g_bIsBinary = false;

    void FormatDocument(HWND hEditorWnd) {
        if (g_bIsBinary) return;
        
        // Get full text
        int len = GetWindowTextLength(hEditorWnd);
        if (len <= 0) return;

        std::vector<wchar_t> buffer(len + 1);
        GetWindowText(hEditorWnd, buffer.data(), len + 1);
        std::wstring text(buffer.data());

        std::wstring filename = GetCurrentFilePath(hEditorWnd);
        std::wstring formatted = RunClangFormat(text, filename);
        
        if (formatted != text && !formatted.empty()) {
            // Save scroll and selection
            POINT pt;
            SendMessage(hEditorWnd, EM_GETSCROLLPOS, 0, (LPARAM)&pt);
            
            CHARRANGE cr;
            SendMessage(hEditorWnd, EM_EXGETSEL, 0, (LPARAM)&cr);

            // Replace all text
            SetWindowText(hEditorWnd, formatted.c_str());

            // Restore scroll and selection
            SendMessage(hEditorWnd, EM_SETSCROLLPOS, 0, (LPARAM)&pt);
            SendMessage(hEditorWnd, EM_EXSETSEL, 0, (LPARAM)&cr);
        }
    }

    PluginMenuItem g_Items[] = {
        { L"Format Document", FormatDocument, L"Ctrl+Shift+I" },
        { L"Clang Format Settings", ShowSettings, NULL }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_Items) / sizeof(PluginMenuItem);
        return g_Items;
    }

    PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
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
