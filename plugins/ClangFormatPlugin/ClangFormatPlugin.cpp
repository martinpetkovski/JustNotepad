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

// Helper to run clang-format
std::wstring RunClangFormat(const std::wstring& input, const std::wstring& filename) {
    // Convert input to UTF-8 for clang-format
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, input.c_str(), (int)input.length(), NULL, 0, NULL, NULL);
    std::string inputUtf8(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, input.c_str(), (int)input.length(), &inputUtf8[0], size_needed, NULL, NULL);

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
    std::wstring cmdLineStr = L"\"" + exePath + L"\" -style=Google";
    
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

    void FormatSelection(HWND hEditorWnd) {
        CHARRANGE cr;
        SendMessage(hEditorWnd, EM_EXGETSEL, 0, (LPARAM)&cr);
        
        bool bSelectAll = false;
        if (cr.cpMin == cr.cpMax) {
            // No selection, select all
            CHARRANGE crAll = { 0, -1 };
            SendMessage(hEditorWnd, EM_EXSETSEL, 0, (LPARAM)&crAll);
            SendMessage(hEditorWnd, EM_EXGETSEL, 0, (LPARAM)&cr); // Update range
            bSelectAll = true;
        }

        // Get text
        int len = cr.cpMax - cr.cpMin;
        if (len <= 0) return;

        // EM_GETTEXTRANGE expects TEXTRANGE struct
        TEXTRANGE tr;
        tr.chrg = cr;
        std::vector<wchar_t> buffer(len + 1);
        tr.lpstrText = buffer.data();
        
        SendMessage(hEditorWnd, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
        
        std::wstring text(buffer.data());
        std::wstring filename = GetCurrentFilePath(hEditorWnd);
        std::wstring formatted = RunClangFormat(text, filename);
        
        if (formatted != text && !formatted.empty()) {
            SendMessage(hEditorWnd, EM_REPLACESEL, TRUE, (LPARAM)formatted.c_str());
        }
        
        // If we selected all automatically, maybe we should deselect? 
        // But usually formatting leaves it selected or cursor at end.
        // Let's leave it as is, consistent with "Format Document".
    }

    PluginMenuItem g_Items[] = {
        { L"Format Selection", FormatSelection }
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
