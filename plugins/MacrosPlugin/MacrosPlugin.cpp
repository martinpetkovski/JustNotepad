#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <strsafe.h>
#include <vector>
#include <string>
#include <shlwapi.h>
#include <filesystem>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")

namespace fs = std::filesystem;

std::wstring g_CurrentFile;
std::vector<std::wstring> g_MacroFiles;
std::vector<PluginMenuItem> g_MenuItems;
std::vector<std::wstring> g_MenuNames;
std::vector<std::wstring> g_Shortcuts;

void ExecuteMacro(int index, HWND hEditor) {
    if (index < 0 || index >= g_MacroFiles.size()) return;

    TCHAR szExePath[MAX_PATH];
    GetModuleFileName(NULL, szExePath, MAX_PATH);
    PathRemoveFileSpec(szExePath);
    
    std::wstring macroPath = std::wstring(szExePath) + L"\\macros\\" + g_MacroFiles[index];
    
    // Construct command: powershell.exe -ExecutionPolicy Bypass -File "macroPath" "currentFile"
    
    std::wstring params = L"-ExecutionPolicy Bypass -File \"" + macroPath + L"\"";
    if (!g_CurrentFile.empty()) {
        params += L" \"" + g_CurrentFile + L"\"";
    } else {
        params += L" \"\"";
    }

    // Debug: Show what we are executing
    // MessageBox(hEditor, params.c_str(), L"Debug Macro", MB_OK);

    ShellExecute(NULL, L"open", L"powershell.exe", params.c_str(), NULL, SW_HIDE);
}

template<int N>
void MacroCallback(HWND hEditor) {
    ExecuteMacro(N, hEditor);
}

#define MACRO_CALLBACK(N) MacroCallback<N>
PluginCallback g_Callbacks[] = {
    MACRO_CALLBACK(0), MACRO_CALLBACK(1), MACRO_CALLBACK(2), MACRO_CALLBACK(3), MACRO_CALLBACK(4),
    MACRO_CALLBACK(5), MACRO_CALLBACK(6), MACRO_CALLBACK(7), MACRO_CALLBACK(8), MACRO_CALLBACK(9),
    MACRO_CALLBACK(10), MACRO_CALLBACK(11), MACRO_CALLBACK(12), MACRO_CALLBACK(13), MACRO_CALLBACK(14),
    MACRO_CALLBACK(15), MACRO_CALLBACK(16), MACRO_CALLBACK(17), MACRO_CALLBACK(18), MACRO_CALLBACK(19)
};
const int MAX_MACROS = 20;

extern "C" {

PLUGIN_API const TCHAR* GetPluginName() {
    return _T("Macros");
}

PLUGIN_API const TCHAR* GetPluginDescription() {
    return _T("Executes PowerShell macros from the 'macros' folder.");
}

PLUGIN_API const TCHAR* GetPluginVersion() {
    return _T("1.0");
}

PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
    if (filePath) {
        g_CurrentFile = filePath;
    } else {
        g_CurrentFile.clear();
    }
}

PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
    g_MacroFiles.clear();
    g_MenuItems.clear();
    g_MenuNames.clear();
    g_Shortcuts.clear();

    TCHAR szExePath[MAX_PATH];
    GetModuleFileName(NULL, szExePath, MAX_PATH);
    PathRemoveFileSpec(szExePath);
    std::wstring macrosDir = std::wstring(szExePath) + L"\\macros";

    if (fs::exists(macrosDir)) {
        for (const auto& entry : fs::directory_iterator(macrosDir)) {
            if (entry.path().extension() == L".ps1") {
                g_MacroFiles.push_back(entry.path().filename().wstring());
                if (g_MacroFiles.size() >= MAX_MACROS) break;
            }
        }
    }

    for (size_t i = 0; i < g_MacroFiles.size(); ++i) {
        std::wstring name = g_MacroFiles[i];
        size_t lastDot = name.find_last_of(L".");
        if (lastDot != std::wstring::npos) {
            name = name.substr(0, lastDot);
        }
        
        // Add spaces before capital letters for nicer display?
        // Let's keep it simple for now.
        
        g_MenuNames.push_back(name);
        
        std::wstring shortcut;
        if (i < 9) {
            shortcut = L"Ctrl+Shift+" + std::to_wstring(i + 1);
        }
        g_Shortcuts.push_back(shortcut);

        PluginMenuItem item;
        item.name = g_MenuNames.back().c_str();
        item.callback = g_Callbacks[i];
        item.shortcut = g_Shortcuts.back().empty() ? NULL : g_Shortcuts.back().c_str();
        g_MenuItems.push_back(item);
    }

    *count = (int)g_MenuItems.size();
    return g_MenuItems.data();
}

} // extern "C"

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    return TRUE;
}
