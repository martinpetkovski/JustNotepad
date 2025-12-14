#include "../../src/PluginInterface.h"
#include <windows.h>
#include <string>

#define WM_APP_SET_THEME (WM_APP + 200)

int g_CurrentTheme = 0;
std::wstring g_SettingsPath;

extern "C" {
    PLUGIN_API void Initialize(const wchar_t* settingsPath) {
        g_SettingsPath = settingsPath;
        g_CurrentTheme = GetPrivateProfileInt(L"Settings", L"Theme", 0, g_SettingsPath.c_str());
        
        // Apply theme on startup
        HWND hMain = FindWindow(L"JustNotepadClass", NULL);
        if (hMain) {
            SendMessage(hMain, WM_APP_SET_THEME, g_CurrentTheme, 0);
        }
    }

    PLUGIN_API void Shutdown() {
        if (!g_SettingsPath.empty()) {
            WritePrivateProfileString(L"Settings", L"Theme", std::to_wstring(g_CurrentTheme).c_str(), g_SettingsPath.c_str());
        }
    }

    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Theme Switcher";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Switches the application theme.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.0";
    }

    PLUGIN_API const wchar_t* GetPluginLicense() {
        return L"MIT License\n\n"
               L"Copyright (c) 2025 Just Notepad Contributors\n\n"
               L"Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\n"
               L"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\n"
               L"THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO, THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
    }

    void SetThemeLight(HWND hEditorWnd) {
        g_CurrentTheme = 0;
        HWND hMain = GetParent(hEditorWnd);
        SendMessage(hMain, WM_APP_SET_THEME, 0, 0);
    }

    void SetThemeDark(HWND hEditorWnd) {
        g_CurrentTheme = 1;
        HWND hMain = GetParent(hEditorWnd);
        SendMessage(hMain, WM_APP_SET_THEME, 1, 0);
    }

    void SetThemeSolarized(HWND hEditorWnd) {
        g_CurrentTheme = 2;
        HWND hMain = GetParent(hEditorWnd);
        SendMessage(hMain, WM_APP_SET_THEME, 2, 0);
    }

    PluginMenuItem g_Items[] = {
        { L"Light Theme", SetThemeLight },
        { L"Dark Theme", SetThemeDark },
        { L"Solarized Dark Theme", SetThemeSolarized }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_Items) / sizeof(g_Items[0]);
        return g_Items;
    }
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    return TRUE;
}
