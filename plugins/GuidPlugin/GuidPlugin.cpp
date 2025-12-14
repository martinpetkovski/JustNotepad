#include "../../src/PluginInterface.h"
#include <windows.h>
#include <objbase.h>
#include <tchar.h>
#include <richedit.h>
#include <string>
#include <vector>

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"GUID Generator";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Generates and inserts GUIDs.";
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

    void InsertGuid(HWND hEditorWnd, bool braces, bool hyphens) {
        GUID guid;
        if (CoCreateGuid(&guid) == S_OK) {
            wchar_t buffer[64];
            // StringFromGUID2 produces {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
            wchar_t temp[64];
            if (StringFromGUID2(guid, temp, 64) != 0) {
                std::wstring s(temp);
                
                // Remove braces if not requested
                if (!braces) {
                    if (s.length() > 2) {
                        s = s.substr(1, s.length() - 2);
                    }
                }
                
                // Remove hyphens if not requested
                if (!hyphens) {
                    std::wstring clean;
                    for (wchar_t c : s) {
                        if (c != L'-') clean += c;
                    }
                    s = clean;
                }
                
                wcscpy_s(buffer, s.c_str());
                SendMessage(hEditorWnd, EM_REPLACESEL, TRUE, (LPARAM)buffer);
            }
        }
    }

    void InsertGuidDefault(HWND h) { InsertGuid(h, true, true); }
    void InsertGuidNoBraces(HWND h) { InsertGuid(h, false, true); }
    void InsertGuidClean(HWND h) { InsertGuid(h, false, false); }

    PluginMenuItem g_Items[] = {
        { L"Insert GUID", InsertGuidDefault, L"Ctrl+Shift+Alt+G" },
        { L"Insert GUID (No Braces)", InsertGuidNoBraces },
        { L"Insert GUID (Clean)", InsertGuidClean }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_Items) / sizeof(PluginMenuItem);
        return g_Items;
    }
}
