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
        { L"Insert GUID", InsertGuidDefault },
        { L"Insert GUID (No Braces)", InsertGuidNoBraces },
        { L"Insert GUID (Clean)", InsertGuidClean }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_Items) / sizeof(PluginMenuItem);
        return g_Items;
    }
}
