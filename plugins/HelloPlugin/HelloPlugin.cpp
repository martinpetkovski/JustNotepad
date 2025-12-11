#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <richedit.h>

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Hello Plugin";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"A sample plugin that inserts text.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.0";
    }

    void InsertHello(HWND hEditorWnd) {
        SendMessage(hEditorWnd, EM_REPLACESEL, TRUE, (LPARAM)L"Hello from Sample Plugin!");
    }

    PluginMenuItem g_Items[] = {
        { L"Insert Hello", InsertHello }
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
