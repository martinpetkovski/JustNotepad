#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <richedit.h>
#include <string>
#include <vector>
#include <ctime>

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Date and Time";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Inserts date and time in various formats.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.0";
    }

    void InsertFormat(HWND hEditorWnd, const wchar_t* format) {
        time_t rawtime;
        struct tm* timeinfo;
        wchar_t buffer[256];

        time(&rawtime);
        timeinfo = localtime(&rawtime);

        wcsftime(buffer, 256, format, timeinfo);
        SendMessage(hEditorWnd, EM_REPLACESEL, TRUE, (LPARAM)buffer);
    }

    void InsertShortDate(HWND h) { InsertFormat(h, L"%Y-%m-%d"); }
    void InsertLongDate(HWND h) { InsertFormat(h, L"%A, %B %d, %Y"); }
    void InsertShortTime(HWND h) { InsertFormat(h, L"%H:%M"); }
    void InsertLongTime(HWND h) { InsertFormat(h, L"%H:%M:%S"); }
    void InsertDateTime(HWND h) { InsertFormat(h, L"%Y-%m-%d %H:%M:%S"); }
    void InsertISO8601(HWND h) { InsertFormat(h, L"%Y-%m-%dT%H:%M:%S"); }
    void InsertRFC2822(HWND h) { InsertFormat(h, L"%a, %d %b %Y %H:%M:%S"); }

    PluginMenuItem g_Items[] = {
        { L"Short Date (YYYY-MM-DD)", InsertShortDate },
        { L"Long Date", InsertLongDate },
        { L"Short Time (HH:MM)", InsertShortTime },
        { L"Long Time (HH:MM:SS)", InsertLongTime },
        { L"Date & Time", InsertDateTime, L"F5" },
        { L"ISO 8601", InsertISO8601 },
        { L"RFC 2822", InsertRFC2822 }
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
