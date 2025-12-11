#pragma once
#include <windows.h>

#ifdef JUSTNOTEPAD_PLUGIN_EXPORTS
#define PLUGIN_API __declspec(dllexport)
#else
#define PLUGIN_API __declspec(dllimport)
#endif

typedef void (*PluginCallback)(HWND);

struct PluginMenuItem {
    const wchar_t* name;
    PluginCallback callback;
};

extern "C" {
    // Function to get the plugin name
    PLUGIN_API const wchar_t* GetPluginName();
    
    // Function to get the plugin description
    PLUGIN_API const wchar_t* GetPluginDescription();
    
    // Function to get the plugin version
    PLUGIN_API const wchar_t* GetPluginVersion();

    // Function to get the status string for a specific file (optional)
    // Returns NULL or empty string if the plugin has no status for this file
    PLUGIN_API const wchar_t* GetPluginStatus(const wchar_t* filePath);

    // Function to get the menu items provided by the plugin
    // Returns a pointer to an array of PluginMenuItem, and sets count to the number of items
    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count);

    // Optional: Called when a file is loaded or saved
    PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType);
}

