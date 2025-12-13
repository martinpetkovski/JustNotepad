#pragma once
#include <windows.h>

#ifdef JUSTNOTEPAD_PLUGIN_EXPORTS
#define PLUGIN_API __declspec(dllexport)
#else
#define PLUGIN_API __declspec(dllimport)
#endif

typedef void (*PluginCallback)(HWND);

struct HostFunctions {
    void (*SetProgress)(int percent);
};

struct PluginMenuItem {
    const wchar_t* name;
    PluginCallback callback;
    const wchar_t* shortcut; // e.g. "Ctrl+Shift+M"
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

    // Optional: Called before saving a file. Return true if the plugin handled the save.
    PLUGIN_API bool OnSaveFile(const wchar_t* filePath, HWND hEditor);

    // Optional: Called when text is modified in the editor
    PLUGIN_API void OnTextModified(HWND hEditor);

    // Optional: Called when the plugin is loaded
    PLUGIN_API void Initialize(const wchar_t* settingsPath);

    // Optional: Called when the plugin is unloaded
    PLUGIN_API void Shutdown();

    // Optional: Called by the host to provide callback functions
    PLUGIN_API void SetHostFunctions(HostFunctions* functions);

    // Optional: Get maximum supported file size in bytes. Return 0 for no limit.
    PLUGIN_API long long GetMaxFileSize();
}

