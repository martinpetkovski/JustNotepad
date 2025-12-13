#pragma once
#include <windows.h>
#pragma once
#include <windows.h>
#include <vector>
#include <string>
#include "PluginInterface.h"

struct PluginItem {
    std::wstring name;
    PluginCallback callback;
    int commandId;
    std::wstring shortcut;
    UINT vk;
    UINT mods;
};

struct PluginInfo {
    HMODULE hModule;
    std::wstring name;
    std::wstring description;
    std::wstring version;
    std::wstring path;
    std::wstring filename;
    bool enabled;
    std::vector<PluginItem> items;
    
    // Function pointers
    typedef const wchar_t* (*GetPluginNameFunc)();
    typedef const wchar_t* (*GetPluginDescriptionFunc)();
    typedef const wchar_t* (*GetPluginVersionFunc)();
    typedef const wchar_t* (*GetPluginStatusFunc)(const wchar_t*);
    typedef PluginMenuItem* (*GetPluginMenuItemsFunc)(int*);
    typedef void (*OnFileEventFunc)(const wchar_t*, HWND, const wchar_t*);
    typedef bool (*OnSaveFileFunc)(const wchar_t*, HWND);
    typedef void (*OnTextModifiedFunc)(HWND);
    typedef void (*InitializeFunc)(const wchar_t*);
    typedef void (*ShutdownFunc)();
    
    GetPluginStatusFunc GetPluginStatus;
    OnFileEventFunc OnFileEvent;
    OnSaveFileFunc OnSaveFile;
    OnTextModifiedFunc OnTextModified;
    InitializeFunc Initialize;
    ShutdownFunc Shutdown;
};

struct PluginCommand {
    std::wstring pluginName;
    std::wstring commandName;
    int commandId;
    std::wstring shortcut;
};

class PluginManager {
public:
    PluginManager();
    ~PluginManager();

    void LoadPlugins(const std::wstring& pluginsDir);
    void UnloadPlugins();
    const std::vector<PluginInfo>& GetPlugins() const;
    std::vector<PluginCommand> GetAllCommands() const;
    void ExecutePluginCommand(int commandId, HWND hEditor);

    void LoadSettings(const std::wstring& iniPath);
    void SaveSettings(const std::wstring& iniPath);
    void SetPluginEnabled(const std::wstring& filename, bool enabled);
    
    void NotifyFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType);
    bool NotifySaveFile(const wchar_t* filePath, HWND hEditor);
    void NotifyTextModified(HWND hEditor);
    
    bool TranslateAccelerator(MSG* pMsg);

private:
    std::vector<PluginInfo> m_plugins;
    std::vector<PluginCallback> m_callbacks; // Map from (CommandID - Base) to Callback
    std::vector<std::wstring> m_enabledPlugins;
};

