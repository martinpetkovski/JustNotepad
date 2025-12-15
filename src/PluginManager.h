#pragma once
#include <windows.h>
#include <vector>
#include <string>
#include "PluginInterface.h"

// Forward declarations
struct PluginItem;
struct PluginInfo;

struct PluginCommand {
    std::wstring pluginName;
    std::wstring commandName;
    int commandId;
    std::wstring shortcut;
};

struct PluginItem {
    std::wstring name;
    PluginCallback callback;
    std::wstring shortcut;
    UINT vk;
    UINT mods;
    int commandId;
};

struct PluginInfo {
    HMODULE hModule;
    std::wstring path;
    std::wstring filename;
    bool enabled;
    std::wstring name;
    std::wstring description;
    std::wstring version;
    std::wstring license;

    // Function pointers types
    typedef const wchar_t* (*GetPluginNameFunc)();
    typedef const wchar_t* (*GetPluginDescriptionFunc)();
    typedef const wchar_t* (*GetPluginVersionFunc)();
    typedef const wchar_t* (*GetPluginLicenseFunc)();
    typedef const wchar_t* (*GetPluginStatusFunc)(const wchar_t*);
    typedef PluginMenuItem* (*GetPluginMenuItemsFunc)(int*);
    typedef void (*OnFileEventFunc)(const wchar_t*, HWND, const wchar_t*);
    typedef bool (*OnSaveFileFunc)(const wchar_t*, HWND);
    typedef void (*OnTextModifiedFunc)(HWND);
    typedef void (*InitializeFunc)(const wchar_t*);
    typedef void (*ShutdownFunc)();
    typedef void (*SetHostFunctionsFunc)(HostFunctions*);
    typedef long long (*GetMaxFileSizeFunc)();

    GetPluginStatusFunc GetPluginStatus;
    OnFileEventFunc OnFileEvent;
    OnSaveFileFunc OnSaveFile;
    OnTextModifiedFunc OnTextModified;
    InitializeFunc Initialize;
    ShutdownFunc Shutdown;
    SetHostFunctionsFunc SetHostFunctions;
    GetMaxFileSizeFunc GetMaxFileSize;

    std::vector<PluginItem> items;
};

class PluginManager {
public:
    PluginManager();
    ~PluginManager();

    void LoadPlugins(const std::wstring& pluginsDir);
    void UnloadPlugins();
    void LoadSettings(const std::wstring& iniPath);
    void SaveSettings(const std::wstring& iniPath);
    void SetPluginEnabled(const std::wstring& filename, bool enabled);
    
    const std::vector<PluginInfo>& GetPlugins() const;
    std::vector<PluginCommand> GetAllCommands() const;
    void ExecutePluginCommand(int commandId, HWND hEditor);
    
    void NotifyFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType);
    void NotifyTextModified(HWND hEditor);
    bool NotifySaveFile(const wchar_t* filePath, HWND hEditor);
    
    bool TranslateAccelerator(MSG* pMsg);
    void SetHostFunctions(HostFunctions functions);

private:
    std::vector<PluginInfo> m_plugins;
    std::vector<std::wstring> m_enabledPlugins;
    std::vector<PluginCallback> m_callbacks;
    HostFunctions m_hostFunctions;
};


