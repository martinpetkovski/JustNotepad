#include "PluginManager.h"
#include "resource.h"
#include <filesystem>
#include <algorithm>

namespace fs = std::filesystem;

PluginManager::PluginManager() {}

PluginManager::~PluginManager() {
    UnloadPlugins();
}

void PluginManager::LoadPlugins(const std::wstring& pluginsDir) {
    UnloadPlugins();

    if (!fs::exists(pluginsDir)) {
        MessageBoxW(NULL, (L"Plugins directory not found: " + pluginsDir).c_str(), L"Debug", MB_OK);
        return;
    }

    bool foundAny = false;
    for (const auto& entry : fs::directory_iterator(pluginsDir)) {
        if (entry.path().extension() == L".dll") {
            foundAny = true;
            HMODULE hModule = LoadLibraryW(entry.path().c_str());
            if (hModule) {
                PluginInfo info;
                info.hModule = hModule;
                info.path = entry.path().wstring();

                auto getName = (PluginInfo::GetPluginNameFunc)GetProcAddress(hModule, "GetPluginName");
                auto getDesc = (PluginInfo::GetPluginDescriptionFunc)GetProcAddress(hModule, "GetPluginDescription");
                auto getVer = (PluginInfo::GetPluginVersionFunc)GetProcAddress(hModule, "GetPluginVersion");
                auto getItems = (PluginInfo::GetPluginMenuItemsFunc)GetProcAddress(hModule, "GetPluginMenuItems");

                if (getName && getItems) {
                    info.name = getName();
                    if (getDesc) info.description = getDesc();
                    if (getVer) info.version = getVer();
                    
                    int count = 0;
                    PluginMenuItem* items = getItems(&count);
                    
                    // Debug
                    // wchar_t buf[256];
                    // wsprintf(buf, L"Plugin loaded: %s\nItems count: %d", info.name.c_str(), count);
                    // MessageBoxW(NULL, buf, L"Debug", MB_OK);

                    if (items && count > 0) {
                        for (int i = 0; i < count; ++i) {
                            PluginItem item;
                            item.name = items[i].name;
                            item.callback = items[i].callback;
                            
                            // Assign ID
                            if (ID_PLUGIN_FIRST + m_callbacks.size() <= ID_PLUGIN_LAST) {
                                item.commandId = ID_PLUGIN_FIRST + (int)m_callbacks.size();
                                m_callbacks.push_back(item.callback);
                                info.items.push_back(item);
                            }
                        }
                    } else {
                         MessageBoxW(NULL, L"Plugin loaded but no items returned.", L"Debug", MB_OK);
                    }

                    m_plugins.push_back(info);
                } else {
                    std::wstring error = L"Failed to load exports from: " + entry.path().filename().wstring();
                    if (!getName) error += L"\nMissing GetPluginName";
                    if (!getItems) error += L"\nMissing GetPluginMenuItems";
                    MessageBoxW(NULL, error.c_str(), L"Debug", MB_OK);
                    FreeLibrary(hModule);
                }
            } else {
                MessageBoxW(NULL, (L"LoadLibrary failed for: " + entry.path().wstring()).c_str(), L"Debug", MB_OK);
            }
        }
    }

    if (!foundAny) {
         MessageBoxW(NULL, L"No DLLs found in plugins directory", L"Debug", MB_OK);
    }

    // Sort plugins by name
    std::sort(m_plugins.begin(), m_plugins.end(), [](const PluginInfo& a, const PluginInfo& b) {
        return a.name < b.name;
    });
}


void PluginManager::UnloadPlugins() {
    for (auto& plugin : m_plugins) {
        if (plugin.hModule) {
            FreeLibrary(plugin.hModule);
        }
    }
    m_plugins.clear();
    m_callbacks.clear();
}

const std::vector<PluginInfo>& PluginManager::GetPlugins() const {
    return m_plugins;
}

std::vector<PluginCommand> PluginManager::GetAllCommands() const {
    std::vector<PluginCommand> commands;
    for (const auto& plugin : m_plugins) {
        for (const auto& item : plugin.items) {
            PluginCommand cmd;
            cmd.pluginName = plugin.name;
            cmd.commandName = item.name;
            cmd.commandId = item.commandId;
            commands.push_back(cmd);
        }
    }
    return commands;
}

void PluginManager::ExecutePluginCommand(int commandId, HWND hEditor) {
    int index = commandId - ID_PLUGIN_FIRST;
    if (index >= 0 && index < m_callbacks.size()) {
        if (m_callbacks[index]) {
            m_callbacks[index](hEditor);
        }
    }
}

