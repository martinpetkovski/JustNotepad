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
                info.filename = entry.path().filename().wstring();
                
                // Check if disabled
                info.enabled = true;
                for (const auto& disabled : m_disabledPlugins) {
                    if (_wcsicmp(disabled.c_str(), info.filename.c_str()) == 0) {
                        info.enabled = false;
                        break;
                    }
                }

                auto getName = (PluginInfo::GetPluginNameFunc)GetProcAddress(hModule, "GetPluginName");
                auto getDesc = (PluginInfo::GetPluginDescriptionFunc)GetProcAddress(hModule, "GetPluginDescription");
                auto getVer = (PluginInfo::GetPluginVersionFunc)GetProcAddress(hModule, "GetPluginVersion");
                auto getStatus = (PluginInfo::GetPluginStatusFunc)GetProcAddress(hModule, "GetPluginStatus");
                auto getItems = (PluginInfo::GetPluginMenuItemsFunc)GetProcAddress(hModule, "GetPluginMenuItems");
                auto onFileEvent = (PluginInfo::OnFileEventFunc)GetProcAddress(hModule, "OnFileEvent");

                if (getName && getItems) {
                    info.name = getName();
                    if (getDesc) info.description = getDesc();
                    if (getVer) info.version = getVer();
                    info.GetPluginStatus = getStatus;
                    info.OnFileEvent = onFileEvent;
                    
                    int count = 0;
                    PluginMenuItem* items = getItems(&count);
                    
                    if (items && count > 0) {
                        for (int i = 0; i < count; ++i) {
                            PluginItem item;
                            item.name = items[i].name;
                            item.callback = items[i].callback;
                            
                            // Only assign ID and add callback if enabled
                            if (info.enabled) {
                                if (ID_PLUGIN_FIRST + m_callbacks.size() <= ID_PLUGIN_LAST) {
                                    item.commandId = ID_PLUGIN_FIRST + (int)m_callbacks.size();
                                    m_callbacks.push_back(item.callback);
                                    info.items.push_back(item);
                                }
                            } else {
                                // Still add to items so we know what it has, but no ID
                                item.commandId = 0;
                                info.items.push_back(item);
                            }
                        }
                    } else {
                         // MessageBoxW(NULL, L"Plugin loaded but no items returned.", L"Debug", MB_OK);
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

void PluginManager::LoadSettings(const std::wstring& iniPath) {
    m_disabledPlugins.clear();
    
    // Read disabled plugins from [DisabledPlugins] section
    // Format: PluginFilename=1
    
    const int bufSize = 4096;
    std::vector<TCHAR> buffer(bufSize);
    GetPrivateProfileSection(L"DisabledPlugins", buffer.data(), bufSize, iniPath.c_str());
    
    TCHAR* p = buffer.data();
    while (*p) {
        std::wstring line = p;
        size_t eq = line.find(L'=');
        if (eq != std::wstring::npos) {
            std::wstring key = line.substr(0, eq);
            std::wstring val = line.substr(eq + 1);
            if (val == L"1") {
                m_disabledPlugins.push_back(key);
            }
        }
        p += line.length() + 1;
    }
}

void PluginManager::SaveSettings(const std::wstring& iniPath) {
    // Clear section first
    WritePrivateProfileSection(L"DisabledPlugins", NULL, iniPath.c_str());
    
    for (const auto& name : m_disabledPlugins) {
        WritePrivateProfileString(L"DisabledPlugins", name.c_str(), L"1", iniPath.c_str());
    }
}

void PluginManager::SetPluginEnabled(const std::wstring& filename, bool enabled) {
    if (enabled) {
        // Remove from disabled list
        auto it = std::remove_if(m_disabledPlugins.begin(), m_disabledPlugins.end(), 
            [&](const std::wstring& s) { return _wcsicmp(s.c_str(), filename.c_str()) == 0; });
        m_disabledPlugins.erase(it, m_disabledPlugins.end());
    } else {
        // Add to disabled list if not present
        bool found = false;
        for (const auto& s : m_disabledPlugins) {
            if (_wcsicmp(s.c_str(), filename.c_str()) == 0) {
                found = true;
                break;
            }
        }
        if (!found) {
            m_disabledPlugins.push_back(filename);
        }
    }

    // Update runtime state
    for (auto& plugin : m_plugins) {
        if (_wcsicmp(plugin.filename.c_str(), filename.c_str()) == 0) {
            if (plugin.enabled != enabled) {
                plugin.enabled = enabled;
                
                if (enabled) {
                    // If enabling, ensure items have command IDs and callbacks are registered
                    for (auto& item : plugin.items) {
                        if (item.commandId == 0) {
                            if (ID_PLUGIN_FIRST + m_callbacks.size() <= ID_PLUGIN_LAST) {
                                item.commandId = ID_PLUGIN_FIRST + (int)m_callbacks.size();
                                m_callbacks.push_back(item.callback);
                            }
                        }
                    }
                }
            }
            break;
        }
    }
}

const std::vector<PluginInfo>& PluginManager::GetPlugins() const {
    return m_plugins;
}

std::vector<PluginCommand> PluginManager::GetAllCommands() const {
    std::vector<PluginCommand> commands;
    for (const auto& plugin : m_plugins) {
        if (!plugin.enabled) continue; // Skip disabled plugins
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

void PluginManager::NotifyFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
    for (const auto& plugin : m_plugins) {
        if (plugin.enabled && plugin.OnFileEvent) {
            plugin.OnFileEvent(filePath, hEditor, eventType);
        }
    }
}

