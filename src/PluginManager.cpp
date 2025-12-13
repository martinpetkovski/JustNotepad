#include "PluginManager.h"
#include "resource.h"
#include <filesystem>
#include <algorithm>
#include <richedit.h>

namespace fs = std::filesystem;

static bool ParseShortcut(const std::wstring& shortcut, UINT& key, UINT& modifiers);

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
                
                // Check if enabled
                info.enabled = false;
                for (const auto& enabled : m_enabledPlugins) {
                    if (_wcsicmp(enabled.c_str(), info.filename.c_str()) == 0) {
                        info.enabled = true;
                        break;
                    }
                }

                auto getName = (PluginInfo::GetPluginNameFunc)GetProcAddress(hModule, "GetPluginName");
                auto getDesc = (PluginInfo::GetPluginDescriptionFunc)GetProcAddress(hModule, "GetPluginDescription");
                auto getVer = (PluginInfo::GetPluginVersionFunc)GetProcAddress(hModule, "GetPluginVersion");
                auto getStatus = (PluginInfo::GetPluginStatusFunc)GetProcAddress(hModule, "GetPluginStatus");
                auto getItems = (PluginInfo::GetPluginMenuItemsFunc)GetProcAddress(hModule, "GetPluginMenuItems");
                auto onFileEvent = (PluginInfo::OnFileEventFunc)GetProcAddress(hModule, "OnFileEvent");
                auto onSaveFile = (PluginInfo::OnSaveFileFunc)GetProcAddress(hModule, "OnSaveFile");
                auto onTextModified = (PluginInfo::OnTextModifiedFunc)GetProcAddress(hModule, "OnTextModified");
                auto initialize = (PluginInfo::InitializeFunc)GetProcAddress(hModule, "Initialize");
                auto shutdown = (PluginInfo::ShutdownFunc)GetProcAddress(hModule, "Shutdown");
                auto setHostFunctions = (PluginInfo::SetHostFunctionsFunc)GetProcAddress(hModule, "SetHostFunctions");
                auto getMaxFileSize = (PluginInfo::GetMaxFileSizeFunc)GetProcAddress(hModule, "GetMaxFileSize");

                if (getName && getItems) {
                    info.name = getName();
                    if (getDesc) info.description = getDesc();
                    if (getVer) info.version = getVer();
                    info.GetPluginStatus = getStatus;
                    info.OnFileEvent = onFileEvent;
                    info.OnSaveFile = onSaveFile;
                    info.OnTextModified = onTextModified;
                    info.Initialize = initialize;
                    info.Shutdown = shutdown;
                    info.SetHostFunctions = setHostFunctions;
                    info.GetMaxFileSize = getMaxFileSize;

                    if (info.enabled) {
                        if (info.SetHostFunctions) {
                            info.SetHostFunctions(&m_hostFunctions);
                        }
                        if (info.Initialize) {
                            std::wstring settingsPath = entry.path().parent_path().wstring() + L"\\settings\\" + info.filename + L".ini";
                            fs::create_directories(fs::path(settingsPath).parent_path());
                            info.Initialize(settingsPath.c_str());
                        }
                    }
                    
                    int count = 0;
                    PluginMenuItem* items = getItems(&count);
                    
                    if (items && count > 0) {
                        for (int i = 0; i < count; ++i) {
                            PluginItem item;
                            item.name = items[i].name;
                            item.callback = items[i].callback;
                            if (items[i].shortcut) {
                                item.shortcut = items[i].shortcut;
                                ParseShortcut(item.shortcut, item.vk, item.mods);
                            } else {
                                item.vk = 0;
                                item.mods = 0;
                            }
                            
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
        if (plugin.enabled && plugin.Shutdown) {
            plugin.Shutdown();
        }
        if (plugin.hModule) {
            FreeLibrary(plugin.hModule);
        }
    }
    m_plugins.clear();
    m_callbacks.clear();
}

void PluginManager::LoadSettings(const std::wstring& iniPath) {
    m_enabledPlugins.clear();
    
    // Read enabled plugins from [EnabledPlugins] section
    // Format: PluginFilename=1
    
    const int bufSize = 4096;
    std::vector<TCHAR> buffer(bufSize);
    GetPrivateProfileSection(L"EnabledPlugins", buffer.data(), bufSize, iniPath.c_str());
    
    TCHAR* p = buffer.data();
    while (*p) {
        std::wstring line = p;
        size_t eq = line.find(L'=');
        if (eq != std::wstring::npos) {
            std::wstring key = line.substr(0, eq);
            std::wstring val = line.substr(eq + 1);
            if (val == L"1") {
                m_enabledPlugins.push_back(key);
            }
        }
        p += line.length() + 1;
    }
}

void PluginManager::SaveSettings(const std::wstring& iniPath) {
    // Clear section first
    WritePrivateProfileSection(L"EnabledPlugins", NULL, iniPath.c_str());
    
    for (const auto& name : m_enabledPlugins) {
        WritePrivateProfileString(L"EnabledPlugins", name.c_str(), L"1", iniPath.c_str());
    }
}

void PluginManager::SetPluginEnabled(const std::wstring& filename, bool enabled) {
    if (!enabled) {
        // Remove from enabled list
        auto it = std::remove_if(m_enabledPlugins.begin(), m_enabledPlugins.end(), 
            [&](const std::wstring& s) { return _wcsicmp(s.c_str(), filename.c_str()) == 0; });
        m_enabledPlugins.erase(it, m_enabledPlugins.end());
    } else {
        // Add to enabled list if not present
        bool found = false;
        for (const auto& s : m_enabledPlugins) {
            if (_wcsicmp(s.c_str(), filename.c_str()) == 0) {
                found = true;
                break;
            }
        }
        if (!found) {
            m_enabledPlugins.push_back(filename);
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
            cmd.shortcut = item.shortcut;
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
    long long fileSize = -1;
    if (filePath && *filePath) {
        WIN32_FILE_ATTRIBUTE_DATA fad;
        if (GetFileAttributesExW(filePath, GetFileExInfoStandard, &fad)) {
            LARGE_INTEGER size;
            size.HighPart = fad.nFileSizeHigh;
            size.LowPart = fad.nFileSizeLow;
            fileSize = size.QuadPart;
        }
    }

    for (const auto& plugin : m_plugins) {
        if (plugin.enabled && plugin.OnFileEvent) {
            if (fileSize != -1 && plugin.GetMaxFileSize) {
                long long max = plugin.GetMaxFileSize();
                if (max > 0 && fileSize > max) continue;
            }
            plugin.OnFileEvent(filePath, hEditor, eventType);
        }
    }
}

void PluginManager::NotifyTextModified(HWND hEditor) {
    long long textSize = -1;
    
    for (const auto& plugin : m_plugins) {
        if (plugin.enabled && plugin.OnTextModified) {
            if (plugin.GetMaxFileSize) {
                long long max = plugin.GetMaxFileSize();
                if (max > 0) {
                    if (textSize == -1) {
                        // Calculate only if needed
                        GETTEXTLENGTHEX gtl = { GTL_DEFAULT | GTL_PRECISE, 1200 }; 
                        textSize = SendMessage(hEditor, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 0);
                        // Approximate byte size for UTF-16
                        textSize *= 2;
                    }
                    if (textSize > max) continue;
                }
            }
            plugin.OnTextModified(hEditor);
        }
    }
}

static bool ParseShortcut(const std::wstring& shortcut, UINT& key, UINT& modifiers) {
    if (shortcut.empty()) return false;
    
    modifiers = 0;
    key = 0;
    
    std::wstring s = shortcut;
    std::transform(s.begin(), s.end(), s.begin(), ::towupper);
    
    size_t pos = 0;
    while ((pos = s.find(L"+")) != std::wstring::npos) {
        std::wstring mod = s.substr(0, pos);
        if (mod == L"CTRL") modifiers |= MOD_CONTROL;
        else if (mod == L"SHIFT") modifiers |= MOD_SHIFT;
        else if (mod == L"ALT") modifiers |= MOD_ALT;
        s.erase(0, pos + 1);
    }
    
    // Last part is the key
    if (s.length() == 1) {
        key = s[0];
    } else if (s.substr(0, 1) == L"F") {
        int f = _wtoi(s.substr(1).c_str());
        if (f >= 1 && f <= 12) key = VK_F1 + (f - 1);
    } else if (s == L"DEL" || s == L"DELETE") {
        key = VK_DELETE;
    } else if (s == L"INS" || s == L"INSERT") {
        key = VK_INSERT;
    } else if (s == L"HOME") {
        key = VK_HOME;
    } else if (s == L"END") {
        key = VK_END;
    } else if (s == L"PGUP" || s == L"PAGEUP") {
        key = VK_PRIOR;
    } else if (s == L"PGDN" || s == L"PAGEDOWN") {
        key = VK_NEXT;
    }
    
    return key != 0;
}

bool PluginManager::TranslateAccelerator(MSG* pMsg) {
    if (pMsg->message != WM_KEYDOWN && pMsg->message != WM_SYSKEYDOWN) return false;
    
    for (const auto& plugin : m_plugins) {
        if (!plugin.enabled) continue;
        for (const auto& item : plugin.items) {
            if (item.vk == 0) continue;
            
            if (pMsg->wParam == item.vk) {
                // Check modifiers
                bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
                
                bool reqCtrl = (item.mods & MOD_CONTROL) != 0;
                bool reqShift = (item.mods & MOD_SHIFT) != 0;
                bool reqAlt = (item.mods & MOD_ALT) != 0;
                
                if (ctrl == reqCtrl && shift == reqShift && alt == reqAlt) {
                    if (item.callback) {
                        // We need hEditor. Since we don't have it here, we can try to find it.
                        // Or we can pass it to TranslateAccelerator.
                        // For now, let's assume the active window or find it.
                        // Actually, ExecutePluginCommand takes hEditor.
                        // Let's assume the main window is the parent of the focused window or similar.
                        // But wait, ExecutePluginCommand uses m_callbacks which is indexed by ID.
                        // We can just call the callback directly if we have hEditor.
                        
                        // Better: Post a WM_COMMAND to the main window with the command ID.
                        // This handles hEditor correctly in main.cpp
                        HWND hMain = GetAncestor(pMsg->hwnd, GA_ROOT);
                        SendMessage(hMain, WM_COMMAND, item.commandId, 0);
                        return true;
                    }
                }
            }
        }
    }
    return false;
}

bool PluginManager::NotifySaveFile(const wchar_t* filePath, HWND hEditor) {
    long long textSize = -1;

    for (const auto& plugin : m_plugins) {
        if (plugin.enabled && plugin.OnSaveFile) {
            if (plugin.GetMaxFileSize) {
                long long max = plugin.GetMaxFileSize();
                if (max > 0) {
                    if (textSize == -1) {
                        GETTEXTLENGTHEX gtl = { GTL_DEFAULT | GTL_PRECISE, 1200 }; 
                        textSize = SendMessage(hEditor, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 0);
                        textSize *= 2;
                    }
                    if (textSize > max) continue;
                }
            }

            if (plugin.OnSaveFile(filePath, hEditor)) {
                return true;
            }
        }
    }
    return false;
}

void PluginManager::SetHostFunctions(HostFunctions functions) {
    m_hostFunctions = functions;
}

