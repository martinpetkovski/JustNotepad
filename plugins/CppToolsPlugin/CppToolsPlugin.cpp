#include "../../src/PluginInterface.h"
#include <windows.h>
#include <string>
#include <vector>
#include <filesystem>
#include <algorithm>

namespace fs = std::filesystem;

HostFunctions* g_HostFunctions = nullptr;

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"C/C++ Tools";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Tools for C/C++ development.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.0";
    }

    PLUGIN_API const wchar_t* GetPluginLicense() {
        return L"MIT";
    }

    PLUGIN_API void SetHostFunctions(HostFunctions* functions) {
        g_HostFunctions = functions;
    }

    void SwitchSourceHeader(HWND hEditor) {
        if (!g_HostFunctions) return;

        // Save current file
        if (g_HostFunctions->SaveFile) {
            g_HostFunctions->SaveFile();
        }

        // Get current file path
        wchar_t currentPath[MAX_PATH] = {0};
        if (g_HostFunctions->GetCurrentFilePath) {
            g_HostFunctions->GetCurrentFilePath(currentPath, MAX_PATH);
        }

        if (wcslen(currentPath) == 0) return;

        fs::path path(currentPath);
        std::wstring ext = path.extension().wstring();
        std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);

        std::vector<std::wstring> sourceExts = {L".cpp", L".c", L".cxx", L".cc"};
        std::vector<std::wstring> headerExts = {L".h", L".hpp", L".hxx"};

        bool isSource = false;

        for (const auto& e : sourceExts) {
            if (ext == e) {
                isSource = true;
                break;
            }
        }

        std::vector<std::wstring> targetExts;
        if (isSource) {
            targetExts = headerExts;
        } else {
            // Check if it is header
            bool isHeader = false;
            for (const auto& e : headerExts) {
                if (ext == e) {
                    isHeader = true;
                    break;
                }
            }
            if (isHeader) {
                targetExts = sourceExts;
            } else {
                return; // Not a C/C++ file
            }
        }

        // Try to find the target file
        fs::path basePath = path;
        basePath.replace_extension(L""); // Remove extension (this might be wrong if filename has multiple dots, but standard replace_extension handles it)
        
        // Actually replace_extension replaces the last extension.
        // If we have file.tar.gz, replace_extension("") -> file.tar
        // But for source files it's usually file.cpp -> file
        
        // However, we should use the stem and parent path to be safe.
        fs::path parent = path.parent_path();
        std::wstring stem = path.stem().wstring();

        for (const auto& targetExt : targetExts) {
            fs::path targetPath = parent / (stem + targetExt);
            
            if (fs::exists(targetPath)) {
                if (g_HostFunctions->OpenFile) {
                    g_HostFunctions->OpenFile(targetPath.c_str());
                }
                return;
            }
        }
        
        MessageBox(hEditor, L"Corresponding source/header file not found.", L"C/C++ Tools", MB_ICONINFORMATION);
    }

    PluginMenuItem g_Items[] = {
        { L"Switch Source/Header", SwitchSourceHeader, L"Alt+O" }
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
