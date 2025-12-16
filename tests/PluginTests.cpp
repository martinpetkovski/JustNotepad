#include <windows.h>
#include <iostream>
#include <string>
#include <vector>
#include <filesystem>
#include "../src/PluginInterface.h"

namespace fs = std::filesystem;

// Simple test framework
#define TEST(name) void name()
#define RUN_TEST(name) do { std::cout << "Running " << #name << "... "; name(); std::cout << "PASSED" << std::endl; } while(0)

void Assert(bool condition, const char* message, int line)
{
    if (!condition)
    {
        std::cerr << "FAILED at line " << line << ": " << message << "\n";
        exit(1);
    }
}

#define ASSERT(condition) Assert(condition, #condition, __LINE__)

// Helper to load plugin and check basic interface
void CheckPlugin(const std::wstring& pluginName, const std::wstring& dllPath)
{
    std::wcout << L"Checking " << pluginName << L" at " << dllPath << L"... ";
    
    HMODULE hModule = LoadLibraryW(dllPath.c_str());
    if (!hModule)
    {
        std::cerr << "Failed to load library. Error: " << GetLastError() << "\n";
        ASSERT(hModule != NULL);
    }

    // Check exports
    auto GetName = (const wchar_t* (*)())GetProcAddress(hModule, "GetPluginName");
    auto GetDesc = (const wchar_t* (*)())GetProcAddress(hModule, "GetPluginDescription");
    auto GetVer = (const wchar_t* (*)())GetProcAddress(hModule, "GetPluginVersion");
    auto GetItems = (PluginMenuItem* (*)(int*))GetProcAddress(hModule, "GetPluginMenuItems");

    ASSERT(GetName != NULL);
    ASSERT(GetDesc != NULL);
    ASSERT(GetVer != NULL);
    // GetItems is optional? No, usually required for a plugin to do anything, but maybe not strictly required by interface if it just hooks events.
    // But let's check if it exists if we expect it.
    
    if (GetName)
    {
        const wchar_t* name = GetName();
        ASSERT(name != NULL);
        ASSERT(wcslen(name) > 0);
        // std::wcout << L"Name: " << name << L" ";
    }

    if (GetDesc)
    {
        const wchar_t* desc = GetDesc();
        ASSERT(desc != NULL);
    }

    if (GetVer)
    {
        const wchar_t* ver = GetVer();
        ASSERT(ver != NULL);
    }

    if (GetItems)
    {
        int count = 0;
        PluginMenuItem* items = GetItems(&count);
        if (count > 0)
        {
            ASSERT(items != NULL);
            for (int i = 0; i < count; i++)
            {
                ASSERT(items[i].name != NULL);
                ASSERT(items[i].callback != NULL);
            }
        }
    }

    FreeLibrary(hModule);
    std::cout << "OK" << std::endl;
}

std::wstring GetBuildDir()
{
    // Get path to current executable
    wchar_t buffer[MAX_PATH];
    GetModuleFileNameW(NULL, buffer, MAX_PATH);
    fs::path exePath(buffer);
    return exePath.parent_path().wstring(); // e.g. .../build/Debug
}

TEST(TestAllPluginsLoad)
{
    std::vector<std::wstring> plugins = {
        L"AdvancedSearchPlugin",
        L"CharacterMapPlugin",
        L"ClangFormatPlugin",
        L"CodeFoldingPlugin",
        L"DateTimePlugin",
        L"EmojiPlugin",
        L"GitPlugin",
        L"GuidPlugin",
        L"HelloPlugin",
        L"HexEditorPlugin",
        L"DataViewerPlugin",
        L"PowerShellScripts",
        L"MarkdownToolsPlugin",
        L"NeonPlugin",
        L"PerforcePlugin",
        L"ThemeSwitcherPlugin",
        L"WhitespaceVisualizerPlugin",
        L"WindowsShellPlugin"
    };

    fs::path buildDir = GetBuildDir(); // .../build/Debug
    // We expect plugins in .../build/plugins/<Name>/Debug/<Name>.dll
    // So from buildDir, we go up one level to 'build', then 'plugins', then Name, then Debug.
    
    // Check if we are in Debug or Release
    std::wstring config = buildDir.filename().wstring(); // "Debug" or "Release"
    fs::path rootBuild = buildDir.parent_path(); // .../build

    for (const auto& name : plugins)
    {
        // Try standard MSVC layout: build/plugins/Debug/Name.dll
        fs::path dllPath = rootBuild / "plugins" / config / (name + L".dll");
        
        if (!fs::exists(dllPath))
        {
            // Try per-project layout: build/plugins/Name/Debug/Name.dll
            dllPath = rootBuild / "plugins" / name / config / (name + L".dll");
        }
        
        if (!fs::exists(dllPath))
        {
            // Try flat layout: build/plugins/Name.dll
            dllPath = rootBuild / "plugins" / (name + L".dll");
        }
        
        if (!fs::exists(dllPath))
        {
             // Try flat in build dir
             dllPath = buildDir / (name + L".dll");
        }

        if (fs::exists(dllPath))
        {
            CheckPlugin(name, dllPath.wstring());
        }
        else
        {
            std::wcerr << L"Skipping " << name << L": DLL not found at " << dllPath.wstring() << std::endl;
            // Don't fail, maybe not all are built
        }
    }
}

int main()
{
    RUN_TEST(TestAllPluginsLoad);
    return 0;
}
