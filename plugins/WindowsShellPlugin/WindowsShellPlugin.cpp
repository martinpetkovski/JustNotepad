#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <strsafe.h>
#include <shlobj.h>
#include <shlwapi.h>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "shell32.lib")

HINSTANCE g_hInst = NULL;

BOOL IsUserAdmin()
{
    BOOL b;
    SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
    PSID AdministratorsGroup; 
    b = AllocateAndInitializeSid(&NtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &AdministratorsGroup); 
    if(b) 
    {
        if (!CheckTokenMembership( NULL, AdministratorsGroup, &b)) 
        {
             b = FALSE;
        } 
        FreeSid(AdministratorsGroup); 
    }

    return b;
}

BOOL IsAppInShellMenu()
{
    HKEY hKey;
    if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\*\\shell\\JustNotepad"), 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        RegCloseKey(hKey);
        return TRUE;
    }
    return FALSE;
}

void AddAppToShellMenu()
{
    TCHAR szPath[MAX_PATH];
    GetModuleFileName(NULL, szPath, MAX_PATH);
    
    HKEY hKey;
    if (RegCreateKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\*\\shell\\JustNotepad"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("Edit in Just Notepad"), (DWORD)(_tcslen(_T("Edit in Just Notepad")) + 1) * sizeof(TCHAR));
        RegSetValueEx(hKey, _T("Icon"), 0, REG_SZ, (BYTE*)szPath, (DWORD)(_tcslen(szPath) + 1) * sizeof(TCHAR));
        
        HKEY hKeyCmd;
        if (RegCreateKeyEx(hKey, _T("command"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyCmd, NULL) == ERROR_SUCCESS)
        {
            TCHAR szCmd[MAX_PATH + 10];
            StringCchPrintf(szCmd, MAX_PATH + 10, _T("\"%s\" \"%%1\""), szPath);
            RegSetValueEx(hKeyCmd, NULL, 0, REG_SZ, (BYTE*)szCmd, (DWORD)(_tcslen(szCmd) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyCmd);
        }
        RegCloseKey(hKey);
    }
    
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
}

void RemoveAppFromShellMenu()
{
    HMODULE hAdvApi32 = LoadLibrary(_T("Advapi32.dll"));
    if (hAdvApi32)
    {
        typedef LSTATUS (APIENTRY *PFN_RegDeleteTree)(HKEY, LPCTSTR);
        PFN_RegDeleteTree pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeW");
        #ifdef UNICODE
        if (!pfnRegDeleteTree) pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeW");
        #else
        if (!pfnRegDeleteTree) pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeA");
        #endif
        
        if (pfnRegDeleteTree)
        {
            pfnRegDeleteTree(HKEY_CURRENT_USER, _T("Software\\Classes\\*\\shell\\JustNotepad"));
        }
        FreeLibrary(hAdvApi32);
    }
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
}

BOOL IsTxtAssociated()
{
    HKEY hKey;
    TCHAR szValue[256];
    DWORD dwSize = sizeof(szValue);
    if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\.txt"), 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        if (RegQueryValueEx(hKey, NULL, NULL, NULL, (LPBYTE)szValue, &dwSize) == ERROR_SUCCESS)
        {
            RegCloseKey(hKey);
            return _tcscmp(szValue, _T("JustNotepad.txt")) == 0;
        }
        RegCloseKey(hKey);
    }
    return FALSE;
}

void AssociateTxtFiles(HWND hMain)
{
    if (!IsUserAdmin())
    {
        int result = MessageBox(hMain, _T("This feature requires Administrator privileges to set system-wide associations.\nDo you want to restart Just Notepad as Administrator?"), _T("Administrator Required"), MB_YESNO | MB_ICONINFORMATION);
        if (result == IDYES)
        {
            TCHAR szPath[MAX_PATH];
            GetModuleFileName(NULL, szPath, MAX_PATH);
            ShellExecute(NULL, _T("runas"), szPath, NULL, NULL, SW_SHOWNORMAL);
            SendMessage(hMain, WM_CLOSE, 0, 0);
        }
        return;
    }

    TCHAR szPath[MAX_PATH];
    GetModuleFileName(NULL, szPath, MAX_PATH);

    HKEY hKey;
    // .txt -> JustNotepad.txt (HKCR)
    if (RegCreateKeyEx(HKEY_CLASSES_ROOT, _T(".txt"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("JustNotepad.txt"), (DWORD)(_tcslen(_T("JustNotepad.txt")) + 1) * sizeof(TCHAR));
        RegCloseKey(hKey);
    }

    // JustNotepad.txt -> Text Document (HKCR)
    if (RegCreateKeyEx(HKEY_CLASSES_ROOT, _T("JustNotepad.txt"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("Text Document"), (DWORD)(_tcslen(_T("Text Document")) + 1) * sizeof(TCHAR));
        
        // DefaultIcon
        HKEY hKeyIcon;
        if (RegCreateKeyEx(hKey, _T("DefaultIcon"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyIcon, NULL) == ERROR_SUCCESS)
        {
            TCHAR szIcon[MAX_PATH + 5];
            StringCchPrintf(szIcon, MAX_PATH + 5, _T("%s,0"), szPath);
            RegSetValueEx(hKeyIcon, NULL, 0, REG_SZ, (BYTE*)szIcon, (DWORD)(_tcslen(szIcon) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyIcon);
        }
        
        // shell\open\command
        HKEY hKeyCmd;
        if (RegCreateKeyEx(hKey, _T("shell\\open\\command"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyCmd, NULL) == ERROR_SUCCESS)
        {
            TCHAR szCmd[MAX_PATH + 10];
            StringCchPrintf(szCmd, MAX_PATH + 10, _T("\"%s\" \"%%1\""), szPath);
            RegSetValueEx(hKeyCmd, NULL, 0, REG_SZ, (BYTE*)szCmd, (DWORD)(_tcslen(szCmd) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyCmd);
        }
        
        RegCloseKey(hKey);
    }
    
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
    MessageBox(hMain, _T("Association successful!"), _T("Success"), MB_OK | MB_ICONINFORMATION);
}

void ToggleShellMenuCallback(HWND hEditor) {
    HWND hMain = GetParent(hEditor);
    if (IsAppInShellMenu()) {
        RemoveAppFromShellMenu();
        MessageBox(hMain, _T("Removed from Explorer Context Menu."), _T("Windows Shell"), MB_OK | MB_ICONINFORMATION);
    } else {
        AddAppToShellMenu();
        MessageBox(hMain, _T("Added to Explorer Context Menu."), _T("Windows Shell"), MB_OK | MB_ICONINFORMATION);
    }
}

void AssociateTxtCallback(HWND hEditor) {
    HWND hMain = GetParent(hEditor);
    AssociateTxtFiles(hMain);
}

PluginMenuItem g_MenuItems[] = {
    { _T("Toggle Explorer Context Menu"), ToggleShellMenuCallback },
    { _T("Associate with .txt files"), AssociateTxtCallback }
};

extern "C" {

PLUGIN_API const TCHAR* GetPluginName() {
    return _T("WindowsShell");
}

PLUGIN_API const TCHAR* GetPluginDescription() {
    return _T("Provides Windows Shell integration features.");
}

PLUGIN_API const TCHAR* GetPluginVersion() {
    return _T("1.0");
}

PLUGIN_API const wchar_t* GetPluginLicense() {
    return L"MIT License\n\n"
           L"Copyright (c) 2025 Just Notepad Contributors\n\n"
           L"Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\n"
           L"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\n"
           L"THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO, THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
}

PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
    *count = 2;
    return g_MenuItems;
}

} // extern "C"

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}
