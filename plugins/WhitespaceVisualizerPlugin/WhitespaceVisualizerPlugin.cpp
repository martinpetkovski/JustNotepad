#include "../../src/PluginInterface.h"
#include <windows.h>
#include <richedit.h>
#include <tchar.h>
#include <vector>
#include <string>

// Global state
static HWND g_hEditor = NULL;
static WNDPROC g_OldEditorProc = NULL;
static COLORREF g_WhitespaceColor = RGB(150, 150, 150);
static HPEN g_hPen = NULL;
static HBRUSH g_hBrush = NULL;
static bool g_bIsBinary = false;

// Initialize drawing resources
void InitializeDrawingResources() {
    if (!g_hPen) {
        g_hPen = CreatePen(PS_SOLID, 1, g_WhitespaceColor);
        g_hBrush = CreateSolidBrush(g_WhitespaceColor);
    }
}

// Custom window procedure for the editor
LRESULT CALLBACK EditorSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    if (uMsg == WM_PAINT) {
        // Call the default paint first
        LRESULT result = CallWindowProc(g_OldEditorProc, hwnd, uMsg, wParam, lParam);
        
        if (g_bIsBinary) return result;

        // Now draw our whitespace indicators
        HDC hdc = GetDC(hwnd);
        if (hdc) {
            InitializeDrawingResources();
            
            // Get client rect
            RECT rcClient;
            GetClientRect(hwnd, &rcClient);
            
            // Get visible text range
            CHARRANGE crVisible;
            SendMessage(hwnd, EM_GETRECT, 0, (LPARAM)&rcClient);
            
            POINTL ptStart = {rcClient.left, rcClient.top};
            POINTL ptEnd = {rcClient.right, rcClient.bottom};
            int startChar = (int)SendMessage(hwnd, EM_CHARFROMPOS, 0, (LPARAM)&ptStart);
            int endChar = (int)SendMessage(hwnd, EM_CHARFROMPOS, 0, (LPARAM)&ptEnd);
            
            if (startChar >= 0 && endChar >= startChar) {
                // Get the text range
                int rangeLen = endChar - startChar + 1;
                if (rangeLen > 0 && rangeLen < 1000000) { // Safety check
                    static std::vector<wchar_t> buffer;
                    if (buffer.size() < rangeLen + 1) buffer.resize(rangeLen + 1);
                    
                    TEXTRANGE tr;
                    tr.chrg.cpMin = startChar;
                    tr.chrg.cpMax = endChar;
                    tr.lpstrText = buffer.data();
                    
                    int actualLen = (int)SendMessage(hwnd, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                    
                    if (actualLen > 0) {
                        // Set up drawing
                        HPEN hOldPen = (HPEN)SelectObject(hdc, g_hPen);
                        HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, g_hBrush);
                        SetTextColor(hdc, g_WhitespaceColor);
                        SetBkMode(hdc, TRANSPARENT);
                        
                        // Get font size for proper positioning
                        CHARFORMAT2 cf;
                        ZeroMemory(&cf, sizeof(cf));
                        cf.cbSize = sizeof(cf);
                        SendMessage(hwnd, EM_GETCHARFORMAT, SCF_DEFAULT, (LPARAM)&cf);
                        int dotOffset = cf.yHeight / 40;
                        
                        // Draw whitespace for each character
                        for (int i = 0; i < actualLen; i++) {
                            wchar_t ch = buffer[i];
                            if (ch == L' ' || ch == L'\t') {
                                // Get position efficiently
                                POINTL pt;
                                SendMessage(hwnd, EM_POSFROMCHAR, (WPARAM)&pt, startChar + i);
                                
                                if (ch == L' ') {
                                    // Draw a small dot
                                    int centerX = pt.x + 3;
                                    int centerY = pt.y + dotOffset;
                                    Ellipse(hdc, centerX - 1, centerY - 1, centerX + 2, centerY + 2);
                                } else {
                                    // Draw arrow
                                    TextOutW(hdc, pt.x + 2, pt.y, L"→", 1);
                                }
                            }
                        }
                        
                        // Clean up
                        SelectObject(hdc, hOldPen);
                        SelectObject(hdc, hOldBrush);
                    }
                }
            }
            
            ReleaseDC(hwnd, hdc);
        }
        
        return result;
    }
    
    return CallWindowProc(g_OldEditorProc, hwnd, uMsg, wParam, lParam);
}

// Initialize plugin - subclass editor on first menu call
void EnsureSubclassed(HWND hEditorWnd) {
    if (g_OldEditorProc == NULL) {
        g_hEditor = hEditorWnd;
        g_OldEditorProc = (WNDPROC)SetWindowLongPtr(hEditorWnd, GWLP_WNDPROC, (LONG_PTR)EditorSubclassProc);
        InvalidateRect(hEditorWnd, NULL, TRUE);
    }
}

// Show options dialog
void ShowOptions(HWND hEditorWnd) {
    EnsureSubclassed(hEditorWnd);
    
    // Simple color picker using ChooseColor
    CHOOSECOLOR cc;
    static COLORREF acrCustClr[16];
    
    ZeroMemory(&cc, sizeof(cc));
    cc.lStructSize = sizeof(cc);
    cc.hwndOwner = hEditorWnd;
    cc.lpCustColors = (LPDWORD)acrCustClr;
    cc.rgbResult = g_WhitespaceColor;
    cc.Flags = CC_FULLOPEN | CC_RGBINIT;
    
    if (ChooseColor(&cc)) {
        g_WhitespaceColor = cc.rgbResult;
        
        // Recreate pen and brush
        if (g_hPen) DeleteObject(g_hPen);
        if (g_hBrush) DeleteObject(g_hBrush);
        g_hPen = CreatePen(PS_SOLID, 1, g_WhitespaceColor);
        g_hBrush = CreateSolidBrush(g_WhitespaceColor);
        
        InvalidateRect(hEditorWnd, NULL, TRUE);
    }
}

std::wstring g_SettingsPath;

extern "C" {
    PLUGIN_API void Initialize(const wchar_t* settingsPath) {
        g_SettingsPath = settingsPath;
        int r = GetPrivateProfileInt(L"Settings", L"ColorR", 150, g_SettingsPath.c_str());
        int g = GetPrivateProfileInt(L"Settings", L"ColorG", 150, g_SettingsPath.c_str());
        int b = GetPrivateProfileInt(L"Settings", L"ColorB", 150, g_SettingsPath.c_str());
        g_WhitespaceColor = RGB(r, g, b);
    }

    PLUGIN_API void Shutdown() {
        if (!g_SettingsPath.empty()) {
            WritePrivateProfileString(L"Settings", L"ColorR", std::to_wstring(GetRValue(g_WhitespaceColor)).c_str(), g_SettingsPath.c_str());
            WritePrivateProfileString(L"Settings", L"ColorG", std::to_wstring(GetGValue(g_WhitespaceColor)).c_str(), g_SettingsPath.c_str());
            WritePrivateProfileString(L"Settings", L"ColorB", std::to_wstring(GetBValue(g_WhitespaceColor)).c_str(), g_SettingsPath.c_str());
        }
    }

    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Whitespace Visualizer";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Visualize spaces and tabs in the editor";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.0";
    }

    PluginMenuItem g_MenuItems[] = {
        { L"Whitespace Color...", ShowOptions, NULL }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_MenuItems) / sizeof(g_MenuItems[0]);
        return g_MenuItems;
    }
    
    PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
        if (wcscmp(eventType, L"Loaded") == 0) {
            g_bIsBinary = false;
            // Check if binary
            HANDLE hFile = CreateFile(filePath, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
            if (hFile != INVALID_HANDLE_VALUE) {
                BYTE buf[1024];
                DWORD read;
                if (ReadFile(hFile, buf, sizeof(buf), &read, NULL) && read > 0) {
                    for (DWORD i = 0; i < read; i++) {
                        BYTE b = buf[i];
                        if (b < 0x20 && b != 0x09 && b != 0x0A && b != 0x0D) {
                            g_bIsBinary = true;
                            break;
                        }
                    }
                }
                CloseHandle(hFile);
            }
        }

        // Auto-enable on file load
        EnsureSubclassed(hEditor);
    }
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
    switch (ul_reason_for_call) {
        case DLL_PROCESS_DETACH:
            // Clean up resources
            if (g_hEditor && g_OldEditorProc) {
                SetWindowLongPtr(g_hEditor, GWLP_WNDPROC, (LONG_PTR)g_OldEditorProc);
            }
            if (g_hPen) DeleteObject(g_hPen);
            if (g_hBrush) DeleteObject(g_hBrush);
            break;
    }
    return TRUE;
}
