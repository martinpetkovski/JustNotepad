#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <vector>
#include <string>
#include <algorithm>
#include "resource.h"

HINSTANCE g_hInst = NULL;

struct UnicodeRange {
    const wchar_t* name;
    unsigned int start;
    unsigned int end;
};

UnicodeRange g_Ranges[] = {
    { L"Basic Latin", 0x0020, 0x007F },
    { L"Latin-1 Supplement", 0x0080, 0x00FF },
    { L"Latin Extended-A", 0x0100, 0x017F },
    { L"Latin Extended-B", 0x0180, 0x024F },
    { L"IPA Extensions", 0x0250, 0x02AF },
    { L"Spacing Modifier Letters", 0x02B0, 0x02FF },
    { L"Combining Diacritical Marks", 0x0300, 0x036F },
    { L"Greek and Coptic", 0x0370, 0x03FF },
    { L"Cyrillic", 0x0400, 0x04FF },
    { L"Cyrillic Supplement", 0x0500, 0x052F },
    { L"Armenian", 0x0530, 0x058F },
    { L"Hebrew", 0x0590, 0x05FF },
    { L"Arabic", 0x0600, 0x06FF },
    { L"Syriac", 0x0700, 0x074F },
    { L"Thaana", 0x0780, 0x07BF },
    { L"Devanagari", 0x0900, 0x097F },
    { L"Bengali", 0x0980, 0x09FF },
    { L"Gurmukhi", 0x0A00, 0x0A7F },
    { L"Gujarati", 0x0A80, 0x0AFF },
    { L"Oriya", 0x0B00, 0x0B7F },
    { L"Tamil", 0x0B80, 0x0BFF },
    { L"Telugu", 0x0C00, 0x0C7F },
    { L"Kannada", 0x0C80, 0x0CFF },
    { L"Malayalam", 0x0D00, 0x0D7F },
    { L"Sinhala", 0x0D80, 0x0DFF },
    { L"Thai", 0x0E00, 0x0E7F },
    { L"Lao", 0x0E80, 0x0EFF },
    { L"Tibetan", 0x0F00, 0x0FFF },
    { L"Myanmar", 0x1000, 0x109F },
    { L"Georgian", 0x10A0, 0x10FF },
    { L"Hangul Jamo", 0x1100, 0x11FF },
    { L"Ethiopic", 0x1200, 0x137F },
    { L"Cherokee", 0x13A0, 0x13FF },
    { L"Unified Canadian Aboriginal Syllabics", 0x1400, 0x167F },
    { L"Ogham", 0x1680, 0x169F },
    { L"Runic", 0x16A0, 0x16FF },
    { L"Tagalog", 0x1700, 0x171F },
    { L"Hanunoo", 0x1720, 0x173F },
    { L"Buhid", 0x1740, 0x175F },
    { L"Tagbanwa", 0x1760, 0x177F },
    { L"Khmer", 0x1780, 0x17FF },
    { L"Mongolian", 0x1800, 0x18AF },
    { L"Limbu", 0x1900, 0x194F },
    { L"Tai Le", 0x1950, 0x197F },
    { L"Khmer Symbols", 0x19E0, 0x19FF },
    { L"Phonetic Extensions", 0x1D00, 0x1D7F },
    { L"Latin Extended Additional", 0x1E00, 0x1EFF },
    { L"Greek Extended", 0x1F00, 0x1FFF },
    { L"General Punctuation", 0x2000, 0x206F },
    { L"Superscripts and Subscripts", 0x2070, 0x209F },
    { L"Currency Symbols", 0x20A0, 0x20CF },
    { L"Combining Diacritical Marks for Symbols", 0x20D0, 0x20FF },
    { L"Letterlike Symbols", 0x2100, 0x214F },
    { L"Number Forms", 0x2150, 0x218F },
    { L"Arrows", 0x2190, 0x21FF },
    { L"Mathematical Operators", 0x2200, 0x22FF },
    { L"Miscellaneous Technical", 0x2300, 0x23FF },
    { L"Control Pictures", 0x2400, 0x243F },
    { L"Optical Character Recognition", 0x2440, 0x245F },
    { L"Enclosed Alphanumerics", 0x2460, 0x24FF },
    { L"Box Drawing", 0x2500, 0x257F },
    { L"Block Elements", 0x2580, 0x259F },
    { L"Geometric Shapes", 0x25A0, 0x25FF },
    { L"Miscellaneous Symbols", 0x2600, 0x26FF },
    { L"Dingbats", 0x2700, 0x27BF },
    { L"Miscellaneous Mathematical Symbols-A", 0x27C0, 0x27EF },
    { L"Supplemental Arrows-A", 0x27F0, 0x27FF },
    { L"Braille Patterns", 0x2800, 0x28FF },
    { L"Supplemental Arrows-B", 0x2900, 0x297F },
    { L"Miscellaneous Mathematical Symbols-B", 0x2980, 0x29FF },
    { L"Supplemental Mathematical Operators", 0x2A00, 0x2AFF },
    { L"Miscellaneous Symbols and Arrows", 0x2B00, 0x2BFF },
    { L"CJK Radicals Supplement", 0x2E80, 0x2EFF },
    { L"Kangxi Radicals", 0x2F00, 0x2FDF },
    { L"Ideographic Description Characters", 0x2FF0, 0x2FFF },
    { L"CJK Symbols and Punctuation", 0x3000, 0x303F },
    { L"Hiragana", 0x3040, 0x309F },
    { L"Katakana", 0x30A0, 0x30FF },
    { L"Bopomofo", 0x3100, 0x312F },
    { L"Hangul Compatibility Jamo", 0x3130, 0x318F },
    { L"Kanbun", 0x3190, 0x319F },
    { L"Bopomofo Extended", 0x31A0, 0x31BF },
    { L"Katakana Phonetic Extensions", 0x31F0, 0x31FF },
    { L"Enclosed CJK Letters and Months", 0x3200, 0x32FF },
    { L"CJK Compatibility", 0x3300, 0x33FF },
    { L"CJK Unified Ideographs Extension A", 0x3400, 0x4DBF },
    { L"Yijing Hexagram Symbols", 0x4DC0, 0x4DFF },
    { L"CJK Unified Ideographs", 0x4E00, 0x9FFF },
    { L"Yi Syllables", 0xA000, 0xA48F },
    { L"Yi Radicals", 0xA490, 0xA4CF },
    { L"Hangul Syllables", 0xAC00, 0xD7AF },
    { L"High Surrogates", 0xD800, 0xDB7F },
    { L"High Private Use Surrogates", 0xDB80, 0xDBFF },
    { L"Low Surrogates", 0xDC00, 0xDFFF },
    { L"Private Use Area", 0xE000, 0xF8FF },
    { L"CJK Compatibility Ideographs", 0xF900, 0xFAFF },
    { L"Alphabetic Presentation Forms", 0xFB00, 0xFB4F },
    { L"Arabic Presentation Forms-A", 0xFB50, 0xFDFF },
    { L"Variation Selectors", 0xFE00, 0xFE0F },
    { L"Vertical Forms", 0xFE10, 0xFE1F },
    { L"Combining Half Marks", 0xFE20, 0xFE2F },
    { L"CJK Compatibility Forms", 0xFE30, 0xFE4F },
    { L"Small Form Variants", 0xFE50, 0xFE6F },
    { L"Arabic Presentation Forms-B", 0xFE70, 0xFEFF },
    { L"Halfwidth and Fullwidth Forms", 0xFF00, 0xFFEF },
    { L"Specials", 0xFFF0, 0xFFFF }
};

HFONT g_hFont = NULL;
int g_SelectedChar = -1;
int g_CurrentRangeIndex = 0;
WNDPROC g_OldListBoxProc = NULL;
std::wstring g_SettingsPath;
std::wstring g_LastFont = L"Segoe UI";

extern "C" {
    PLUGIN_API void Initialize(const wchar_t* settingsPath) {
        g_SettingsPath = settingsPath;
        g_CurrentRangeIndex = GetPrivateProfileInt(L"Settings", L"RangeIndex", 0, g_SettingsPath.c_str());
        
        WCHAR buf[LF_FACESIZE];
        GetPrivateProfileString(L"Settings", L"Font", L"Segoe UI", buf, LF_FACESIZE, g_SettingsPath.c_str());
        g_LastFont = buf;
    }

    PLUGIN_API void Shutdown() {
        if (!g_SettingsPath.empty()) {
            WritePrivateProfileString(L"Settings", L"RangeIndex", std::to_wstring(g_CurrentRangeIndex).c_str(), g_SettingsPath.c_str());
            WritePrivateProfileString(L"Settings", L"Font", g_LastFont.c_str(), g_SettingsPath.c_str());
        }
    }
}

#define COLS 16
#define CELL_SIZE 24

int GetCharFromPoint(HWND hList, POINT pt) {
    RECT rcClient;
    GetClientRect(hList, &rcClient);
    
    int rowHeight = (int)SendMessage(hList, LB_GETITEMHEIGHT, 0, 0);
    if (rowHeight == 0) return -1;
    
    int topIndex = (int)SendMessage(hList, LB_GETTOPINDEX, 0, 0);
    
    int row = topIndex + pt.y / rowHeight;
    int col = pt.x / CELL_SIZE;
    
    if (col >= COLS) return -1;
    
    int charIndex = row * COLS + col;
    int rangeStart = g_Ranges[g_CurrentRangeIndex].start;
    int rangeEnd = g_Ranges[g_CurrentRangeIndex].end;
    int count = rangeEnd - rangeStart + 1;
    
    if (charIndex >= count) return -1;
    
    return rangeStart + charIndex;
}

LRESULT CALLBACK ListBoxSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_LBUTTONDOWN: {
            POINT pt = { LOWORD(lParam), HIWORD(lParam) };
            int charCode = GetCharFromPoint(hwnd, pt);
            if (charCode != -1) {
                g_SelectedChar = charCode;
                InvalidateRect(hwnd, NULL, FALSE);
                
                // Handle double click behavior if needed, but standard listbox doesn't do double click on cells easily
            }
            break;
        }
        case WM_LBUTTONDBLCLK: {
            POINT pt = { LOWORD(lParam), HIWORD(lParam) };
            int charCode = GetCharFromPoint(hwnd, pt);
            if (charCode != -1) {
                g_SelectedChar = charCode;
                SendMessage(GetParent(hwnd), WM_COMMAND, IDOK, 0);
            }
            break;
        }
    }
    return CallWindowProc(g_OldListBoxProc, hwnd, uMsg, wParam, lParam);
}

int CALLBACK EnumFontFamExProc(const LOGFONT *lpelfe, const TEXTMETRIC *lpntme, DWORD FontType, LPARAM lParam) {
    HWND hCombo = (HWND)lParam;
    if (SendMessage(hCombo, CB_FINDSTRINGEXACT, -1, (LPARAM)lpelfe->lfFaceName) == CB_ERR) {
        SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)lpelfe->lfFaceName);
    }
    return 1;
}

void UpdateGrid(HWND hDlg) {
    HWND hList = GetDlgItem(hDlg, IDC_CHAR_GRID);
    int rangeStart = g_Ranges[g_CurrentRangeIndex].start;
    int rangeEnd = g_Ranges[g_CurrentRangeIndex].end;
    int count = rangeEnd - rangeStart + 1;
    int rows = (count + COLS - 1) / COLS;
    
    SendMessage(hList, LB_RESETCONTENT, 0, 0);
    SendMessage(hList, LB_SETCOUNT, rows, 0);
    
    // Reset selection if out of range
    if (g_SelectedChar < rangeStart || g_SelectedChar > rangeEnd) {
        g_SelectedChar = -1;
    }
    InvalidateRect(hList, NULL, FALSE);
}

INT_PTR CALLBACK CharMapDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
    {
        // Populate Fonts
        HWND hFontCombo = GetDlgItem(hDlg, IDC_FONT_COMBO);
        HDC hdc = GetDC(hDlg);
        LOGFONT lf = { 0 };
        lf.lfCharSet = DEFAULT_CHARSET;
        EnumFontFamiliesEx(hdc, &lf, (FONTENUMPROC)EnumFontFamExProc, (LPARAM)hFontCombo, 0);
        ReleaseDC(hDlg, hdc);
        
        // Select default font (Segoe UI or Arial)
        if (SendMessage(hFontCombo, CB_SELECTSTRING, -1, (LPARAM)g_LastFont.c_str()) == CB_ERR) {
            SendMessage(hFontCombo, CB_SETCURSEL, 0, 0);
        }
        
        // Populate Ranges
        HWND hRangeCombo = GetDlgItem(hDlg, IDC_RANGE_COMBO);
        for (int i = 0; i < sizeof(g_Ranges) / sizeof(g_Ranges[0]); ++i) {
            SendMessage(hRangeCombo, CB_ADDSTRING, 0, (LPARAM)g_Ranges[i].name);
        }
        SendMessage(hRangeCombo, CB_SETCURSEL, g_CurrentRangeIndex, 0);
        
        // Subclass ListBox
        HWND hList = GetDlgItem(hDlg, IDC_CHAR_GRID);
        g_OldListBoxProc = (WNDPROC)SetWindowLongPtr(hList, GWLP_WNDPROC, (LONG_PTR)ListBoxSubclassProc);
        
        UpdateGrid(hDlg);
        
        return TRUE;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_FONT_COMBO && HIWORD(wParam) == CBN_SELCHANGE) {
            HWND hFontCombo = GetDlgItem(hDlg, IDC_FONT_COMBO);
            int idx = (int)SendMessage(hFontCombo, CB_GETCURSEL, 0, 0);
            if (idx != CB_ERR) {
                wchar_t szFontName[LF_FACESIZE];
                SendMessage(hFontCombo, CB_GETLBTEXT, idx, (LPARAM)szFontName);
                g_LastFont = szFontName;
                
                if (g_hFont) DeleteObject(g_hFont);
                g_hFont = CreateFont(20, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 
                    OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, szFontName);
                
                InvalidateRect(GetDlgItem(hDlg, IDC_CHAR_GRID), NULL, FALSE);
            }
        }
        else if (LOWORD(wParam) == IDC_RANGE_COMBO && HIWORD(wParam) == CBN_SELCHANGE) {
            HWND hRangeCombo = GetDlgItem(hDlg, IDC_RANGE_COMBO);
            int idx = (int)SendMessage(hRangeCombo, CB_GETCURSEL, 0, 0);
            if (idx != CB_ERR) {
                g_CurrentRangeIndex = idx;
                UpdateGrid(hDlg);
            }
        }
        else if (LOWORD(wParam) == IDOK) {
            if (g_SelectedChar != -1) {
                EndDialog(hDlg, (INT_PTR)g_SelectedChar);
            }
            return TRUE;
        }
        else if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, -1);
            return TRUE;
        }
        break;
        
    case WM_MEASUREITEM:
        if (wParam == IDC_CHAR_GRID) {
            LPMEASUREITEMSTRUCT lpmis = (LPMEASUREITEMSTRUCT)lParam;
            lpmis->itemHeight = CELL_SIZE;
        }
        break;
        
    case WM_DRAWITEM:
        if (wParam == IDC_CHAR_GRID) {
            LPDRAWITEMSTRUCT lpdis = (LPDRAWITEMSTRUCT)lParam;
            
            if (lpdis->itemID == -1) return TRUE;
            
            HDC hdc = lpdis->hDC;
            int row = lpdis->itemID;
            int rangeStart = g_Ranges[g_CurrentRangeIndex].start;
            int rangeEnd = g_Ranges[g_CurrentRangeIndex].end;
            int count = rangeEnd - rangeStart + 1;
            
            HFONT hOldFont = (HFONT)SelectObject(hdc, g_hFont ? g_hFont : GetStockObject(DEFAULT_GUI_FONT));
            
            // Fill background
            FillRect(hdc, &lpdis->rcItem, (HBRUSH)GetStockObject(WHITE_BRUSH));
            
            for (int col = 0; col < COLS; ++col) {
                int charIndex = row * COLS + col;
                if (charIndex >= count) break;
                
                int charCode = rangeStart + charIndex;
                
                RECT rcCell;
                rcCell.left = lpdis->rcItem.left + col * CELL_SIZE;
                rcCell.top = lpdis->rcItem.top;
                rcCell.right = rcCell.left + CELL_SIZE;
                rcCell.bottom = rcCell.top + CELL_SIZE;
                
                // Draw Selection
                if (charCode == g_SelectedChar) {
                    HBRUSH hBr = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT));
                    FillRect(hdc, &rcCell, hBr);
                    DeleteObject(hBr);
                    SetTextColor(hdc, GetSysColor(COLOR_HIGHLIGHTTEXT));
                } else {
                    SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));
                    
                    // Draw grid lines
                    HPEN hPen = CreatePen(PS_SOLID, 1, RGB(200, 200, 200));
                    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                    MoveToEx(hdc, rcCell.right - 1, rcCell.top, NULL);
                    LineTo(hdc, rcCell.right - 1, rcCell.bottom);
                    MoveToEx(hdc, rcCell.left, rcCell.bottom - 1, NULL);
                    LineTo(hdc, rcCell.right, rcCell.bottom - 1);
                    SelectObject(hdc, hOldPen);
                    DeleteObject(hPen);
                }
                
                SetBkMode(hdc, TRANSPARENT);
                
                wchar_t ch[2] = { (wchar_t)charCode, 0 };
                DrawText(hdc, ch, 1, &rcCell, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            }
            
            SelectObject(hdc, hOldFont);
            return TRUE;
        }
        break;
        
    case WM_DESTROY:
        if (g_hFont) DeleteObject(g_hFont);
        // Unsubclass
        if (g_OldListBoxProc) {
            SetWindowLongPtr(GetDlgItem(hDlg, IDC_CHAR_GRID), GWLP_WNDPROC, (LONG_PTR)g_OldListBoxProc);
        }
        break;
    }
    return FALSE;
}

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Character Map";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Insert Unicode characters.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.0";
    }

    void OpenCharacterMap(HWND hEditorWnd) {
        INT_PTR result = DialogBox(g_hInst, MAKEINTRESOURCE(IDD_CHARMAP), hEditorWnd, CharMapDlgProc);
        if (result != -1) {
            wchar_t ch[2] = { (wchar_t)result, 0 };
            SendMessage(hEditorWnd, EM_REPLACESEL, TRUE, (LPARAM)ch);
        }
    }

    PluginMenuItem g_Items[] = {
        { L"Character Map...", OpenCharacterMap }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_Items) / sizeof(PluginMenuItem);
        return g_Items;
    }
}

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
