#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <string>
#include <vector>
#include <commdlg.h>
// A robust, virtual-paged Hex Viewer implemented with GDI to avoid RichEdit limitations

#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <string>
#include <vector>
#include <commdlg.h>

static const wchar_t* kFontName = L"Consolas";
static const int kFontSize = 14;
static const int kBytesPerLine = 16;

struct HexViewState {
    HANDLE hFile;
    DWORD totalSize;
    DWORD topLine; // first visible line index
    int lineHeight;
    int charWidth;
    HFONT hFont;
    // Editing state
    bool hasSelection;
    DWORD selOffset; // absolute byte offset
    int selNibble;   // 0 = high, 1 = low
    // Buffer and dirty tracking
    std::vector<BYTE> cache; // cached window of bytes
    DWORD cacheStart;        // starting offset of cache
    DWORD cacheSize;         // size of cache window
    bool dirty;              // any changes pending
};

static wchar_t g_HexLUT[16] = {L'0',L'1',L'2',L'3',L'4',L'5',L'6',L'7',L'8',L'9',L'A',L'B',L'C',L'D',L'E',L'F'};

static void FormatHexLineW(DWORD offset, const BYTE* data, DWORD len, std::wstring& out)
{
    wchar_t buf[128];
    wchar_t* p = buf;
    // Offset: %08X
    for (int shift = 28; shift >= 0; shift -= 4) *p++ = g_HexLUT[(offset >> shift) & 0xF];
    *p++ = L' ';
    *p++ = L' ';
    for (int i = 0; i < kBytesPerLine; ++i) {
        if (i < (int)len) {
            BYTE b = data[i];
            *p++ = g_HexLUT[(b >> 4) & 0xF];
            *p++ = g_HexLUT[b & 0xF];
            *p++ = L' ';
        } else {
            *p++ = L' ';
            *p++ = L' ';
            *p++ = L' ';
        }
        if (i == 7) *p++ = L' ';
    }
    *p++ = L' ';
    *p++ = L'|';
    for (DWORD i = 0; i < len; ++i) {
        unsigned char c = data[i];
        *p++ = (c >= 32 && c <= 126) ? (wchar_t)c : L'.';
    }
    *p++ = L'|';
    *p = 0;
    out.assign(buf);
}

static void MeasureFont(HWND hWnd, HexViewState* st)
{
    HDC hdc = GetDC(hWnd);
    HFONT hOld = (HFONT)SelectObject(hdc, st->hFont);
    TEXTMETRIC tm{};
    GetTextMetrics(hdc, &tm);
    st->lineHeight = tm.tmHeight + tm.tmExternalLeading;
    SIZE sz{};
    const wchar_t W = L'W';
    GetTextExtentPoint32(hdc, &W, 1, &sz);
    st->charWidth = sz.cx;
    SelectObject(hdc, hOld);
    ReleaseDC(hWnd, hdc);
}

static void UpdateScrollBar(HWND hWnd, HexViewState* st)
{
    DWORD totalLines = (st->totalSize + kBytesPerLine - 1) / kBytesPerLine;
    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = (int)totalLines;
    RECT rc; GetClientRect(hWnd, &rc);
    si.nPage = max(1, (rc.bottom / max(1, st->lineHeight)));
    si.nPos = (int)st->topLine;
    SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
}

static void PaintHex(HWND hWnd, HexViewState* st)
{
    PAINTSTRUCT ps; HDC hdc = BeginPaint(hWnd, &ps);
    RECT rc; GetClientRect(hWnd, &rc);
    FillRect(hdc, &rc, (HBRUSH)(COLOR_WINDOW+1));
    HFONT hOld = (HFONT)SelectObject(hdc, st->hFont);
    SetBkMode(hdc, TRANSPARENT);

    int linesVisible = (rc.bottom / max(1, st->lineHeight));
    DWORD byteOffset = st->topLine * kBytesPerLine;
    int xHexStart = st->charWidth * 10; // 8 offset + 2 spaces
    for (int i = 0; i < linesVisible; ++i) {
        if (byteOffset >= st->totalSize) break;
        BYTE buf[16]; DWORD remain = (st->totalSize - byteOffset);
        DWORD toRead = remain < 16 ? remain : 16; DWORD read = toRead;
        // Prefer cache if it covers this range
        if (st->cacheSize && byteOffset >= st->cacheStart && (byteOffset + toRead) <= (st->cacheStart + st->cacheSize)) {
            memcpy(buf, st->cache.data() + (byteOffset - st->cacheStart), toRead);
        } else {
            SetFilePointer(st->hFile, byteOffset, NULL, FILE_BEGIN);
            ReadFile(st->hFile, buf, toRead, &read, NULL);
        }

        std::wstring line; FormatHexLineW(byteOffset, buf, read, line);
        int y = i * st->lineHeight;
        // Draw selection highlight behind the selected byte hex if visible
        if (st->hasSelection && st->selOffset >= byteOffset && st->selOffset < byteOffset + read) {
            int idx = (int)(st->selOffset - byteOffset);
            int cell = idx * 3 + (idx >= 8 ? 1 : 0);
            RECT hi{ xHexStart + cell * st->charWidth, y, xHexStart + (cell + 2) * st->charWidth, y + st->lineHeight };
            HBRUSH hBr = CreateSolidBrush(RGB(200, 220, 255));
            FillRect(hdc, &hi, hBr);
            DeleteObject(hBr);
        }
        TextOut(hdc, 0, y, line.c_str(), (int)line.size());
        byteOffset += read;
    }

    SelectObject(hdc, hOld);
    EndPaint(hWnd, &ps);
}

static void EnsureCache(HWND hDlg, HexViewState* st)
{
    if (st->cacheSize && st->selOffset >= st->cacheStart && st->selOffset < (st->cacheStart + st->cacheSize)) {
        return;
    }
    // Save if dirty
    if (st->dirty && st->cacheSize) {
        SetFilePointer(st->hFile, st->cacheStart, NULL, FILE_BEGIN);
        DWORD written = 0; WriteFile(st->hFile, st->cache.data(), st->cacheSize, &written, NULL);
        st->dirty = false;
    }
    // Load new window
    DWORD windowStart = (st->selOffset / 4096) * 4096;
    DWORD remain = (st->totalSize > windowStart) ? (st->totalSize - windowStart) : 0;
    DWORD windowSize = (remain < 4096) ? remain : 4096;
    st->cache.resize(windowSize);
    st->cacheStart = windowStart; st->cacheSize = windowSize;
    SetFilePointer(st->hFile, windowStart, NULL, FILE_BEGIN);
    DWORD read = 0; ReadFile(st->hFile, st->cache.data(), windowSize, &read, NULL);
    InvalidateRect(hDlg, NULL, TRUE);
}

static INT_PTR CALLBACK HexViewDlg(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    HexViewState* st = (HexViewState*)GetWindowLongPtr(hDlg, GWLP_USERDATA);
    switch (msg) {
    case WM_INITDIALOG: {
        st = (HexViewState*)lParam;
        SetWindowLongPtr(hDlg, GWLP_USERDATA, (LONG_PTR)st);
        st->hFont = CreateFont(kFontSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, kFontName);
        st->hasSelection = false; st->selOffset = 0; st->selNibble = 0;
        st->dirty = false; st->cacheStart = 0; st->cacheSize = 0; st->cache.clear();
        MeasureFont(hDlg, st);
        UpdateScrollBar(hDlg, st);
        // Add Save/Discard buttons at bottom
        RECT rc; GetClientRect(hDlg, &rc);
        int btnW = 80, btnH = 24, padding = 8;
        CreateWindow(L"BUTTON", L"Save", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                     rc.right - btnW*2 - padding*2, rc.bottom - btnH - padding, btnW, btnH, hDlg, (HMENU)IDOK, GetModuleHandle(NULL), NULL);
        CreateWindow(L"BUTTON", L"Discard", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                 rc.right - btnW - padding, rc.bottom - btnH - padding, btnW, btnH, hDlg, (HMENU)IDCANCEL, GetModuleHandle(NULL), NULL);
        SetFocus(hDlg);
        return FALSE;
    }
    case WM_GETDLGCODE: {
        SetWindowLongPtr(hDlg, DWLP_MSGRESULT, DLGC_WANTALLKEYS | DLGC_WANTCHARS);
        return TRUE;
    }
    case WM_LBUTTONDOWN: {
        if (!st) break;
        SetFocus(hDlg);
        POINTS pts = MAKEPOINTS(lParam);
        POINT pt{ pts.x, pts.y };
        // Map y to line
        int lineIdx = pt.y / max(1, st->lineHeight);
        DWORD lineStartOffset = (st->topLine + lineIdx) * kBytesPerLine;
        if (lineStartOffset >= st->totalSize) break;
        // Compute hex area start
        int xHexStart = st->charWidth * 10;
        if (pt.x < xHexStart) break;
        int rel = (pt.x - xHexStart) / st->charWidth; // character cell index
        // Map rel to byte index: pattern 0..(with spaces)
        // Each byte takes 3 cells; after 8 bytes an extra space (+1)
        int idx = -1;
        int cell = 0;
        for (int i = 0; i < 16; ++i) {
            int startCell = cell;
            int endCell = cell + 2; // hex chars span two cells
            if (rel >= startCell && rel <= endCell) { idx = i; break; }
            cell += 3;
            if (i == 7) cell += 1;
        }
        if (idx >= 0) {
            DWORD sel = lineStartOffset + idx;
            if (sel < st->totalSize) {
                st->hasSelection = true;
                st->selOffset = sel;
                // Determine nibble from click position within the two cells
                int nibbleRel = rel - (idx * 3 + (idx >= 8 ? 1 : 0));
                st->selNibble = (nibbleRel <= 0) ? 0 : 1;
                // Ensure cache covers this selection
                EnsureCache(hDlg, st);
            }
        }
        InvalidateRect(hDlg, NULL, TRUE);
        return TRUE;
    }
    // WM_CHAR removed as it is handled in WM_KEYDOWN
    case WM_KEYDOWN: {
        if (!st) break;

        // Handle Hex Input
        if (st->hasSelection) {
            int val = -1;
            if (wParam >= '0' && wParam <= '9') val = wParam - '0';
            else if (wParam >= 'A' && wParam <= 'F') val = 10 + (wParam - 'A');
            else if (wParam >= VK_NUMPAD0 && wParam <= VK_NUMPAD9) val = wParam - VK_NUMPAD0;
            
            if (val >= 0) {
                EnsureCache(hDlg, st);
                if (st->cacheSize && st->selOffset >= st->cacheStart && st->selOffset < (st->cacheStart + st->cacheSize)) {
                    DWORD cacheIndex = st->selOffset - st->cacheStart;
                    BYTE b = st->cache[cacheIndex];
                    if (st->selNibble == 0) { b = (BYTE)((b & 0x0F) | (val << 4)); st->selNibble = 1; }
                    else { b = (BYTE)((b & 0xF0) | val); st->selNibble = 0; /* advance to next byte */ if (st->selOffset + 1 < st->totalSize) st->selOffset++; }
                    st->cache[cacheIndex] = b;
                    st->dirty = true;
                    InvalidateRect(hDlg, NULL, TRUE);
                }
                return TRUE;
            }
        }

        switch (wParam) {
            case VK_LEFT:
                if (st->hasSelection && st->selNibble == 1) st->selNibble = 0;
                else if (st->hasSelection && st->selOffset > 0) { st->selOffset--; st->selNibble = 1; }
                EnsureCache(hDlg, st); InvalidateRect(hDlg, NULL, TRUE); return TRUE;
            case VK_RIGHT:
                if (st->hasSelection && st->selNibble == 0) st->selNibble = 1;
                else if (st->hasSelection && st->selOffset + 1 < st->totalSize) { st->selOffset++; st->selNibble = 0; }
                EnsureCache(hDlg, st); InvalidateRect(hDlg, NULL, TRUE); return TRUE;
            case VK_UP:
                if (st->hasSelection && st->selOffset >= kBytesPerLine) st->selOffset -= kBytesPerLine;
                EnsureCache(hDlg, st); InvalidateRect(hDlg, NULL, TRUE); return TRUE;
            case VK_DOWN:
                if (st->hasSelection && st->selOffset + kBytesPerLine < st->totalSize) st->selOffset += kBytesPerLine;
                EnsureCache(hDlg, st); InvalidateRect(hDlg, NULL, TRUE); return TRUE;
            case VK_HOME:
                if (st->hasSelection) st->selOffset = (st->selOffset / kBytesPerLine) * kBytesPerLine;
                EnsureCache(hDlg, st); InvalidateRect(hDlg, NULL, TRUE); return TRUE;
            case VK_END:
                if (st->hasSelection) {
                    DWORD lineStart = (st->selOffset / kBytesPerLine) * kBytesPerLine;
                    DWORD lineEnd = min(lineStart + kBytesPerLine - 1, st->totalSize ? st->totalSize - 1 : 0);
                    st->selOffset = lineEnd;
                }
                EnsureCache(hDlg, st); InvalidateRect(hDlg, NULL, TRUE); return TRUE;
        }
        break;
    }
    case WM_VSCROLL: {
        if (!st) break;
        SCROLLINFO si{}; si.cbSize = sizeof(si); si.fMask = SIF_ALL; GetScrollInfo(hDlg, SB_VERT, &si);
        int pos = si.nPos;
        switch (LOWORD(wParam)) {
            case SB_LINEUP: pos -= 1; break;
            case SB_LINEDOWN: pos += 1; break;
            case SB_PAGEUP: pos -= (int)si.nPage; break;
            case SB_PAGEDOWN: pos += (int)si.nPage; break;
            case SB_THUMBTRACK: pos = (int)si.nTrackPos; break;
        }
        if (pos < si.nMin) pos = si.nMin;
        if (pos > si.nMax - (int)si.nPage) pos = si.nMax - (int)si.nPage;
        st->topLine = (DWORD)pos;
        si.fMask = SIF_POS; si.nPos = pos; SetScrollInfo(hDlg, SB_VERT, &si, TRUE);
        InvalidateRect(hDlg, NULL, TRUE);
        return TRUE;
    }
    case WM_SIZE: {
        if (st) { UpdateScrollBar(hDlg, st); InvalidateRect(hDlg, NULL, TRUE); }
        return TRUE;
    }
    case WM_PAINT: {
        if (st) PaintHex(hDlg, st);
        return 0;
    }
    case WM_COMMAND: {
        if (LOWORD(wParam) == IDOK) {
            // Save cache window back to file if dirty
            if (st && st->dirty && st->cacheSize) {
                SetFilePointer(st->hFile, st->cacheStart, NULL, FILE_BEGIN);
                DWORD written = 0; WriteFile(st->hFile, st->cache.data(), st->cacheSize, &written, NULL);
                st->dirty = false;
            }
            if (st) {
                if (st->hFile && st->hFile != INVALID_HANDLE_VALUE) CloseHandle(st->hFile);
                if (st->hFont) DeleteObject(st->hFont);
                delete st; SetWindowLongPtr(hDlg, GWLP_USERDATA, 0);
            }
            EndDialog(hDlg, IDOK); return TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL) {
            // Discard changes: do not write cache back
            if (st) {
                if (st->hFile && st->hFile != INVALID_HANDLE_VALUE) CloseHandle(st->hFile);
                if (st->hFont) DeleteObject(st->hFont);
                delete st;
                SetWindowLongPtr(hDlg, GWLP_USERDATA, 0);
            }
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
        }
        break;
    }
    }
    return FALSE;
}

// Helper to align data on DWORD boundary
static LPWORD lpwAlign(LPWORD lpIn)
{
    ULONG ul = (ULONG)(ULONG_PTR)lpIn; ul += 3; ul >>= 2; ul <<= 2; return (LPWORD)(ULONG_PTR)ul;
}

static void ShowHexWindow(HWND hParent, HexViewState* st)
{
    HGLOBAL hgbl = GlobalAlloc(GMEM_ZEROINIT, 2048);
    if (!hgbl) return;
    LPDLGTEMPLATE lpdt = (LPDLGTEMPLATE)GlobalLock(hgbl);
    lpdt->style = WS_POPUP | WS_BORDER | WS_SYSMENU | WS_CAPTION | DS_MODALFRAME | DS_CENTER | WS_THICKFRAME | WS_VISIBLE | WS_VSCROLL;
    lpdt->cdit = 0;
    lpdt->x = 10; lpdt->y = 10; lpdt->cx = 500; lpdt->cy = 400;
    LPWORD lpw = (LPWORD)(lpdt + 1);
    *lpw++ = 0; *lpw++ = 0;
    LPWSTR lpwsz = (LPWSTR)lpw; wcscpy(lpwsz, L"Hex Viewer"); lpw += wcslen(lpwsz) + 1;
    GlobalUnlock(hgbl);
    DialogBoxIndirectParam(GetModuleHandle(NULL), (LPDLGTEMPLATE)hgbl, hParent, HexViewDlg, (LPARAM)st);
    GlobalFree(hgbl);
}

static bool IsBinaryBuffer(const BYTE* buf, DWORD read)
{
    for (DWORD i = 0; i < read; ++i) {
        BYTE b = buf[i];
        if (b < 0x20 && b != 0x09 && b != 0x0A && b != 0x0D) return true;
    }
    return false;
}

static std::wstring g_CurrentFile;

static void OpenHexEditor(HWND hEditor) {
    if (g_CurrentFile.empty()) {
        MessageBox(hEditor, L"No file loaded.", L"Hex Editor", MB_ICONWARNING);
        return;
    }
    HANDLE hFile = CreateFile(g_CurrentFile.c_str(), GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) {
        MessageBox(hEditor, L"Could not open file.", L"Hex Editor", MB_ICONERROR);
        return;
    }
    DWORD size = GetFileSize(hFile, NULL);
    if (size == 0) { CloseHandle(hFile); return; }
    
    HexViewState* st = new HexViewState();
    st->hFile = hFile; st->totalSize = size; st->topLine = 0; st->lineHeight = 18; st->charWidth = 8; st->hFont = NULL;
    HWND hParent = GetAncestor(hEditor, GA_ROOT);
    ShowHexWindow(hParent ? hParent : hEditor, st);
}

static PluginMenuItem g_MenuItems[] = {
    { L"Open in Hex Editor", OpenHexEditor, L"Ctrl+Shift+H" }
};

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() { return L"Hex Editor"; }
    PLUGIN_API const wchar_t* GetPluginDescription() { return L"View files in a Hex format (read-only)."; }
    PLUGIN_API const wchar_t* GetPluginVersion() { return L"2.0"; }

    PLUGIN_API const wchar_t* GetPluginLicense() {
        return L"MIT License\n\n"
               L"Copyright (c) 2025 Just Notepad Contributors\n\n"
               L"Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\n\n"
               L"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n\n"
               L"THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO, THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
    }
    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) { 
        *count = 1; 
        return g_MenuItems; 
    }

    PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
        if (_wcsicmp(eventType, L"Loaded") == 0) {
            g_CurrentFile = filePath;
        } else if (_wcsicmp(eventType, L"New") == 0) {
            g_CurrentFile.clear();
        } else if (_wcsicmp(eventType, L"Saved") == 0) {
            g_CurrentFile = filePath;
        }
    }

    PLUGIN_API bool OnSaveFile(const wchar_t* filePath, HWND hEditor) { return false; }
}

BOOL APIENTRY DllMain(HMODULE, DWORD, LPVOID) { return TRUE; }
