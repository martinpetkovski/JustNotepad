#ifndef UNICODE
#define UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include "theme.h"

#include <dwmapi.h>
#include <uxtheme.h>
#include <vssym32.h>
#include <commctrl.h>
#include <string>

#include <algorithm>
namespace Gdiplus {
using std::max;
using std::min;
}  // namespace Gdiplus
#include <objidl.h>
#include <gdiplus.h>
#pragma comment(lib, "gdiplus.lib")

#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")

namespace {

// uxtheme.dll, by ordinal (Windows 10 1809 and later).
enum PreferredAppMode { AppModeDefault, AllowDark, ForceDark, ForceLight };
using SetPreferredAppModeFn = PreferredAppMode(WINAPI*)(PreferredAppMode);
using AllowDarkModeForWindowFn = BOOL(WINAPI*)(HWND, BOOL);
using FlushMenuThemesFn = void(WINAPI*)();
using RefreshPolicyFn = void(WINAPI*)();

SetPreferredAppModeFn pSetPreferredAppMode = nullptr;
AllowDarkModeForWindowFn pAllowDarkModeForWindow = nullptr;
FlushMenuThemesFn pFlushMenuThemes = nullptr;
RefreshPolicyFn pRefreshPolicy = nullptr;

HBRUSH g_faceBrush = nullptr, g_bgBrush = nullptr, g_hotBrush = nullptr;
COLORREF g_faceCol = CLR_INVALID, g_bgCol = CLR_INVALID, g_hotCol = CLR_INVALID;

HBRUSH brushFor(HBRUSH& b, COLORREF& have, COLORREF want) {
    if (!b || have != want) {
        if (b) DeleteObject(b);
        b = CreateSolidBrush(want);
        have = want;
    }
    return b;
}

// Undocumented messages Windows sends to let an app draw its own menu bar.
const UINT WM_UAHDRAWMENU = 0x0091, WM_UAHDRAWMENUITEM = 0x0092;
struct UAHMENU {
    HMENU hmenu;
    HDC hdc;
    DWORD dwFlags;
};
union UAHMENUITEMMETRICS {
    struct { DWORD cx, cy; } rgsizeBar[2];
    struct { DWORD cx, cy; } rgsizePopup[4];
};
struct UAHMENUPOPUPMETRICS {
    DWORD rgcx[4];
    DWORD fUpdateMaxWidths : 2;
};
struct UAHMENUITEM {
    int iPosition;
    UAHMENUITEMMETRICS umim;
    UAHMENUPOPUPMETRICS umpm;
};
struct UAHDRAWMENUITEM {
    DRAWITEMSTRUCT dis;
    UAHMENU um;
    UAHMENUITEM umi;
};

// The menu bar area in window coordinates.
bool menuBarRect(HWND hwnd, RECT& r) {
    MENUBARINFO mbi = {sizeof(mbi)};
    if (!GetMenuBarInfo(hwnd, OBJID_MENU, 0, &mbi)) return false;
    RECT wr;
    GetWindowRect(hwnd, &wr);
    r = mbi.rcBar;
    OffsetRect(&r, -wr.left, -wr.top);
    return true;
}

// Windows draws a one-pixel light line under the menu bar; cover it.
void paintMenuBarLine(HWND hwnd, const Palette& p) {
    RECT rc;
    GetClientRect(hwnd, &rc);
    MapWindowPoints(hwnd, nullptr, (POINT*)&rc, 2);
    RECT wr;
    GetWindowRect(hwnd, &wr);
    OffsetRect(&rc, -wr.left, -wr.top);
    RECT line = rc;
    line.bottom = line.top;
    line.top--;
    HDC dc = GetWindowDC(hwnd);
    FillRect(dc, &line, brushFor(g_faceBrush, g_faceCol, p.face));
    ReleaseDC(hwnd, dc);
}

// Windows 11 rounds the corners of windows and menus; ask for square ones.
// (33 = DWMWA_WINDOW_CORNER_PREFERENCE, 1 = DWMWCP_DONOTROUND; ignored before Windows 11.)
void squareCorners(HWND hwnd) {
    DWORD pref = 1;
    DwmSetWindowAttribute(hwnd, 33, &pref, sizeof(pref));
}

// Menus are windows of class "#32768" that Windows creates on our thread;
// catch each one as it is created and square it too.
HHOOK g_menuHook = nullptr;
LRESULT CALLBACK menuHookProc(int code, WPARAM wp, LPARAM lp) {
    if (code == HC_ACTION) {
        const CWPRETSTRUCT* m = (const CWPRETSTRUCT*)lp;
        if (m->message == WM_CREATE) {
            wchar_t cls[16] = {};
            GetClassNameW(m->hwnd, cls, 16);
            if (lstrcmpW(cls, L"#32768") == 0) squareCorners(m->hwnd);
        }
    }
    return CallNextHookEx(g_menuHook, code, wp, lp);
}

}  // namespace

Palette makePalette(bool dark) {
    Palette p;
    p.dark = dark;
    if (dark) {
        p.bg = RGB(30, 30, 30);
        p.fg = RGB(212, 212, 212);
        p.face = RGB(43, 43, 43);
        p.faceText = RGB(222, 222, 222);
        p.border = RGB(64, 64, 64);
        p.hot = RGB(66, 66, 66);
        p.accent = RGB(76, 160, 255);
        p.selection = RGB(38, 79, 120);
    } else {
        p.bg = RGB(255, 255, 255);
        p.fg = RGB(28, 28, 28);
        p.face = RGB(243, 243, 243);
        p.faceText = RGB(28, 28, 28);
        p.border = RGB(214, 214, 214);
        p.hot = RGB(224, 236, 250);
        p.accent = RGB(0, 103, 192);
        p.selection = RGB(173, 214, 255);
    }
    return p;
}

bool systemUsesDarkMode() {
    DWORD v = 1, size = sizeof(v);
    if (RegGetValueW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                     L"AppsUseLightTheme", RRF_RT_REG_DWORD, nullptr, &v, &size) != ERROR_SUCCESS)
        return false;
    return v == 0;
}

void themeInit() {
    HMODULE ux = LoadLibraryExW(L"uxtheme.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!ux) return;
    pRefreshPolicy = (RefreshPolicyFn)(void*)GetProcAddress(ux, MAKEINTRESOURCEA(104));
    pAllowDarkModeForWindow = (AllowDarkModeForWindowFn)(void*)GetProcAddress(ux, MAKEINTRESOURCEA(133));
    pSetPreferredAppMode = (SetPreferredAppModeFn)(void*)GetProcAddress(ux, MAKEINTRESOURCEA(135));
    pFlushMenuThemes = (FlushMenuThemesFn)(void*)GetProcAddress(ux, MAKEINTRESOURCEA(136));
    if (!g_menuHook) g_menuHook = SetWindowsHookExW(WH_CALLWNDPROCRET, menuHookProc, nullptr, GetCurrentThreadId());
}

void themeSetAppDark(bool dark) {
    if (pSetPreferredAppMode) pSetPreferredAppMode(dark ? ForceDark : ForceLight);
    if (pRefreshPolicy) pRefreshPolicy();
    if (pFlushMenuThemes) pFlushMenuThemes();
}

void themeWindow(HWND hwnd, bool dark) {
    if (pAllowDarkModeForWindow) pAllowDarkModeForWindow(hwnd, dark);
    BOOL on = dark;
    // 20 = DWMWA_USE_IMMERSIVE_DARK_MODE; 19 on Windows 10 before 20H1.
    if (FAILED(DwmSetWindowAttribute(hwnd, 20, &on, sizeof(on)))) DwmSetWindowAttribute(hwnd, 19, &on, sizeof(on));
    squareCorners(hwnd);
    SetWindowTheme(hwnd, dark ? L"DarkMode_Explorer" : nullptr, nullptr);
    SendMessageW(hwnd, WM_THEMECHANGED, 0, 0);
}

void themeDialog(HWND dlg, const Palette& p) {
    themeWindow(dlg, p.dark);
    squareCorners(dlg);
    for (HWND c = GetWindow(dlg, GW_CHILD); c; c = GetWindow(c, GW_HWNDNEXT)) {
        wchar_t cls[32] = {};
        GetClassNameW(c, cls, 32);
        LONG style = GetWindowLongW(c, GWL_STYLE);
        if (pAllowDarkModeForWindow) pAllowDarkModeForWindow(c, p.dark);
        if (lstrcmpiW(cls, L"Button") == 0) {
            LONG kind = style & BS_TYPEMASK;
            bool textual = kind == BS_CHECKBOX || kind == BS_AUTOCHECKBOX || kind == BS_RADIOBUTTON ||
                           kind == BS_AUTORADIOBUTTON || kind == BS_GROUPBOX || kind == BS_3STATE ||
                           kind == BS_AUTO3STATE;
            // Themed check boxes ignore WM_CTLCOLORSTATIC text colour, so in
            // dark mode they are drawn unthemed (still crisp, and readable).
            // Classic (unthemed) controls are square in every Windows version.
            if (textual) {
                SetWindowTheme(c, L"", L"");
            } else if (kind == BS_PUSHBUTTON || kind == BS_DEFPUSHBUTTON) {
                // Push buttons are drawn by themeDrawDialogButton; remember which is the default.
                SetWindowLongPtrW(c, GWLP_USERDATA, kind == BS_DEFPUSHBUTTON ? 1 : 0);
                SetWindowLongW(c, GWL_STYLE, (style & ~BS_TYPEMASK) | BS_OWNERDRAW);
            }
        } else if (lstrcmpiW(cls, L"ComboBox") == 0 || lstrcmpiW(cls, L"Edit") == 0) {
            SetWindowTheme(c, L"", L"");
        } else if (lstrcmpiW(cls, L"SysListView32") == 0) {
            SetWindowTheme(c, p.dark ? L"DarkMode_Explorer" : L"Explorer", nullptr);
            SendMessageW(c, LVM_SETBKCOLOR, 0, (LPARAM)p.bg);
            SendMessageW(c, LVM_SETTEXTBKCOLOR, 0, (LPARAM)p.bg);
            SendMessageW(c, LVM_SETTEXTCOLOR, 0, (LPARAM)p.fg);
            HWND hdr = (HWND)SendMessageW(c, LVM_GETHEADER, 0, 0);
            if (hdr) SetWindowTheme(hdr, p.dark ? L"DarkMode_ItemsView" : nullptr, nullptr);
        } else {
            SetWindowTheme(c, p.dark ? L"DarkMode_Explorer" : nullptr, nullptr);
        }
        SendMessageW(c, WM_THEMECHANGED, 0, 0);
    }
    InvalidateRect(dlg, nullptr, TRUE);
}

INT_PTR themeCtlColor(UINT msg, WPARAM wp, const Palette& p) {
    HDC dc = (HDC)wp;
    switch (msg) {
    case WM_CTLCOLORDLG:
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLORBTN:
        SetTextColor(dc, p.faceText);
        SetBkColor(dc, p.face);
        return (INT_PTR)brushFor(g_faceBrush, g_faceCol, p.face);
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORLISTBOX:
        SetTextColor(dc, p.fg);
        SetBkColor(dc, p.bg);
        return (INT_PTR)brushFor(g_bgBrush, g_bgCol, p.bg);
    }
    return 0;
}

bool themeMenuBarMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, LRESULT* result, const Palette& p) {
    switch (msg) {
    case WM_UAHDRAWMENU: {
        UAHMENU* m = (UAHMENU*)lp;
        RECT r;
        if (menuBarRect(hwnd, r)) FillRect(m->hdc, &r, brushFor(g_faceBrush, g_faceCol, p.face));
        *result = 0;
        return true;
    }
    case WM_UAHDRAWMENUITEM: {
        UAHDRAWMENUITEM* d = (UAHDRAWMENUITEM*)lp;
        wchar_t text[256] = {};
        MENUITEMINFOW mii = {sizeof(mii)};
        mii.fMask = MIIM_STRING;
        mii.dwTypeData = text;
        mii.cch = 255;
        GetMenuItemInfoW(d->um.hmenu, d->umi.iPosition, TRUE, &mii);
        const UINT st = d->dis.itemState;
        const bool hot = (st & (ODS_HOTLIGHT | ODS_SELECTED)) != 0;
        FillRect(d->um.hdc, &d->dis.rcItem,
                 hot ? brushFor(g_hotBrush, g_hotCol, p.hot) : brushFor(g_faceBrush, g_faceCol, p.face));
        SetBkMode(d->um.hdc, TRANSPARENT);
        SetTextColor(d->um.hdc, (st & (ODS_GRAYED | ODS_DISABLED)) ? RGB(128, 128, 128) : p.faceText);
        DrawTextW(d->um.hdc, text, -1, &d->dis.rcItem,
                  DT_CENTER | DT_SINGLELINE | DT_VCENTER | ((st & ODS_NOACCEL) ? DT_HIDEPREFIX : 0));
        *result = 0;
        return true;
    }
    case WM_NCPAINT:
    case WM_NCACTIVATE:
        *result = DefWindowProcW(hwnd, msg, wp, lp);
        paintMenuBarLine(hwnd, p);
        return true;
    }
    return false;
}

namespace {
ULONG_PTR g_gdiplusToken = 0;
}

void gdiplusStart() {
    Gdiplus::GdiplusStartupInput in;
    Gdiplus::GdiplusStartup(&g_gdiplusToken, &in, nullptr);
}

void gdiplusStop() {
    if (g_gdiplusToken) Gdiplus::GdiplusShutdown(g_gdiplusToken);
    g_gdiplusToken = 0;
}

void drawDiffBand(HDC dc, int W, float a0, float a1, float b0, float b1, COLORREF fill, COLORREF edge, float penW) {
    using namespace Gdiplus;
    if (!g_gdiplusToken || W < 2) return;
    Graphics gr(dc);
    gr.SetSmoothingMode(SmoothingModeAntiAlias);
    gr.SetPixelOffsetMode(PixelOffsetModeHalf);
    // Run a little past both sides so the curves meet the panes' lines cleanly.
    const float x0 = -1.0f, x1 = (float)W + 1.0f, mid = ((float)W) / 2.0f;
    if (a1 < a0 + penW) a1 = a0 + penW;
    if (b1 < b0 + penW) b1 = b0 + penW;
    GraphicsPath path;
    path.AddBezier(x0, a0, mid, a0, mid, b0, x1, b0);
    path.AddLine(x1, b0, x1, b1);
    path.AddBezier(x1, b1, mid, b1, mid, a1, x0, a1);
    path.CloseFigure();
    SolidBrush br(Color(255, GetRValue(fill), GetGValue(fill), GetBValue(fill)));
    gr.FillPath(&br, &path);
    // Borders sit just inside the band, exactly where the panes draw theirs.
    Pen pen(Color(255, GetRValue(edge), GetGValue(edge), GetBValue(edge)), penW);
    const float h = penW / 2.0f;
    gr.DrawBezier(&pen, x0, a0 + h, mid, a0 + h, mid, b0 + h, x1, b0 + h);
    gr.DrawBezier(&pen, x0, a1 - h, mid, a1 - h, mid, b1 - h, x1, b1 - h);
}

// ---------------------------------------------------------------- 3D buttons

namespace {

Gdiplus::Color gp(COLORREF c, BYTE a = 255) { return Gdiplus::Color(a, GetRValue(c), GetGValue(c), GetBValue(c)); }

COLORREF mixc(COLORREF a, COLORREF b, float t) {  // t of a
    return RGB((int)(GetRValue(a) * t + GetRValue(b) * (1 - t)), (int)(GetGValue(a) * t + GetGValue(b) * (1 - t)),
               (int)(GetBValue(a) * t + GetBValue(b) * (1 - t)));
}


}  // namespace

// Late-90s bevel, square and pixel-exact, in the theme's colours: two-pixel
// edges lit from the top left (light outer, softer inner), inverted when the
// button is in. A toggle that is on gets the classic dithered face, tinted.
void drawButton3D(HDC dc, const RECT& rc, ButtonState st, const Palette& p, COLORREF tint) {
    auto fill = [&](int l, int t, int r, int b, COLORREF c) {
        if (r <= l || b <= t) return;
        RECT x = {l, t, r, b};
        SetBkColor(dc, c);
        ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &x, L"", 0, nullptr);
    };
    const COLORREF white = RGB(255, 255, 255), black = RGB(0, 0, 0);
    const bool in = st == BTN_PRESSED || st == BTN_LATCHED || st == BTN_LATCHED_HOT;
    const bool hot = st == BTN_HOT || st == BTN_LATCHED_HOT;
    const bool latched = st == BTN_LATCHED || st == BTN_LATCHED_HOT;

    // The face: a step up from the panel, a touch brighter under the mouse.
    COLORREF face = p.dark ? mixc(white, p.face, 0.09f) : mixc(white, p.face, 0.35f);
    if (hot) face = p.dark ? mixc(white, face, 0.07f) : mixc(white, face, 0.40f);
    // Edge colours: highlight, light, shadow, dark (outer to inner, both sides).
    const COLORREF hi = p.dark ? mixc(white, face, 0.30f) : white;
    const COLORREF light = p.dark ? mixc(white, face, 0.12f) : mixc(white, face, 0.55f);
    const COLORREF shadow = p.dark ? mixc(black, face, 0.35f) : mixc(black, face, 0.30f);
    const COLORREF dark = p.dark ? mixc(black, face, 0.75f) : mixc(black, face, 0.62f);

    const int L = rc.left, T = rc.top, R = rc.right, B = rc.bottom;
    const int bw = 1;  // one bevel pixel per ring: crisp at any size
    if (latched) {
        // Dithered face: every other pixel the tint, the 90s "pressed toggle".
        COLORREF a = mixc(tint, face, p.dark ? 0.40f : 0.30f), b2 = p.dark ? mixc(black, face, 0.15f) : mixc(white, face, 0.6f);
        if (hot) a = mixc(white, a, p.dark ? 0.08f : 0.15f);
        fill(L, T, R, B, b2);
        for (int y = T + 2; y < B - 2; y++)
            for (int x = L + 2 + ((y - T) & 1); x < R - 2; x += 2) SetPixel(dc, x, y, a);
    } else {
        // Modern touch: the face gets a faint top-to-bottom shade.
        const int mid = T + (B - T) / 2;
        fill(L, T, R, mid, in ? mixc(black, face, 0.05f) : mixc(white, face, p.dark ? 0.03f : 0.25f));
        fill(L, mid, R, B, in ? face : mixc(black, face, p.dark ? 0.05f : 0.03f));
    }
    // Outer ring.
    fill(L, T, R - bw, T + bw, in ? dark : hi);          // top
    fill(L, T, L + bw, B - bw, in ? dark : hi);          // left
    fill(L, B - bw, R, B, in ? hi : dark);               // bottom
    fill(R - bw, T, R, B, in ? hi : dark);               // right
    // Inner ring.
    fill(L + bw, T + bw, R - 2 * bw, T + 2 * bw, in ? shadow : light);
    fill(L + bw, T + bw, L + 2 * bw, B - 2 * bw, in ? shadow : light);
    fill(L + bw, B - 2 * bw, R - bw, B - bw, in ? light : shadow);
    fill(R - 2 * bw, T + bw, R - bw, B - bw, in ? light : shadow);
}

void drawGlyph(HDC dc, const RECT& rc, Glyph gl, COLORREF c1, COLORREF c2, bool pushed) {
    using namespace Gdiplus;
    if (!g_gdiplusToken) return;
    Graphics gr(dc);
    gr.SetSmoothingMode(SmoothingModeAntiAlias);
    gr.SetPixelOffsetMode(PixelOffsetModeHalf);
    const float w = (float)(rc.right - rc.left), h = (float)(rc.bottom - rc.top);
    const float cx = rc.left + w / 2.0f + (pushed ? 1.0f : 0.0f), cy = rc.top + h / 2.0f + (pushed ? 1.0f : 0.0f);
    const float s = (w < h ? w : h) * 0.40f;  // half the glyph's size
    SolidBrush b1(gp(c1)), b2(gp(c2));
    // Triangles: the bounding box is centred, then nudged a little toward the
    // point, because the eye puts a triangle's middle nearer its flat side.
    const float nudge = s * 0.18f;
    switch (gl) {
    case GLYPH_LEFT: {
        PointF pt[3] = {{cx + s * 0.85f - nudge, cy - s}, {cx + s * 0.85f - nudge, cy + s}, {cx - s * 0.85f - nudge, cy}};
        gr.FillPolygon(&b1, pt, 3);
        break;
    }
    case GLYPH_RIGHT: {
        PointF pt[3] = {{cx - s * 0.85f + nudge, cy - s}, {cx - s * 0.85f + nudge, cy + s}, {cx + s * 0.85f + nudge, cy}};
        gr.FillPolygon(&b1, pt, 3);
        break;
    }
    case GLYPH_BOTH: {
        PointF l[3] = {{cx, cy - s}, {cx, cy + s}, {cx - s, cy}};
        PointF r[3] = {{cx, cy - s}, {cx + s, cy}, {cx, cy + s}};
        gr.FillPolygon(&b1, l, 3);
        gr.FillPolygon(&b2, r, 3);
        break;
    }
    case GLYPH_UP: {
        PointF pt[3] = {{cx - s, cy + s * 0.6f - nudge}, {cx + s, cy + s * 0.6f - nudge}, {cx, cy - s * 0.6f - nudge}};
        gr.FillPolygon(&b1, pt, 3);
        break;
    }
    case GLYPH_DOWN: {
        PointF pt[3] = {{cx - s, cy - s * 0.6f + nudge}, {cx + s, cy - s * 0.6f + nudge}, {cx, cy + s * 0.6f + nudge}};
        gr.FillPolygon(&b1, pt, 3);
        break;
    }
    case GLYPH_CLOSE: {
        Pen pen(gp(c1), (w < 16 ? 1.3f : w / 11.0f));
        pen.SetStartCap(LineCapRound);
        pen.SetEndCap(LineCapRound);
        float k = s * 0.75f;
        gr.DrawLine(&pen, cx - k, cy - k, cx + k, cy + k);
        gr.DrawLine(&pen, cx + k, cy - k, cx - k, cy + k);
        break;
    }
    }
}

// ---------------------------------------------------------------- menus

namespace {

struct MenuItem {
    std::wstring text;  // "label\tshortcut"
    bool radio = false, sub = false, separator = false, bar = false;
};

HBRUSH g_menuBrush = nullptr;
COLORREF g_menuBrushCol = CLR_INVALID;

void ownerDrawPopup(HMENU m) {
    int n = GetMenuItemCount(m);
    for (int i = 0; i < n; i++) {
        wchar_t text[256] = {};
        MENUITEMINFOW mii = {sizeof(mii)};
        mii.fMask = MIIM_FTYPE | MIIM_STRING | MIIM_SUBMENU | MIIM_DATA;
        mii.dwTypeData = text;
        mii.cch = 255;
        if (!GetMenuItemInfoW(m, i, TRUE, &mii)) continue;
        if (mii.fType & MFT_OWNERDRAW) continue;
        MenuItem* it = new MenuItem();  // lives as long as the menu (the whole run)
        it->text = text;
        it->separator = (mii.fType & MFT_SEPARATOR) != 0;
        it->radio = (mii.fType & MFT_RADIOCHECK) != 0;
        it->sub = mii.hSubMenu != nullptr;
        MENUITEMINFOW set = {sizeof(set)};
        set.fMask = MIIM_FTYPE | MIIM_DATA;
        set.fType = (mii.fType & ~MFT_STRING) | MFT_OWNERDRAW;
        set.dwItemData = (ULONG_PTR)it;
        SetMenuItemInfoW(m, i, TRUE, &set);
        if (mii.hSubMenu) ownerDrawPopup(mii.hSubMenu);
    }
}

MenuItem* itemFor(HMENU bar, UINT id) {
    MENUITEMINFOW mii = {sizeof(mii)};
    mii.fMask = MIIM_DATA | MIIM_FTYPE;
    if (!GetMenuItemInfoW(bar, id, FALSE, &mii) || !(mii.fType & MFT_OWNERDRAW)) return nullptr;
    return (MenuItem*)mii.dwItemData;
}

void applyMenuBrush(HMENU m, HBRUSH b) {
    MENUINFO mi = {sizeof(mi)};
    mi.fMask = MIM_BACKGROUND;
    mi.hbrBack = b;
    SetMenuInfo(m, &mi);
    int n = GetMenuItemCount(m);
    for (int i = 0; i < n; i++)
        if (HMENU s = GetSubMenu(m, i)) applyMenuBrush(s, b);
}

}  // namespace

void themeOwnerDrawMenu(HMENU bar) {
    int n = GetMenuItemCount(bar);
    for (int i = 0; i < n; i++)
        if (HMENU s = GetSubMenu(bar, i)) ownerDrawPopup(s);  // the bar itself is drawn by the UAH messages
}

void themeSetMenuText(HMENU bar, UINT id, const wchar_t* text) {
    if (MenuItem* it = itemFor(bar, id)) it->text = text;
}

void themeSetMenuRadio(HMENU bar, UINT id) {
    if (MenuItem* it = itemFor(bar, id)) it->radio = true;
}

void themeMenuColors(HMENU bar, const Palette& p) {
    HBRUSH b = brushFor(g_menuBrush, g_menuBrushCol, p.face);
    int n = GetMenuItemCount(bar);
    for (int i = 0; i < n; i++)
        if (HMENU s = GetSubMenu(bar, i)) applyMenuBrush(s, b);
}

void themeMeasureMenuItem(HWND hwnd, MEASUREITEMSTRUCT* mi, HFONT font) {
    if (mi->CtlType != ODT_MENU) return;
    const MenuItem* it = (const MenuItem*)mi->itemData;
    const UINT dpi = GetDpiForWindow(hwnd);
    if (!it || it->separator) {
        mi->itemWidth = 0;
        mi->itemHeight = MulDiv(7, dpi, 96);
        return;
    }
    HDC dc = GetDC(hwnd);
    HGDIOBJ of = SelectObject(dc, font);
    std::wstring label = it->text, keys;
    size_t tab = label.find(L'\t');
    if (tab != std::wstring::npos) {
        keys = label.substr(tab + 1);
        label.resize(tab);
    }
    RECT a = {}, b = {};
    DrawTextW(dc, label.c_str(), -1, &a, DT_CALCRECT | DT_SINGLELINE);
    if (!keys.empty()) DrawTextW(dc, keys.c_str(), -1, &b, DT_CALCRECT | DT_SINGLELINE);
    TEXTMETRICW tm;
    GetTextMetricsW(dc, &tm);
    SelectObject(dc, of);
    ReleaseDC(hwnd, dc);
    // Padding: room for a check mark on the left, the shortcut on the right.
    const int check = MulDiv(28, dpi, 96), padR = MulDiv(24, dpi, 96), gap = MulDiv(32, dpi, 96);
    mi->itemWidth = check + (a.right - a.left) + (keys.empty() ? 0 : gap + (b.right - b.left)) + padR;
    mi->itemHeight = tm.tmHeight + MulDiv(10, dpi, 96);
}

void themeDrawMenuItem(DRAWITEMSTRUCT* di, HFONT font, const Palette& p) {
    if (di->CtlType != ODT_MENU) return;
    const MenuItem* it = (const MenuItem*)di->itemData;
    HDC dc = di->hDC;
    RECT r = di->rcItem;  // the whole item: the highlight fills it edge to edge
    const UINT dpi = GetDpiForWindow(WindowFromDC(dc) ? WindowFromDC(dc) : GetDesktopWindow());
    auto fill = [&](const RECT& x, COLORREF c) {
        SetBkColor(dc, c);
        ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &x, L"", 0, nullptr);
    };
    if (!it || it->separator) {
        fill(r, p.face);
        // 90s etched line: a dark line over a light one.
        int y = (r.top + r.bottom) / 2 - 1, pad = MulDiv(6, dpi, 96);
        RECT a = {r.left + pad, y, r.right - pad, y + 1}, b = {r.left + pad, y + 1, r.right - pad, y + 2};
        fill(a, mixc(RGB(0, 0, 0), p.face, p.dark ? 0.45f : 0.20f));
        fill(b, mixc(RGB(255, 255, 255), p.face, p.dark ? 0.08f : 0.70f));
        return;
    }
    const bool sel = (di->itemState & ODS_SELECTED) != 0;
    const bool off = (di->itemState & (ODS_GRAYED | ODS_DISABLED)) != 0;
    const bool checked = (di->itemState & ODS_CHECKED) != 0;
    const COLORREF bg = sel && !off ? p.accent : p.face;
    const COLORREF fg = off ? mixc(p.faceText, p.face, 0.40f) : sel ? RGB(255, 255, 255) : p.faceText;
    fill(r, bg);

    const int check = MulDiv(28, dpi, 96), padR = MulDiv(24, dpi, 96);
    if (checked) {
        // A tick, or a dot for one-of choices; square ends, centred in the check column.
        RECT cr = {r.left, r.top, r.left + check, r.bottom};
        if (g_gdiplusToken) {
            using namespace Gdiplus;
            Graphics gr(dc);
            gr.SetSmoothingMode(SmoothingModeAntiAlias);
            const float cx = (cr.left + cr.right) / 2.0f, cy = (cr.top + cr.bottom) / 2.0f, s = MulDiv(4, dpi, 96);
            if (it->radio) {
                SolidBrush b(gp(fg));
                gr.FillRectangle(&b, cx - s * 0.6f, cy - s * 0.6f, s * 1.2f, s * 1.2f);  // square dot
            } else {
                Pen pen(gp(fg), MulDiv(2, dpi, 96) * 0.9f);
                PointF pts[3] = {{cx - s, cy}, {cx - s * 0.3f, cy + s * 0.7f}, {cx + s, cy - s * 0.8f}};
                gr.DrawLines(&pen, pts, 3);
            }
        }
    }
    std::wstring label = it->text, keys;
    size_t tab = label.find(L'\t');
    if (tab != std::wstring::npos) {
        keys = label.substr(tab + 1);
        label.resize(tab);
    }
    HGDIOBJ of = SelectObject(dc, font);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, fg);
    const UINT pre = (di->itemState & ODS_NOACCEL) ? DT_HIDEPREFIX : 0;
    RECT t = {r.left + check, r.top, r.right - padR, r.bottom};
    DrawTextW(dc, label.c_str(), -1, &t, DT_LEFT | DT_VCENTER | DT_SINGLELINE | pre);
    if (!keys.empty()) DrawTextW(dc, keys.c_str(), -1, &t, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    SelectObject(dc, of);
    if (it->sub && g_gdiplusToken) {
        using namespace Gdiplus;
        Graphics gr(dc);
        gr.SetSmoothingMode(SmoothingModeAntiAlias);
        const float cx = r.right - padR / 2.0f, cy = (r.top + r.bottom) / 2.0f, s = MulDiv(3, dpi, 96);
        SolidBrush b(gp(fg));
        PointF pts[3] = {{cx - s * 0.6f, cy - s * 1.2f}, {cx - s * 0.6f, cy + s * 1.2f}, {cx + s * 0.8f, cy}};
        gr.FillPolygon(&b, pts, 3);
    }
    // Stop Windows from drawing its own submenu arrow over ours.
    ExcludeClipRect(dc, r.left, r.top, r.right, r.bottom);
}

// ---------------------------------------------------------------- dialog buttons

bool themeDrawDialogButton(DRAWITEMSTRUCT* di, const Palette& p) {
    if (di->CtlType != ODT_BUTTON) return false;
    HDC dc = di->hDC;
    RECT r = di->rcItem;
    const bool pressed = (di->itemState & ODS_SELECTED) != 0;
    const bool off = (di->itemState & ODS_DISABLED) != 0;
    const bool def = GetWindowLongPtrW(di->hwndItem, GWLP_USERDATA) == 1;
    // Background behind the button (the dialog face).
    SetBkColor(dc, p.face);
    ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &r, L"", 0, nullptr);
    if (def) {
        // The default button gets the classic extra dark frame.
        RECT f = r;
        HBRUSH b = CreateSolidBrush(mixc(RGB(0, 0, 0), p.face, p.dark ? 0.85f : 0.75f));
        FrameRect(dc, &f, b);
        DeleteObject(b);
        InflateRect(&r, -1, -1);
    }
    drawButton3D(dc, r, pressed ? BTN_PRESSED : BTN_NORMAL, p, 0);
    wchar_t text[128] = {};
    GetWindowTextW(di->hwndItem, text, 128);
    HFONT font = (HFONT)SendMessageW(di->hwndItem, WM_GETFONT, 0, 0);
    HGDIOBJ of = SelectObject(dc, font);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, off ? mixc(p.faceText, p.face, 0.40f) : p.faceText);
    RECT t = r;
    if (pressed) OffsetRect(&t, 1, 1);
    const UINT pre = (di->itemState & ODS_NOACCEL) ? DT_HIDEPREFIX : 0;
    DrawTextW(dc, text, -1, &t, DT_CENTER | DT_VCENTER | DT_SINGLELINE | pre);
    SelectObject(dc, of);
    if (di->itemState & ODS_FOCUS) {
        // 90s dotted focus rectangle, inside the bevel.
        RECT fr = r;
        InflateRect(&fr, -4, -4);
        SetTextColor(dc, p.faceText);
        SetBkColor(dc, p.face);
        DrawFocusRect(dc, &fr);
    }
    return true;
}

LRESULT themeMenuChar(HMENU menu, wchar_t ch) {
    wchar_t want = (wchar_t)(UINT_PTR)CharLowerW((LPWSTR)(UINT_PTR)ch);
    int n = GetMenuItemCount(menu);
    for (int i = 0; i < n; i++) {
        MENUITEMINFOW mii = {sizeof(mii)};
        mii.fMask = MIIM_FTYPE | MIIM_DATA | MIIM_STATE;
        if (!GetMenuItemInfoW(menu, i, TRUE, &mii) || !(mii.fType & MFT_OWNERDRAW) || !mii.dwItemData) continue;
        const std::wstring& t = ((const MenuItem*)mii.dwItemData)->text;
        size_t a = t.find(L'&');
        if (a == std::wstring::npos || a + 1 >= t.size()) continue;
        wchar_t k = (wchar_t)(UINT_PTR)CharLowerW((LPWSTR)(UINT_PTR)t[a + 1]);
        if (k == want) return MAKELRESULT(i, MNC_EXECUTE);
    }
    return MAKELRESULT(0, MNC_IGNORE);
}
