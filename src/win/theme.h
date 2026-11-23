#pragma once
// Light / dark theme for everything JustNotepad draws, plus the parts Windows
// draws for us: title bar, menu bar and menus, scroll bars and our dialogs.
//
// Windows has no public switch for dark menus or dark dialog controls. This
// uses the same uxtheme.dll entry points Explorer and Notepad use (looked up
// by ordinal, so on systems without them the app simply stays light there).
#include <windows.h>

struct Palette {
    bool dark = false;
    COLORREF bg, fg;            // text area
    COLORREF face, faceText;    // menu bar, status bar, headers, dialogs
    COLORREF border;            // separators
    COLORREF hot;               // hovered menu item
    COLORREF accent;            // focus, hover rings
    COLORREF selection;         // selected text background
};

enum class ThemeMode { System = 0, Light = 1, Dark = 2 };

Palette makePalette(bool dark);
bool systemUsesDarkMode();

// Call once at startup, before any window exists.
void themeInit();
// Switches menus (and the app as a whole) between light and dark.
void themeSetAppDark(bool dark);
// Title bar, scroll bars, and the frame's menu bar for a top-level window.
void themeWindow(HWND hwnd, bool dark);
// Child controls of a dialog; call from WM_INITDIALOG.
void themeDialog(HWND dlg, const Palette& p);

// Paints the menu bar in the palette's colours. Call first in the frame's
// window procedure; returns true when the message was handled.
bool themeMenuBarMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, LRESULT* result, const Palette& p);

// GDI+ (anti-aliased drawing). Start once before drawing, stop at exit.
void gdiplusStart();
void gdiplusStop();

// The diff band joining rows [a0, a1) on the left to [b0, b1) on the right:
// filled between two S-curves, with a smooth border `penW` pixels wide that
// matches the border lines drawn along the changes in the panes.
void drawDiffBand(HDC dc, int W, float a0, float a1, float b0, float b1, COLORREF fill, COLORREF edge, float penW);

// ---- 3D buttons (drawn with GDI+, anti-aliased)
enum ButtonState { BTN_NORMAL, BTN_HOT, BTN_PRESSED, BTN_LATCHED, BTN_LATCHED_HOT };
enum Glyph { GLYPH_LEFT, GLYPH_BOTH, GLYPH_RIGHT, GLYPH_UP, GLYPH_DOWN, GLYPH_CLOSE };

// A rounded, raised button: lit from above, with a soft shadow. Pressed and
// latched (a toggle that is on) look pushed in; `tint` colours a latched face.
void drawButton3D(HDC dc, const RECT& r, ButtonState st, const Palette& p, COLORREF tint);
// A symbol centred in `r` (optically: triangles sit on their visual centre).
// `c1` / `c2` are its colours (GLYPH_BOTH uses both halves). `pushed` moves it
// down a pixel, so it goes in with the button.
void drawGlyph(HDC dc, const RECT& r, Glyph gl, COLORREF c1, COLORREF c2, bool pushed);

// ---- menus: drawn by us, square, full-width highlight, padded text
// Turns every item below the menu bar into an owner-drawn one (keeps its text).
void themeOwnerDrawMenu(HMENU bar);
void themeSetMenuText(HMENU bar, UINT id, const wchar_t* text);
// The same for a context menu (a popup on its own) and its submenus.
void themeOwnerDrawPopup(HMENU popup);
void themePopupColors(HMENU popup, const Palette& p);
// Frees what themeOwnerDrawPopup attached to a popup; call before DestroyMenu.
void themeReleasePopup(HMENU popup);
void themeSetMenuRadio(HMENU bar, UINT id);  // checked shows a dot instead of a tick
void themeMenuColors(HMENU bar, const Palette& p);
void themeMeasureMenuItem(HWND hwnd, MEASUREITEMSTRUCT* mi, HFONT font);
void themeDrawMenuItem(DRAWITEMSTRUCT* di, HFONT font, const Palette& p);
// WM_MENUCHAR: owner-drawn items still open with their underlined letter.
LRESULT themeMenuChar(HMENU menu, wchar_t ch);

// ---- dialogs: push buttons drawn with drawButton3D (themeDialog makes them owner-drawn)
bool themeDrawDialogButton(DRAWITEMSTRUCT* di, const Palette& p);

// WM_CTLCOLOR* for dialogs. Returns the brush to use, or 0 to let Windows decide.
INT_PTR themeCtlColor(UINT msg, WPARAM wp, const Palette& p);
