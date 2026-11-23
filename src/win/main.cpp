// JustNotepad - Windows front end.
//
// Two windows: a frame (menu, status strip) and one child view that draws the
// text itself. There is no edit control: the view asks the core for the rows
// that are on screen and paints them, so the cost of every paint and every
// keystroke depends on the window size, never on the file size.
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <windowsx.h>
#include <commdlg.h>
#include <shellapi.h>
#include <commctrl.h>
#include <uxtheme.h>

#include <stdlib.h>
#include <string.h>
#include <wchar.h>

#include <algorithm>
#include <cmath>
#include <ctype.h>
#include <memory>
#include <string>
#include <thread>
#include <vector>

#include "../core/diff.h"
#include "../core/highlight.h"
#include "../core/doc.h"
#include "../core/layout.h"
#include "../core/lineindex.h"
#include "../core/utf.h"
#include "resource.h"
#include "theme.h"
#include "updater.h"
#include "vcs.h"
#include "version.h"

using namespace jnp;

namespace {

const wchar_t* const kAppName = L"JustNotepad";
const wchar_t* const kFrameClass = L"JustNotepadFrame";
const wchar_t* const kViewClass = L"JustNotepadView";
const wchar_t* const kDiffClass = L"JustNotepadDiff";
const wchar_t* const kStripClass = L"JustNotepadDiffStrip";

#define JN_WIDEN2(x) L##x
#define JN_WIDEN(x) JN_WIDEN2(x)

// Files up to this size are read into memory, so other programs can keep
// writing to them. Larger files are memory-mapped and open instantly.
const u64 kHeapMax = 32ull << 20;
// Files up to this size get their line index before the first paint (a few
// milliseconds). Larger ones are indexed on a background thread.
const u64 kSyncIndexMax = 8ull << 20;
const u64 kMaxClipboard = 256ull << 20;

// Change markers and the diff view are computed for files up to this size.
const u64 kDiffMax = 16ull << 20;
// Below this size the diff is refreshed on every edit; above, shortly after typing stops.
const u64 kDiffInstant = 256u << 10;

const UINT WM_APP_INDEX = WM_APP + 1, WM_APP_VCS = WM_APP + 2, WM_APP_P4ASK = WM_APP + 3,
           WM_APP_P4EDITED = WM_APP + 4;
const UINT_PTR kTimerBlink = 1, kTimerIndex = 2, kTimerDrag = 3, kTimerDiff = 4, kTimerSession = 5;

const COLORREF kAdded = RGB(40, 160, 70), kChanged = RGB(40, 110, 220), kRemoved = RGB(210, 60, 60);
// The two sides of the diff, as in P4Merge: the depot / committed version on
// the left, your file on the right. Used for headers and the choice buttons.
const COLORREF kDepotCol = RGB(52, 120, 206), kYoursCol = RGB(64, 160, 72);
enum HunkChoice { kUseDepot = 0, kUseBoth = 1, kKeepYours = 2 };

enum FileEnc { FE_UTF8, FE_UTF8BOM, FE_UTF16LE, FE_UTF16BE, FE_ANSI };

struct Backing {
    HANDLE file = INVALID_HANDLE_VALUE;
    HANDLE map = nullptr;
    const u8* view = nullptr;
    std::vector<u8> heap;
    bool mapped = false;

    void close() {
        if (view) UnmapViewOfFile(view);
        if (map) CloseHandle(map);
        if (file != INVALID_HANDLE_VALUE) CloseHandle(file);
        view = nullptr;
        map = nullptr;
        file = INVALID_HANDLE_VALUE;
        mapped = false;
        std::vector<u8>().swap(heap);
    }
    void take(Backing& o) {
        close();
        file = o.file;
        map = o.map;
        view = o.view;
        mapped = o.mapped;
        heap = std::move(o.heap);
        o.file = INVALID_HANDLE_VALUE;
        o.map = nullptr;
        o.view = nullptr;
        o.mapped = false;
    }
};

// Everything that belongs to one open file (one tab).
struct DocState {
    Doc doc;
    std::unique_ptr<LineIndex> index;
    Backing back;
    u64 bomSkip = 0;
    unsigned indexGen = 0;
    bool indexing = false;
    std::wstring path;
    FileEnc enc = FE_UTF8;
    bool crlf = true;

    // where the view was
    u64 top = 0, caret = 0, anchor = 0;
    int wantCol = -1;
    int hcol = 0;
    int topFill = 0;  // filler rows shown above the top line of the editor

    // version control and diff
    VcsInfo vcs;
    unsigned vcsGen = 0;
    bool vcsPending = false, p4Asked = false, checkingOut = false;
    std::vector<u8> baseText;   // base revision, in the same form as the document
    LineIndex baseIdx;
    Doc baseDoc;
    std::vector<u64> baseHash;  // one hash per base line
    std::vector<u64> curHash;   // one hash per current line
    std::vector<Hunk> hunks;    // base -> current
    u64 diffDocVersion = NPOS;
    // Choices made with the three buttons, kept until the base changes
    // (submit, revert, sync). Keyed by the base lines they cover, which do
    // not move while you edit.
    struct Choice { u32 aStart, aCount; std::string yours; int choice; };
    std::vector<Choice> choices;
    // Aligned diff: blank filler rows on the shorter side of each change so
    // both sides keep the same number of rows and scroll together.
    struct Fill { u64 line, count, cum; };  // `count` rows before `line`; cum = running total
    std::vector<Fill> fillA, fillB;

    // Syntax highlighting: the language, and the tokenizer state at the start
    // of every kHlStep-th line (for this text and for the base revision).
    struct HlCache {
        std::vector<u32> states;
        u64 version = NPOS;
    };
    const Lang* lang = nullptr;
    HlCache hl, hlBase;

    // Session backup of unsaved text (see saveSession).
    std::wstring backupName;
    u64 backupVersion = NPOS;
};

struct ClosedTab {
    std::wstring path;
    u64 caret, top;
};

struct App {
    std::vector<ClosedTab> closed;  // for Ctrl+Shift+T, most recent last
    // open files; `d` is the one shown
    std::vector<std::unique_ptr<DocState>> tabs;
    DocState* d = nullptr;
    int tabHover = -1, tabCloseHover = -1;
    std::vector<RECT> tabRects, tabCloseRects;

    HINSTANCE inst = nullptr;
    HWND frame = nullptr, view = nullptr;
    HMENU menu = nullptr;
    HACCEL accel = nullptr;


    // view state
    bool wrap = false, lineNumbers = true;
    std::vector<Row> rows;
    int cliW = 0, cliH = 0, visRows = 1, visCols = 1, gutterPx = 0, vshift = 0;
    bool caretOn = true, focused = false, selecting = false;
    int wheelAcc = 0;
    wchar_t pendingHigh = 0;
    bool inScrollUpdate = false;

    // drawing
    HFONT font = nullptr, uiFont = nullptr;
    std::wstring fontFace = L"Consolas";
    int fontPt = 11;
    int dpi = 96, cw = 8, ch = 16, statusH = 22;
    HDC memDC = nullptr;
    HBITMAP memBmp = nullptr;
    HGDIOBJ memOld = nullptr;
    int memW = 0, memH = 0;

    // single-byte ("ANSI") files
    UINT ansiCp = 1252;
    u16 ansi[256] = {};
    bool ansiReady = false;

    // find / replace
    HWND findDlg = nullptr;
    FINDREPLACEW fr = {};
    wchar_t findBuf[256] = {}, replBuf[256] = {};
    UINT findMsg = 0;
    bool findCase = false, findWord = false;

    // version control and diff
    HWND diffPane = nullptr, strip = nullptr;
    std::vector<u8> rowFill;         // parallel to rows: 1 = filler row
    Palette pal = makePalette(false);
    int themeMode = 0;               // 0 system, 1 light, 2 dark
    bool highlight = true;           // View > Syntax Highlighting
    struct IconHit { RECT rc; int hunk, choice; };
    std::vector<IconHit> icons;      // choice buttons as last painted in the strip
    int hoverHunk = -1, hoverChoice = -1;
    int pressHunk = -1, pressChoice = -1;  // choice button held down
    int navHover = 0, navPress = 0;        // header arrows: 1 previous, 2 next
    int tabClosePress = -1;                // tab whose close button is held down
    std::wstring hint;               // shown in the status bar while hovering a button
    RECT navPrev = {}, navNext = {}; // previous / next change arrows in the header
    bool diffOn = false, wrapBeforeDiff = false;
    int updatePct = -1;        // update download progress, -1 when none

    std::wstring statusLeft, statusRight, title;
    u64 gotoLine = 0;
};

App g;
unsigned g_gen = 0;  // source of unique ids for background work replies

INT_PTR dialogTheme(HWND dlg, UINT msg, WPARAM wp, LPARAM lp);
void scrollRows(int n);
void addTab();
bool isPristine(const DocState* d);
void paintTabs(HDC dc);
bool openInTab(const std::wstring& path);
int tabIndex(const DocState* d);
void activateTab(int i);
void fillRect(HDC dc, int l, int t, int r, int b, COLORREF c);
int edgeW();
COLORREF inlineOf(COLORREF c);
bool inlineSpan(const Doc& doc, u64 s, u64 e, const Doc& other, u64 otherLine, u64& x0, u64& x1);
int headerH();
int tabBarH();
void invalidateTabs();
DocState* tabByVcsGen(unsigned gen);
COLORREF changeColor(u32 aCount, u32 bCount);
COLORREF tintOf(COLORREF c);
COLORREF fillTintOf(COLORREF c);
COLORREF edgeOf(COLORREF c);
COLORREF fillerColorB(u64 line);
void cmdChangelistFiles();
bool isWordByte(u8 b);
bool findNext(bool down);

// ====================================================================== helpers

u64 selMin() { return std::min(g.d->caret, g.d->anchor); }
u64 selMax() { return std::max(g.d->caret, g.d->anchor); }
bool hasSel() { return g.d->caret != g.d->anchor; }

void ensureAnsiTable() {
    if (g.ansiReady) return;
    UINT cp = GetACP();
    CPINFO info;
    if (cp == CP_UTF8 || !GetCPInfo(cp, &info) || info.MaxCharSize != 1) cp = 1252;
    g.ansiCp = cp;
    for (int i = 0; i < 256; i++) {
        char c = (char)i;
        wchar_t w = 0;
        if (MultiByteToWideChar(cp, 0, &c, 1, &w, 1) != 1 || w == 0) w = 0xFFFD;
        g.ansi[i] = (i < 0x80) ? (u16)i : (u16)w;
    }
    g.ansiReady = true;
}

LayoutCfg cfg() {
    LayoutCfg c;
    c.wrapCols = g.wrap ? std::max(g.visCols - 1, 8) : 0;
    c.tabw = 4;
    c.ansi = g.d->enc == FE_ANSI ? g.ansi : nullptr;
    return c;
}

UINT docCodepage() { return g.d->enc == FE_ANSI ? g.ansiCp : CP_UTF8; }

std::string toDocBytes(const wchar_t* w, int n) {
    std::string s;
    if (n <= 0) return s;
    int k = WideCharToMultiByte(docCodepage(), 0, w, n, nullptr, 0, nullptr, nullptr);
    if (k <= 0) return s;
    s.resize((size_t)k);
    WideCharToMultiByte(docCodepage(), 0, w, n, &s[0], k, nullptr, nullptr);
    return s;
}

std::wstring fromDocBytes(const std::string& s) {
    std::wstring w;
    if (s.empty()) return w;
    int k = MultiByteToWideChar(docCodepage(), 0, s.data(), (int)s.size(), nullptr, 0);
    if (k <= 0) return w;
    w.resize((size_t)k);
    MultiByteToWideChar(docCodepage(), 0, s.data(), (int)s.size(), &w[0], k);
    return w;
}

std::wstring widenUtf8(const std::string& s) {
    std::wstring w;
    if (s.empty()) return w;
    int k = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
    if (k <= 0) return w;
    w.resize((size_t)k);
    MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), &w[0], k);
    return w;
}

std::wstring fileNameOf(const std::wstring& p) {
    if (p.empty()) return L"Untitled";
    size_t i = p.find_last_of(L"\\/");
    return i == std::wstring::npos ? p : p.substr(i + 1);
}

void errorBox(const wchar_t* what, const std::wstring& path, DWORD err) {
    wchar_t sys[512] = L"";
    if (err)
        FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, err, 0, sys, 512,
                       nullptr);
    std::wstring msg = std::wstring(what) + L"\n\n" + path;
    if (sys[0]) msg += std::wstring(L"\n\n") + sys;
    MessageBoxW(g.frame, msg.c_str(), kAppName, MB_OK | MB_ICONERROR);
}

std::wstring humanSize(u64 n) {
    wchar_t b[64];
    if (n < 1024) swprintf(b, 64, L"%llu bytes", (unsigned long long)n);
    else if (n < (1ull << 20)) swprintf(b, 64, L"%.1f KB", (double)n / 1024.0);
    else if (n < (1ull << 30)) swprintf(b, 64, L"%.1f MB", (double)n / (double)(1ull << 20));
    else swprintf(b, 64, L"%.2f GB", (double)n / (double)(1ull << 30));
    return b;
}

COLORREF blend(COLORREF a, COLORREF b, int pctA) {
    int r = (GetRValue(a) * pctA + GetRValue(b) * (100 - pctA)) / 100;
    int gg = (GetGValue(a) * pctA + GetGValue(b) * (100 - pctA)) / 100;
    int bb = (GetBValue(a) * pctA + GetBValue(b) * (100 - pctA)) / 100;
    return RGB(r, gg, bb);
}

// ====================================================================== fonts and metrics

void makeFonts() {
    if (g.font) DeleteObject(g.font);
    LOGFONTW lf = {};
    lf.lfHeight = -MulDiv(g.fontPt, g.dpi, 72);
    lf.lfWeight = FW_NORMAL;
    lf.lfCharSet = DEFAULT_CHARSET;
    lf.lfQuality = CLEARTYPE_QUALITY;
    lf.lfPitchAndFamily = FIXED_PITCH | FF_MODERN;
    lstrcpynW(lf.lfFaceName, g.fontFace.c_str(), LF_FACESIZE);
    g.font = CreateFontIndirectW(&lf);

    HDC dc = GetDC(g.frame);
    HGDIOBJ old = SelectObject(dc, g.font);
    TEXTMETRICW tm = {};
    GetTextMetricsW(dc, &tm);
    SelectObject(dc, old);
    ReleaseDC(g.frame, dc);
    g.ch = std::max<int>(1, tm.tmHeight + tm.tmExternalLeading);
    g.cw = std::max<int>(1, tm.tmAveCharWidth);

    if (g.uiFont) DeleteObject(g.uiFont);
    NONCLIENTMETRICSW ncm = {};
    ncm.cbSize = sizeof(ncm);
    if (SystemParametersInfoForDpi(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0, (UINT)g.dpi)) {
        g.uiFont = CreateFontIndirectW(&ncm.lfStatusFont);
        int h = ncm.lfStatusFont.lfHeight < 0 ? -ncm.lfStatusFont.lfHeight : ncm.lfStatusFont.lfHeight;
        g.statusH = h + MulDiv(12, g.dpi, 96);
    } else {
        g.uiFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        g.statusH = MulDiv(22, g.dpi, 96);
    }
}

void computeMetrics() {
    RECT rc;
    GetClientRect(g.view, &rc);
    g.cliW = rc.right;
    g.cliH = rc.bottom;
    if (g.lineNumbers) {
        u64 est = g.d->doc.linesKnown() ? g.d->doc.newlineCount() + 1 : std::max<u64>(g.d->doc.size() / 32, 9999);
        int digits = 1;
        while (est >= 10) {
            est /= 10;
            digits++;
        }
        if (digits < 3) digits = 3;
        g.gutterPx = (digits + 2) * g.cw;
    } else {
        g.gutterPx = g.cw / 2;
    }
    g.visCols = std::max(1, (g.cliW - g.gutterPx) / g.cw);
    g.visRows = std::max(1, g.cliH / g.ch);
}

// ---- aligned diff rows
//
// Each change gets max(lines removed, lines added) rows on both sides; the
// shorter side is padded with filler rows after its lines. Outside changes
// the two sides match line for line, so a row on the left is always level
// with the same row on the right and both scroll exactly together.

bool alignOn() { return g.diffOn && (!g.d->fillA.empty() || !g.d->fillB.empty()); }

u64 fillsUpTo(const std::vector<DocState::Fill>& f, u64 line) {  // filler rows before or at `line`
    auto it = std::upper_bound(f.begin(), f.end(), line, [](u64 l, const DocState::Fill& x) { return l < x.line; });
    return it == f.begin() ? 0 : (it - 1)->cum;
}

u64 fillAt(const std::vector<DocState::Fill>& f, u64 line) {  // filler rows right before `line`
    auto it = std::lower_bound(f.begin(), f.end(), line, [](const DocState::Fill& x, u64 l) { return x.line < l; });
    return it != f.end() && it->line == line ? it->count : 0;
}

u64 alignedA(u64 a) { return a + fillsUpTo(g.d->fillA, a); }
u64 alignedB(u64 b) { return b + fillsUpTo(g.d->fillB, b); }

void buildFills() {
    g.d->fillA.clear();
    g.d->fillB.clear();
    for (const Hunk& h : g.d->hunks) {
        if (h.aCount > h.bCount) g.d->fillB.push_back({(u64)h.bStart + h.bCount, (u64)(h.aCount - h.bCount), 0});
        if (h.bCount > h.aCount) g.d->fillA.push_back({(u64)h.aStart + h.aCount, (u64)(h.bCount - h.aCount), 0});
    }
    for (auto* f : {&g.d->fillA, &g.d->fillB}) {
        u64 cum = 0;
        for (auto& x : *f) x.cum = (cum += x.count);
    }
}

// First aligned row shown at the top of both panes.
u64 alignedTop() {
    if (!g.d->doc.linesKnown()) return 0;
    u64 t = alignedB(g.d->doc.lineOf(g.d->top));
    return t >= (u64)g.d->topFill ? t - (u64)g.d->topFill : 0;
}

void relayout() {
    computeMetrics();
    LayoutCfg c = cfg();
    u64 size = g.d->doc.size();
    if (g.d->top > size) g.d->top = size;
    if (g.d->caret > size) g.d->caret = size;
    if (g.d->anchor > size) g.d->anchor = size;
    g.d->top = rowContaining(g.d->doc, g.d->top, c).start;
    g.rows.clear();
    g.rowFill.clear();
    const bool al = alignOn() && g.d->doc.linesKnown();
    const size_t cap = (size_t)g.visRows + 1;
    auto push = [&](const Row& r, bool fill) {
        g.rows.push_back(r);
        g.rowFill.push_back(fill ? 1 : 0);
    };
    Row r = rowAt(g.d->doc, g.d->top, c);
    u64 line = al ? g.d->doc.lineOf(g.d->top) : 0;
    if (al) {
        g.d->topFill = (int)std::min<u64>((u64)std::max(0, g.d->topFill), fillAt(g.d->fillB, line));
        for (int k = 0; k < g.d->topFill && g.rows.size() < cap; k++) push(r, true);
    } else {
        g.d->topFill = 0;
    }
    push(r, false);
    while (g.rows.size() < cap) {
        Row nx;
        const bool has = nextRow(g.d->doc, r, c, nx);
        if (al && ((r.hard && r.end < r.next) || !has)) {
            line++;
            // Filler rows stand in for lines that exist only on the left. A
            // click on one lands on the real line below it.
            u64 n = fillAt(g.d->fillB, line);
            for (u64 k = 0; k < n && g.rows.size() < cap; k++) push(has ? nx : r, true);
        }
        if (!has || g.rows.size() >= cap) break;
        r = nx;
        push(r, false);
    }
}

// Aligned diff: total rows (the longer of the two padded sides).
u64 alignedTotal() {
    u64 a = alignedA(g.d->baseHash.size()), b = alignedB(g.d->doc.newlineCount() + 1);
    return std::max(a, b);
}

// Puts aligned row `row` at the top of both panes.
void setTopRow(u64 row) {
    const u64 lines = g.d->doc.newlineCount() + 1;
    u64 lo = 0, hi = lines;  // smallest line whose aligned row is at or below `row`
    while (lo < hi) {
        u64 mid = (lo + hi) / 2;
        if (alignedB(mid) < row) lo = mid + 1;
        else hi = mid;
    }
    if (lo >= lines) lo = lines - 1;
    g.d->top = g.d->doc.lineStart(lo);
    u64 at = alignedB(lo);
    g.d->topFill = at > row ? (int)std::min<u64>(at - row, fillAt(g.d->fillB, lo)) : 0;
}

void updateScrollbars() {
    if (g.inScrollUpdate) return;
    g.inScrollUpdate = true;
    SCROLLINFO si = {};
    si.cbSize = sizeof(si);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS | SIF_DISABLENOSCROLL;
    if (alignOn() && g.d->doc.linesKnown()) {
        // In the diff the scroll bar counts aligned rows, so the thumb means
        // the same place in both panes.
        u64 total = alignedTotal();
        si.nMin = 0;
        si.nMax = (int)std::min<u64>(total + (u64)g.visRows / 2, 0x7FFFFFFF);
        si.nPage = (UINT)std::max(1, g.visRows);
        si.nPos = (int)std::min<u64>(alignedTop(), 0x7FFFFFFF);
        g.vshift = 0;
    } else {
        u64 size = g.d->doc.size();
        int sh = 0;
        while ((size >> sh) > 0x3FFFFFFFull) sh++;
        g.vshift = sh;
        u64 visBytes = g.rows.empty() ? 0 : g.rows.back().next - g.d->top;
        si.nMin = 0;
        si.nMax = (int)(size >> sh);
        si.nPage = (UINT)std::max<u64>(1, visBytes >> sh);
        si.nPos = (int)(g.d->top >> sh);
        bool allVisible = g.d->top == 0 && !g.rows.empty() && g.rows.back().next >= size &&
                          (int)g.rows.size() <= g.visRows;
        if (allVisible) si.nPage = (UINT)si.nMax + 1;
    }
    SetScrollInfo(g.view, SB_VERT, &si, TRUE);

    if (!g.wrap) {
        LayoutCfg c = cfg();
        int maxCols = 0;
        for (const Row& r : g.rows) maxCols = std::max(maxCols, colOf(g.d->doc, r, r.end, c));
        si.nMin = 0;
        si.nMax = std::max(maxCols + 1, g.d->hcol + g.visCols - 1);
        si.nPage = (UINT)g.visCols;
        si.nPos = g.d->hcol;
        SetScrollInfo(g.view, SB_HORZ, &si, TRUE);
    }
    g.inScrollUpdate = false;
}

void recomputeDiff() {
    g.d->hunks.clear();
    g.d->diffDocVersion = g.d->doc.version();
    if (!g.d->vcs.haveBase || g.d->doc.size() > kDiffMax || !g.d->doc.linesKnown()) {
        buildFills();
        return;
    }
    std::vector<u64> cur;
    hashDocLines(g.d->doc, cur);
    g.d->hunks = diffLines(g.d->baseHash, cur);
    g.d->curHash.swap(cur);
    buildFills();
}

std::wstring vcsText() {
    if (g.d->vcs.kind == VcsKind::None) return L"";
    std::wstring s = g.d->vcs.kind == VcsKind::Git ? L"git: " : L"p4: ";
    if (g.d->vcs.toolMissing) return s + (g.d->vcs.kind == VcsKind::Git ? L"git.exe not found      " : L"p4.exe not found      ");
    s += g.d->vcs.label;
    if (g.d->checkingOut) return s + L" (checking out\x2026)      ";
    if (g.d->vcs.kind == VcsKind::Perforce) s += g.d->vcs.opened ? L" (" + g.d->vcs.action + L")" : L" (not checked out)";
    else if (!g.d->vcs.tracked) s += L" (new file)";
    if (g.d->vcs.haveBase && g.d->doc.size() <= kDiffMax) {
        wchar_t b[48];
        if (g.d->hunks.empty()) swprintf(b, 48, L", no changes");
        else swprintf(b, 48, L", %d change%ls", (int)g.d->hunks.size(), g.d->hunks.size() == 1 ? L"" : L"s");
        s += b;
    }
    return s + L"      ";
}

void updateStatus() {
    LayoutCfg c = cfg();
    wchar_t b[256];
    Row r = rowContaining(g.d->doc, g.d->caret, c);
    int col = colOf(g.d->doc, r, g.d->caret, c) + 1;
    std::wstring left;
    if (g.d->doc.linesKnown()) {
        swprintf(b, 256, L"Ln %llu, Col %d", (unsigned long long)g.d->doc.lineOf(g.d->caret) + 1, col);
        left = b;
    } else {
        u64 total = g.d->index ? g.d->index->size() : 0;
        u64 done = g.d->index ? g.d->index->progressBytes() : 0;
        int pct = total ? (int)(done * 100 / total) : 0;
        swprintf(b, 256, L"Col %d    Counting lines\x2026 %d%%", col, pct);
        left = b;
    }
    if (hasSel()) {
        swprintf(b, 256, L"    %llu bytes selected", (unsigned long long)(selMax() - selMin()));
        left += b;
    }
    if (!g.hint.empty()) left = g.hint;
    static const wchar_t* const encNames[] = {L"UTF-8", L"UTF-8 with BOM", L"UTF-16 LE", L"UTF-16 BE", L"ANSI"};
    std::wstring upd;
    if (g.updatePct >= 0) {
        swprintf(b, 256, L"Downloading update\x2026 %d%%      ", g.updatePct);
        upd = b;
    }
    std::wstring langText;
    if (g.highlight && g.d->lang) langText = widenUtf8(langName(g.d->lang)) + L"      ";
    std::wstring right = upd + vcsText() + langText + humanSize(g.d->doc.size()) + L"      " + encNames[g.d->enc] + L"      " +
                         (g.d->crlf ? L"Windows (CRLF)" : L"Unix (LF)");
    if (left != g.statusLeft || right != g.statusRight) {
        g.statusLeft = left;
        g.statusRight = right;
        RECT rc;
        GetClientRect(g.frame, &rc);
        rc.top = rc.bottom - g.statusH;
        InvalidateRect(g.frame, &rc, FALSE);
    }
    std::wstring title = (g.d->doc.modified() ? L"*" : L"") + fileNameOf(g.d->path) + L" - " + kAppName;
    if (title != g.title) {
        g.title = title;
        SetWindowTextW(g.frame, title.c_str());
        invalidateTabs();  // the tab shows the modified dot too
    }
    // Perforce keeps files read-only until they are checked out: check it out
    // (or offer to, if automatic checkout is off) the first time it is edited.
    if (g.d->vcs.kind == VcsKind::Perforce && !g.d->vcs.opened && !g.d->p4Asked && g.d->doc.modified()) {
        g.d->p4Asked = true;
        PostMessageW(g.frame, WM_APP_P4ASK, 0, 0);
    }
}

void ensureCaretVisible() {
    relayout();
    LayoutCfg c = cfg();
    int full = std::max(1, g.visRows);
    bool vis = false;
    for (int i = 0; i < (int)g.rows.size() && i < full; i++) {
        if (!g.rowFill[(size_t)i] && rowHas(g.d->doc, g.rows[(size_t)i], g.d->caret)) {
            vis = true;
            break;
        }
    }
    Row r = rowContaining(g.d->doc, g.d->caret, c);
    if (!vis) {
        g.d->topFill = 0;
        if (g.d->caret < g.d->top) {
            g.d->top = r.start;
        } else {
            u64 t = r.start;
            for (int i = 0; i < full - 1; i++) {
                Row pv;
                if (!prevRow(g.d->doc, t, c, pv)) break;
                t = pv.start;
            }
            g.d->top = t;
        }
        relayout();
        // Filler rows can push the caret below the window: scroll on until it shows.
        for (int guard = 0; guard < 4096; guard++) {
            bool shown = false;
            for (int i = 0; i < (int)g.rows.size() && i < full; i++)
                if (!g.rowFill[(size_t)i] && rowHas(g.d->doc, g.rows[(size_t)i], g.d->caret)) shown = true;
            if (shown || g.d->caret < g.d->top) break;
            scrollRows(1);
            relayout();
        }
    }
    if (g.wrap) {
        g.d->hcol = 0;
    } else {
        int col = colOf(g.d->doc, r, g.d->caret, c);
        if (col < g.d->hcol) g.d->hcol = std::max(0, col - 8);
        else if (col >= g.d->hcol + g.visCols - 1) g.d->hcol = std::max(0, col - g.visCols + 2);
    }
}

int headerH();
int tabBarH();
void invalidateTabs();
DocState* tabByVcsGen(unsigned gen);

void refresh(bool followCaret) {
    if (!g.view) return;
    if (g.d->vcs.haveBase && g.d->diffDocVersion != g.d->doc.version()) {
        if (g.d->doc.size() <= kDiffInstant) recomputeDiff();
        else SetTimer(g.frame, kTimerDiff, 150, nullptr);
    }
    if (followCaret) ensureCaretVisible();
    else relayout();
    updateScrollbars();
    InvalidateRect(g.view, nullptr, FALSE);
    if (g.diffOn && g.diffPane) {
        InvalidateRect(g.diffPane, nullptr, FALSE);
        InvalidateRect(g.strip, nullptr, FALSE);
        RECT hd = {0, 0, 32767, tabBarH() + headerH()};
        InvalidateRect(g.frame, &hd, FALSE);
        // Paint all three now, together. Otherwise, while the scroll bar
        // thumb is dragged, Windows repaints the editor (it owns the scroll
        // bar) right away but the left pane and strip only later, and the
        // two sides visibly drift apart until the mouse stops.
        UpdateWindow(g.view);
        UpdateWindow(g.diffPane);
        UpdateWindow(g.strip);
    }
    updateStatus();
}

void scrollRows(int n) {
    LayoutCfg c = cfg();
    if (alignOn() && g.d->doc.linesKnown()) {
        // Step through filler rows too, one row at a time, so the editor and
        // the left pane move by exactly the same amount.
        for (int i = 0; i < std::abs(n); i++) {
            if (n > 0) {
                if (g.d->topFill > 0) {
                    g.d->topFill--;
                    continue;
                }
                Row nx;
                if (!nextRow(g.d->doc, rowAt(g.d->doc, g.d->top, c), c, nx)) break;
                g.d->top = nx.start;
                g.d->topFill = (int)fillAt(g.d->fillB, g.d->doc.lineOf(g.d->top));
            } else {
                if ((u64)g.d->topFill < fillAt(g.d->fillB, g.d->doc.lineOf(g.d->top))) {
                    g.d->topFill++;
                    continue;
                }
                Row pv;
                if (!prevRow(g.d->doc, g.d->top, c, pv)) break;
                g.d->top = pv.start;
                g.d->topFill = 0;
            }
        }
        return;
    }
    Row r = rowAt(g.d->doc, g.d->top, c);
    if (n > 0) {
        for (int i = 0; i < n; i++) {
            Row nx;
            if (!nextRow(g.d->doc, r, c, nx)) break;
            r = nx;
        }
    } else {
        for (int i = 0; i < -n; i++) {
            Row pv;
            if (!prevRow(g.d->doc, r.start, c, pv)) break;
            r = pv;
        }
    }
    g.d->top = r.start;
}

// ====================================================================== painting

void ensureBackBuffer(HDC ref) {
    if (g.memDC && g.memW >= g.cliW && g.memH >= g.cliH) return;
    if (g.memDC) {
        SelectObject(g.memDC, g.memOld);
        DeleteObject(g.memBmp);
        DeleteDC(g.memDC);
    }
    g.memW = std::max(g.cliW, 1) + 64;
    g.memH = std::max(g.cliH, 1) + 64;
    g.memDC = CreateCompatibleDC(ref);
    g.memBmp = CreateCompatibleBitmap(ref, g.memW, g.memH);
    g.memOld = SelectObject(g.memDC, g.memBmp);
}

// ---- syntax highlighting glue

const u64 kHlStep = 256;               // keep the tokenizer state every 256 lines
const u64 kHlLineCap = 64u << 10;      // longer lines are coloured only in their first 64 KB
const u64 kHlBudgetLines = 200000;     // most lines worked through per paint to reach the view

bool hlOn() { return g.highlight && g.d && g.d->lang; }

// Colours (or only follows the state through) line `line` of `d`.
u32 hlLex(const Doc& d, const Lang* lang, u64 line, u32 st, std::vector<u8>* kinds) {
    u64 s = d.lineStart(line), e = d.lineStart(line + 1);
    u64 n = std::min(e - s, kHlLineCap);
    std::string t = d.text(s, n);
    size_t len = t.size();
    while (len && (t[len - 1] == '\n' || t[len - 1] == '\r')) len--;
    if (kinds) kinds->assign(len, (u8)HL_TEXT);
    return highlightLine(lang, (const u8*)t.data(), len, st, kinds ? kinds->data() : nullptr);
}

// Tokenizer state at the start of `line`, from the nearest saved state. Far
// below what has been worked out (huge files), it works forward a bounded
// amount per paint and starts fresh a little above the line meanwhile.
u32 hlStateAt(const Doc& d, const Lang* lang, DocState::HlCache& c, u64 line) {
    if (c.states.empty()) c.states.push_back(0);
    const u64 need = line / kHlStep;
    u64 done = 0;
    while (c.states.size() <= need && done < kHlBudgetLines) {
        u64 l0 = (c.states.size() - 1) * kHlStep;
        u32 st = c.states.back();
        for (u64 l = l0; l < l0 + kHlStep; l++) st = hlLex(d, lang, l, st, nullptr);
        c.states.push_back(st);
        done += kHlStep;
    }
    u64 from;
    u32 st;
    if (c.states.size() > need) {
        from = need * kHlStep;
        st = c.states[need];
    } else {
        from = line > 64 ? line - 64 : 0;
        st = 0;
    }
    for (u64 l = from; l < line; l++) st = hlLex(d, lang, l, st, nullptr);
    return st;
}

// Drops saved states that an edit may have changed.
void hlSync(const Doc& d, DocState::HlCache& c) {
    if (c.version == d.version()) return;
    c.version = d.version();
    if (d.size() < (8u << 20) || !d.linesKnown()) {
        c.states.clear();  // small: start over, it is quick
        return;
    }
    // Large: keep what lies above the caret and the view (edits happen there).
    u64 keep = std::min(d.lineOf(std::min(g.d->caret, g.d->top)) / kHlStep + 1, (u64)c.states.size());
    c.states.resize((size_t)keep);
}

// Colours for one paint: lines in order, each line lexed once.
struct HlPainter {
    const Doc* d = nullptr;
    const Lang* lang = nullptr;
    DocState::HlCache* cache = nullptr;
    u64 line = NPOS;
    u32 endState = 0;
    std::vector<u8> kinds;

    // Kinds for the row starting at byte `rowStart` of line `ln`, or null.
    const u8* row(u64 ln, u64 rowStart, u64 rowEnd) {
        if (!lang || !d->linesKnown()) return nullptr;
        if (ln != line) {
            u32 st = (line != NPOS && ln == line + 1) ? endState : hlStateAt(*d, lang, *cache, ln);
            endState = hlLex(*d, lang, ln, st, &kinds);
            kinds.resize(kinds.size() + 2, (u8)HL_TEXT);  // a line break the row may include
            line = ln;
        }
        u64 off = rowStart - d->lineStart(ln);
        // Rows past the coloured part of a very long line stay plain.
        return off + (rowEnd - rowStart) <= kinds.size() ? kinds.data() + off : nullptr;
    }
};

COLORREF hlColor(u8 kind) {
    // Visual Studio-like colours for each theme.
    static const COLORREF dark[HL_COUNT] = {
        0,                   // text
        RGB(86, 156, 214),   // keyword
        RGB(78, 201, 176),   // type
        RGB(86, 156, 214),   // literal
        RGB(220, 220, 170),  // function
        RGB(214, 157, 133),  // string
        RGB(181, 206, 168),  // number
        RGB(106, 153, 85),   // comment
        RGB(189, 99, 197),   // preprocessor / macros
        RGB(86, 156, 214),   // tag
        RGB(156, 220, 254),  // attribute / key
        RGB(181, 206, 168),  // inserted
        RGB(240, 120, 110),  // deleted
    };
    static const COLORREF light[HL_COUNT] = {
        0,
        RGB(0, 0, 255),
        RGB(43, 145, 175),
        RGB(0, 0, 255),
        RGB(116, 83, 31),
        RGB(163, 21, 21),
        RGB(9, 134, 88),
        RGB(0, 128, 0),
        RGB(128, 0, 128),
        RGB(128, 0, 0),
        RGB(200, 30, 30),
        RGB(9, 134, 88),
        RGB(190, 30, 30),
    };
    if (kind >= HL_COUNT) return g.pal.fg;
    return g.pal.dark ? dark[kind] : light[kind];
}

std::string utf8Of(const std::wstring& w) {
    int k = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string s((size_t)std::max(k, 0), '\0');
    if (k > 0) WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), &s[0], k, nullptr, nullptr);
    return s;
}

// Picks the language of the open file from its name (or a "#!" first line).
void detectLang() {
    std::string first = g.d->doc.text(0, std::min<u64>(g.d->doc.size(), 200));
    size_t nl = first.find('\n');
    if (nl != std::string::npos) first.resize(nl);
    g.d->lang = langForFile(utf8Of(g.d->path), first);
    g.d->hl = DocState::HlCache();
    g.d->hlBase = DocState::HlCache();
}

// Draws one row. `kinds`, if given, has one HlKind per byte of the row
// (from row.start); runs of the same kind are drawn in their colour.
void drawRowText(HDC dc, const Doc& doc, const Row& row, int x0, int xMax, int y, int hcol, int cols,
                 const LayoutCfg& c, const u8* kinds = nullptr) {
    static std::vector<u8> bytes;
    static std::vector<wchar_t> text;
    static std::vector<INT> dx;
    static std::vector<u8> runKind;  // kind of each wchar in `text`
    runKind.clear();
    u64 n = row.end - row.start;
    if (n == 0) return;
    if (bytes.size() < n + 8) bytes.resize((size_t)n + 8);
    doc.copy(row.start, bytes.data(), n);
    const u8* b = bytes.data();
    text.clear();
    dx.clear();
    const int first = hcol, last = hcol + cols + 1;
    int col = 0;
    u64 i = 0;
    while (i < n && col < last) {
        u32 cp;
        const u8 kind = kinds ? kinds[i] : (u8)HL_TEXT;
        int len = decodeChar(b + i, n - i, c.ansi, cp);
        int w = cellsFor(cp, col, c.tabw);
        int c0 = col, c1 = col + w;
        i += (u64)len;
        col = c1;
        if (c1 <= first) continue;  // left of the visible area
        if (cp == '\t' || c0 < first) {
            // Tabs, and wide characters cut by the left edge, become blank cells.
            int visible = c1 - std::max(c0, first);
            for (int k = 0; k < visible; k++) {
                text.push_back(L' ');
                dx.push_back(g.cw);
                runKind.push_back(kind);
            }
            continue;
        }
        if (cp < 0x20 || cp == 0x7F) cp = 0xB7;  // control characters show as a middle dot
        if (cp < 0x10000) {
            text.push_back((wchar_t)cp);
            dx.push_back(w * g.cw);
            runKind.push_back(kind);
        } else {
            u32 v = cp - 0x10000;
            text.push_back((wchar_t)(0xD800 + (v >> 10)));
            dx.push_back(w * g.cw);
            text.push_back((wchar_t)(0xDC00 + (v & 0x3FF)));
            dx.push_back(0);
            runKind.push_back(kind);
            runKind.push_back(kind);
        }
    }
    if (text.empty()) return;
    RECT rc = {x0, y, xMax, y + g.ch};
    if (!kinds) {
        ExtTextOutW(dc, x0, y, ETO_CLIPPED, &rc, text.data(), (UINT)text.size(), dx.data());
        return;
    }
    // One call per run of the same colour.
    const COLORREF plain = GetTextColor(dc);
    int x = x0;
    size_t a = 0;
    while (a < text.size()) {
        size_t e = a;
        int w = 0;
        while (e < text.size() && runKind[e] == runKind[a]) w += dx[e++];
        if (runKind[a] != HL_TEXT || text[a] != L' ' || e - a > 0) {
            SetTextColor(dc, runKind[a] == HL_TEXT ? plain : hlColor(runKind[a]));
            ExtTextOutW(dc, x, y, ETO_CLIPPED, &rc, text.data() + a, (UINT)(e - a), dx.data() + a);
        }
        x += w;
        a = e;
    }
    SetTextColor(dc, plain);
}

void paintView(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(hwnd, &ps);
    ensureBackBuffer(wdc);
    HDC dc = g.memDC;
    LayoutCfg c = cfg();

    const COLORREF bg = g.pal.bg, fg = g.pal.fg;
    const COLORREF selBg = g.pal.selection;
    const COLORREF dim = blend(fg, bg, 45);

    RECT all = {0, 0, g.cliW, g.cliH};
    SetBkColor(dc, bg);
    ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &all, L"", 0, nullptr);
    HGDIOBJ oldFont = SelectObject(dc, g.font);
    SetBkMode(dc, TRANSPARENT);

    const u64 s = selMin(), e = selMax();
    const bool linesOk = g.d->doc.linesKnown();
    HlPainter hp;
    if (hlOn() && linesOk) {
        hlSync(g.d->doc, g.d->hl);
        hp.d = &g.d->doc;
        hp.lang = g.d->lang;
        hp.cache = &g.d->hl;
    }
    const bool knowLines = g.lineNumbers && linesOk;
    u64 lineNo = linesOk ? g.d->doc.lineOf(g.d->top) + 1 : 0;  // line of the current row, 1-based
    bool startsLine = g.d->top == 0 || g.d->doc.at(g.d->top - 1) == '\n';

    for (size_t i = 0; i < g.rows.size(); i++) {
        const Row& r = g.rows[i];
        int y = (int)i * g.ch;
        if (y >= g.cliH) break;
        if (g.rowFill[i]) {
            // Lines that exist only on the left: a blank row in the change's colour.
            u64 below = linesOk ? g.d->doc.lineOf(r.start) : 0;
            RECT f = {0, y, g.cliW, y + g.ch};
            const COLORREF fc = fillerColorB(below);
            SetBkColor(dc, fillTintOf(fc));
            ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &f, L"", 0, nullptr);
            if ((i == 0 && g.d->topFill == 0) || (i > 0 && !g.rowFill[i - 1])) {
                // Top of a run of fillers: a thin line marks where the lines were.
                RECT e = {0, y, g.cliW, y + edgeW()};
                SetBkColor(dc, edgeOf(fc));
                ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &e, L"", 0, nullptr);
            }
            continue;
        }

        const bool endsLine = r.hard && r.end < r.next;
        // Change marker against the base revision.
        if (linesOk && !g.d->hunks.empty()) {
            int hi = hunkAtB(g.d->hunks, (u32)(lineNo - 1));
            if (hi >= 0) {
                const Hunk& h = g.d->hunks[(size_t)hi];
                int bw = std::max(2, MulDiv(3, g.dpi, 96));
                int bx = g.lineNumbers ? g.gutterPx - g.cw / 2 - bw / 2 : 0;
                if (h.bCount == 0) {
                    if (startsLine) {  // lines were removed just above this one
                        RECT m = {bx - bw, y - 1, bx + 2 * bw, y + bw - 1};
                        SetBkColor(dc, kRemoved);
                        ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &m, L"", 0, nullptr);
                    }
                } else {
                    COLORREF col = changeColor(h.aCount, h.bCount);
                    if (g.diffOn) {
                        // The whole row, line number included.
                        fillRect(dc, 0, y, g.cliW, y + g.ch, tintOf(col));
                        const u64 cur = lineNo - 1;
                        const int ew = edgeW();
                        if (startsLine && cur == h.bStart) fillRect(dc, 0, y, g.cliW, y + ew, edgeOf(col));
                        if (endsLine && cur + 1 == (u64)h.bStart + h.bCount)
                            fillRect(dc, 0, y + g.ch - ew, g.cliW, y + g.ch, edgeOf(col));
                        // What changed within the line, against the matching base line.
                        if (h.aCount && cur - h.bStart < h.aCount) {
                            u64 x0, x1;
                            if (inlineSpan(g.d->doc, r.start, r.end, g.d->baseDoc, h.aStart + (cur - h.bStart), x0, x1)) {
                                int c0 = colOf(g.d->doc, r, x0, c), c1 = colOf(g.d->doc, r, x1, c);
                                int px0 = std::max(g.gutterPx, g.gutterPx + (c0 - g.d->hcol) * g.cw);
                                int px1 = std::min(g.cliW, g.gutterPx + (c1 - g.d->hcol) * g.cw);
                                if (px1 > px0) fillRect(dc, px0, y + ew, px1, y + g.ch - ew, inlineOf(col));
                            }
                        }
                    } else {
                        RECT m = {bx, y, bx + bw, y + g.ch};
                        SetBkColor(dc, col);
                        ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &m, L"", 0, nullptr);
                    }
                }
            }
        }
        if (knowLines && startsLine) {
            wchar_t nb[32];
            int k = swprintf(nb, 32, L"%llu", (unsigned long long)lineNo);
            RECT grc = {0, y, g.gutterPx - g.cw, y + g.ch};
            SetTextColor(dc, dim);
            DrawTextW(dc, nb, k, &grc, DT_RIGHT | DT_SINGLELINE | DT_NOPREFIX | DT_NOCLIP);
        }
        // What the next row is: a new line (number it) or a wrap continuation.

        if (s < e && e > r.start && s < r.next) {
            int c0 = s <= r.start ? 0 : colOf(g.d->doc, r, s, c);
            int c1 = e > r.end ? colOf(g.d->doc, r, r.end, c) + 1 : colOf(g.d->doc, r, e, c);
            int x0 = g.gutterPx + (c0 - g.d->hcol) * g.cw;
            int x1 = g.gutterPx + (c1 - g.d->hcol) * g.cw;
            x0 = std::max(x0, g.gutterPx);
            x1 = std::min(x1, g.cliW);
            if (x1 > x0) {
                RECT src = {x0, y, x1, y + g.ch};
                SetBkColor(dc, selBg);
                ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &src, L"", 0, nullptr);
            }
        }
        SetTextColor(dc, fg);
        drawRowText(dc, g.d->doc, r, g.gutterPx, g.cliW, y, g.d->hcol, g.visCols, c,
                    hlOn() && linesOk ? hp.row(lineNo - 1, r.start, r.end) : nullptr);

        if (g.focused && g.caretOn && rowHas(g.d->doc, r, g.d->caret)) {
            int col = colOf(g.d->doc, r, g.d->caret, c);
            int x = g.gutterPx + (col - g.d->hcol) * g.cw;
            if (x >= g.gutterPx && x < g.cliW) {
                RECT crc = {x, y, x + std::max(1, MulDiv(2, g.dpi, 96)), y + g.ch};
                SetBkColor(dc, fg);
                ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &crc, L"", 0, nullptr);
            }
        }
        startsLine = endsLine;
        if (endsLine) lineNo++;
    }
    SelectObject(dc, oldFont);
    BitBlt(wdc, ps.rcPaint.left, ps.rcPaint.top, ps.rcPaint.right - ps.rcPaint.left,
           ps.rcPaint.bottom - ps.rcPaint.top, dc, ps.rcPaint.left, ps.rcPaint.top, SRCCOPY);
    EndPaint(hwnd, &ps);
}

// ---- diff geometry and choices
//
// A "slot" is one place where the three buttons appear: a current difference,
// or a change you already settled with a button (it stays selectable until
// you submit, revert or the base revision changes).

int headerH() { return g.diffOn ? g.statusH + MulDiv(6, g.dpi, 96) : 0; }
int tabBarH() { return g.statusH + MulDiv(8, g.dpi, 96); }
int stripW() { return MulDiv(78, g.dpi, 96); }

struct Slot {
    u32 aStart, aCount, bStart, bCount;
    int choice;  // kUseDepot / kUseBoth / kKeepYours
    int rec;     // index into g.d->choices, or -1 for an untouched difference
};

u32 lineCount(const std::string& t) {
    u32 n = 0;
    for (char ch : t) n += ch == '\n';
    return n + (!t.empty() && t.back() != '\n' ? 1 : 0);
}

// Current line where base line `a` sits, for a base range that starts a
// change (h.aStart == a) or lies in an unchanged stretch.
u32 mapAtoB(u32 a) {
    i64 off = 0;
    for (const Hunk& h : g.d->hunks) {
        if (h.aStart == a && h.aCount) return h.bStart;
        if ((u64)h.aStart + h.aCount <= a && !(h.aStart == a && h.aCount == 0)) {
            off = ((i64)h.bStart + h.bCount) - ((i64)h.aStart + h.aCount);
        } else if (h.aStart > a || (h.aStart == a && h.aCount == 0)) {
            break;
        }
    }
    return (u32)std::max<i64>(0, (i64)a + off);
}

std::vector<Slot> g_slots;  // as last painted in the strip; hover / click indices refer to it

std::vector<Slot> buildSlots() {
    std::vector<Slot> out;
    for (size_t i = 0; i < g.d->choices.size(); i++) {
        const DocState::Choice& c = g.d->choices[i];
        Slot sl = {c.aStart, c.aCount, mapAtoB(c.aStart), 0, c.choice, (int)i};
        u32 mine = lineCount(c.yours);
        sl.bCount = c.choice == kUseDepot ? c.aCount : c.choice == kUseBoth ? c.aCount + mine : mine;
        out.push_back(sl);
    }
    for (const Hunk& h : g.d->hunks) {
        // Differences that belong to a settled choice ("both" leaves one, and
        // so does "keep yours") are already covered by that slot.
        bool owned = false;
        for (const Slot& o : out) {
            u32 e = o.bStart + std::max<u32>(o.bCount, 1);
            if (h.bStart >= o.bStart && h.bStart < e) owned = true;
            if (h.bCount == 0 && h.bStart == o.bStart + o.bCount && o.rec >= 0 &&
                g.d->choices[(size_t)o.rec].choice != kUseDepot)
                owned = true;
        }
        if (!owned) out.push_back({h.aStart, h.aCount, h.bStart, h.bCount, kKeepYours, -1});
    }
    std::sort(out.begin(), out.end(), [](const Slot& x, const Slot& y) { return x.bStart < y.bStart; });
    return out;
}

bool diffReady() {
    return g.d->vcs.haveBase && !g.d->path.empty() && g.d->doc.size() <= kDiffMax && g.d->doc.linesKnown();
}

void fillRect(HDC dc, int l, int t, int r, int b, COLORREF c) {
    RECT rc = {l, t, r, b};
    SetBkColor(dc, c);
    ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &rc, L"", 0, nullptr);
}

// One colour scheme for both panes and the strip, so a change reads as a
// single shape flowing from one side to the other.
COLORREF changeColor(u32 aCount, u32 bCount) {
    return aCount == 0 ? kAdded : bCount == 0 ? kRemoved : kChanged;
}
COLORREF tintOf(COLORREF c) { return blend(c, g.pal.bg, g.pal.dark ? 30 : 17); }      // changed lines
COLORREF fillTintOf(COLORREF c) { return blend(c, g.pal.bg, g.pal.dark ? 14 : 8); }   // filler rows
COLORREF edgeOf(COLORREF c) { return blend(c, g.pal.bg, g.pal.dark ? 70 : 60); }      // outlines
int edgeW() { return std::max(1, MulDiv(1, g.dpi, 96)); }  // one border width everywhere
COLORREF inlineOf(COLORREF c) { return blend(c, g.pal.bg, g.pal.dark ? 55 : 38); }  // changed characters

// The part of line [s, e) of `doc` that differs from line `otherLine` of
// `other`: everything between their common start and common end.
bool inlineSpan(const Doc& doc, u64 s, u64 e, const Doc& other, u64 otherLine, u64& x0, u64& x1) {
    const u64 cap = 1u << 16;  // very long lines: not worth it
    if (e - s > cap) return false;
    u64 os = other.lineStart(otherLine), oe = other.lineStart(otherLine + 1);
    std::string mine = doc.text(s, e - s), theirs = other.text(os, std::min(oe - os, cap));
    while (!theirs.empty() && (theirs.back() == '\n' || theirs.back() == '\r')) theirs.pop_back();
    while (!mine.empty() && (mine.back() == '\n' || mine.back() == '\r')) mine.pop_back();
    size_t p = 0, n = std::min(mine.size(), theirs.size());
    while (p < n && mine[p] == theirs[p]) p++;
    size_t q = 0;
    while (q < n - p && mine[mine.size() - 1 - q] == theirs[theirs.size() - 1 - q]) q++;
    // Do not split a UTF-8 character.
    while (p > 0 && p < mine.size() && ((u8)mine[p] & 0xC0) == 0x80) p--;
    size_t end = mine.size() - q;
    while (end < mine.size() && ((u8)mine[end] & 0xC0) == 0x80) end++;
    if (end <= p) return false;  // nothing of this line's own differs (only removed text)
    if (p == 0 && end == mine.size()) return false;  // the whole line differs: the row tint says it
    x0 = s + p;
    x1 = s + end;
    return true;
}

// First aligned row of a change. A side with no lines has its filler rows at
// that position, which come before the row its next real line sits on.
u64 changeTopRow(u32 aStart, u32 aCount, u32 bStart, u32 bCount) {
    if (aCount == 0) return alignedA(aStart) - fillAt(g.d->fillA, aStart);
    if (bCount == 0) return alignedB(bStart) - fillAt(g.d->fillB, bStart);
    return alignedA(aStart);
}

// Colour of the change that filler rows before current line `line` belong to.
COLORREF fillerColorB(u64 line) {
    for (const Hunk& h : g.d->hunks)
        if ((u64)h.bStart + h.bCount == line && h.aCount > h.bCount) return changeColor(h.aCount, h.bCount);
    return kRemoved;
}

const wchar_t* diffNote() {
    if (g.d->path.empty()) return L"Save the file to compare it with version control.";
    if (g.d->vcs.kind == VcsKind::None)
        return g.d->vcsPending ? L"Checking version control\x2026" : L"This file is not under Git or Perforce.";
    if (g.d->vcs.toolMissing) return g.d->vcs.kind == VcsKind::Git ? L"git.exe was not found." : L"p4.exe was not found.";
    if (!g.d->vcs.haveBase)
        return !g.d->vcs.tracked             ? L"New file: there is no earlier revision to compare with."
               : !g.d->vcs.baseError.empty() ? g.d->vcs.baseError.c_str()
                                          : L"The base revision is too large to compare.";
    if (g.d->doc.size() > kDiffMax) return L"The file is too large to compare (limit 16 MB).";
    if (!g.d->doc.linesKnown()) return L"Counting lines\x2026";
    return nullptr;
}

// Title bar above the two sides: what each side is, and the change navigator.
void paintHeader(HDC dc) {
    RECT rc;
    GetClientRect(g.frame, &rc);
    const int H = headerH(), sw = stripW(), left = std::max(0, ((int)rc.right - sw) / 2);
    const COLORREF face = g.pal.face, text = g.pal.faceText;
    fillRect(dc, 0, 0, rc.right, H, face);
    fillRect(dc, 0, H - 1, rc.right, H, blend(text, face, 20));
    const int pad = MulDiv(8, g.dpi, 96), sq = MulDiv(10, g.dpi, 96);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, text);
    auto side = [&](int x0, int x1, COLORREF col, const std::wstring& label) {
        int cy = H / 2;
        fillRect(dc, x0 + pad, cy - sq / 2, x0 + pad + sq, cy - sq / 2 + sq, col);
        RECT t = {x0 + pad + sq + pad, 0, x1 - pad, H};
        DrawTextW(dc, label.c_str(), (int)label.size(), &t,
                  DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_PATH_ELLIPSIS);
    };
    std::wstring base;
    if (g.d->vcs.kind == VcsKind::Perforce) base = L"Depot  " + (g.d->vcs.baseFrom.empty() ? g.d->vcs.label : g.d->vcs.baseFrom);
    else if (g.d->vcs.kind == VcsKind::Git) base = L"Committed  HEAD (" + g.d->vcs.label + L")";
    else base = L"Base";
    side(0, left, kDepotCol, base);
    side(left + sw, rc.right, kYoursCol, std::wstring(L"Yours  ") + (g.d->doc.modified() ? L"(unsaved)  " : L"") + g.d->path);

    // Previous / next change: two 3D buttons above the strip.
    const int bs = std::min(H - MulDiv(8, g.dpi, 96), MulDiv(20, g.dpi, 96)), gap = MulDiv(4, g.dpi, 96);
    const int bx = left + (sw - 2 * bs - gap) / 2, by = (H - bs) / 2;
    g.navPrev = {bx, by, bx + bs, by + bs};
    g.navNext = {bx + bs + gap, by, bx + 2 * bs + gap, by + bs};
    const bool any = !g.d->hunks.empty() || !g.d->choices.empty();
    const COLORREF arrow = any ? g.pal.faceText : blend(g.pal.faceText, face, 35);
    for (int k = 1; k <= 2; k++) {
        const RECT& r = k == 1 ? g.navPrev : g.navNext;
        const bool hot = any && g.navHover == k, held = hot && g.navPress == k;
        drawButton3D(dc, r, held ? BTN_PRESSED : hot ? BTN_HOT : BTN_NORMAL, g.pal, 0);
        drawGlyph(dc, r, k == 1 ? GLYPH_UP : GLYPH_DOWN, arrow, arrow, held);
    }
}

void paintStatus(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(hwnd, &ps);
    // Everything goes to an off-screen bitmap first and reaches the screen in
    // one copy, so hovering the tabs or buttons never flickers.
    RECT full;
    GetClientRect(hwnd, &full);
    HDC dc = CreateCompatibleDC(wdc);
    HBITMAP bmp = CreateCompatibleBitmap(wdc, std::max(1, (int)full.right), std::max(1, (int)full.bottom));
    HGDIOBJ oldBmp = SelectObject(dc, bmp);
    RECT rc = full;
    rc.top = rc.bottom - g.statusH;
    SetBkColor(dc, g.pal.face);
    ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &rc, L"", 0, nullptr);
    RECT line = {rc.left, rc.top, rc.right, rc.top + 1};
    SetBkColor(dc, blend(g.pal.faceText, g.pal.face, 20));
    ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &line, L"", 0, nullptr);
    HGDIOBJ old = SelectObject(dc, g.uiFont);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, g.pal.faceText);
    int pad = MulDiv(10, g.dpi, 96);
    RECT t = {rc.left + pad, rc.top + 1, rc.right - pad, rc.bottom};
    DrawTextW(dc, g.statusRight.c_str(), (int)g.statusRight.size(), &t,
              DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    DrawTextW(dc, g.statusLeft.c_str(), (int)g.statusLeft.size(), &t,
              DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    paintTabs(dc);
    if (g.diffOn) {
        // Drawn off-screen at (0, 0), then copied below the tab bar.
        RECT cr;
        GetClientRect(hwnd, &cr);
        const int hh = headerH();
        HDC mdc = CreateCompatibleDC(dc);
        HBITMAP mb = CreateCompatibleBitmap(dc, std::max(1, (int)cr.right), std::max(1, hh));
        HGDIOBJ ob = SelectObject(mdc, mb);
        HGDIOBJ of2 = SelectObject(mdc, g.uiFont);
        paintHeader(mdc);
        BitBlt(dc, 0, tabBarH(), cr.right, hh, mdc, 0, 0, SRCCOPY);
        SelectObject(mdc, of2);
        SelectObject(mdc, ob);
        DeleteObject(mb);
        DeleteDC(mdc);
        OffsetRect(&g.navPrev, 0, tabBarH());
        OffsetRect(&g.navNext, 0, tabBarH());
    }
    SelectObject(dc, old);
    BitBlt(wdc, ps.rcPaint.left, ps.rcPaint.top, ps.rcPaint.right - ps.rcPaint.left, ps.rcPaint.bottom - ps.rcPaint.top, dc,
           ps.rcPaint.left, ps.rcPaint.top, SRCCOPY);
    SelectObject(dc, oldBmp);
    DeleteObject(bmp);
    DeleteDC(dc);
    EndPaint(hwnd, &ps);
}

// ====================================================================== caret and editing

void moveCaret(u64 pos, bool extend, bool keepWantCol = false) {
    g.d->caret = std::min(pos, g.d->doc.size());
    if (!extend) g.d->anchor = g.d->caret;
    if (!keepWantCol) g.d->wantCol = -1;
    g.caretOn = true;
    g.d->doc.breakCoalesce();
    refresh(true);
}

void afterEdit(u64 caret) {
    g.d->caret = g.d->anchor = std::min(caret, g.d->doc.size());
    g.d->wantCol = -1;
    g.caretOn = true;
    refresh(true);
}

void replaceSel(const u8* s, u64 n, bool coalesce) {
    if (hasSel()) {
        u64 a = selMin(), len = selMax() - a;
        g.d->doc.beginGroup();
        g.d->doc.erase(a, len);
        g.d->doc.insert(a, s, n);
        g.d->doc.endGroup();
        afterEdit(a + n);
    } else if (n) {
        u64 at = g.d->caret;
        g.d->doc.insert(at, s, n, coalesce);
        afterEdit(at + n);
    }
}

void moveVertical(int rowsDelta, bool extend, bool scrollToo) {
    LayoutCfg c = cfg();
    Row r = rowContaining(g.d->doc, g.d->caret, c);
    if (g.d->wantCol < 0) g.d->wantCol = colOf(g.d->doc, r, g.d->caret, c);
    int moved = 0;
    if (rowsDelta > 0) {
        for (; moved < rowsDelta; moved++) {
            Row nx;
            if (!nextRow(g.d->doc, r, c, nx)) break;
            r = nx;
        }
    } else {
        for (; moved < -rowsDelta; moved++) {
            Row pv;
            if (!prevRow(g.d->doc, r.start, c, pv)) break;
            r = pv;
        }
    }
    if (scrollToo) scrollRows(rowsDelta > 0 ? moved : -moved);
    moveCaret(posOfCol(g.d->doc, r, g.d->wantCol, c), extend, true);
}

void typeNewline() {
    LayoutCfg c = cfg();
    std::string s = g.d->crlf ? "\r\n" : "\n";
    // Auto-indent: repeat the leading whitespace of the current line.
    u64 at = selMin();
    u64 ls = lineStartBefore(g.d->doc, at, c);
    u8 lead[256];
    u64 n = g.d->doc.copy(ls, lead, std::min<u64>(sizeof(lead), at - ls));
    for (u64 i = 0; i < n && (lead[i] == ' ' || lead[i] == '\t'); i++) s += (char)lead[i];
    g.d->doc.breakCoalesce();
    replaceSel((const u8*)s.data(), s.size(), false);
    g.d->doc.breakCoalesce();
}

void typeChar(u32 cp) {
    u8 b[8];
    u64 n;
    if (g.d->enc == FE_ANSI) {
        wchar_t w[2];
        int wn = 1;
        if (cp < 0x10000) {
            w[0] = (wchar_t)cp;
        } else {
            w[0] = (wchar_t)(0xD800 + ((cp - 0x10000) >> 10));
            w[1] = (wchar_t)(0xDC00 + ((cp - 0x10000) & 0x3FF));
            wn = 2;
        }
        int k = WideCharToMultiByte(g.ansiCp, 0, w, wn, (char*)b, 8, nullptr, nullptr);
        if (k <= 0) return;
        n = (u64)k;
    } else {
        n = (u64)utf8Encode(cp, b);
    }
    replaceSel(b, n, true);
}

void deleteSel() {
    if (!hasSel()) return;
    u64 a = selMin();
    g.d->doc.erase(a, selMax() - a);
    afterEdit(a);
}

// Visual Studio marks whole-line copies with this clipboard format; using the
// same name means line copies paste as lines in both editors.
UINT lineClipFormat() {
    static UINT f = RegisterClipboardFormatW(L"VisualStudioEditorOperationsLineCutCopyClipboardTag");
    return f;
}

bool setClipboard(const std::wstring& w, bool wholeLines) {
    if (!OpenClipboard(g.frame)) return false;
    EmptyClipboard();
    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE, (w.size() + 1) * sizeof(wchar_t));
    if (h) {
        wchar_t* p = (wchar_t*)GlobalLock(h);
        memcpy(p, w.c_str(), (w.size() + 1) * sizeof(wchar_t));
        GlobalUnlock(h);
        if (!SetClipboardData(CF_UNICODETEXT, h)) GlobalFree(h);
    }
    if (wholeLines) {
        HGLOBAL t = GlobalAlloc(GMEM_MOVEABLE, 1);
        if (t && !SetClipboardData(lineClipFormat(), t)) GlobalFree(t);
    }
    CloseClipboard();
    return true;
}

// Lines covered by the selection (or the caret's line), as [first, last].
bool selLines(u64& l0, u64& l1) {
    if (!g.d->doc.linesKnown()) return false;
    l0 = g.d->doc.lineOf(selMin());
    l1 = g.d->doc.lineOf(selMax());
    // A selection ending at the very start of a line does not include it.
    if (l1 > l0 && g.d->doc.lineStart(l1) == selMax()) l1--;
    return true;
}
u64 lastLine() { return g.d->doc.newlineCount(); }
const char* eolStr() { return g.d->crlf ? "\r\n" : "\n"; }

// Text of lines [l0, l1], always ending with a line break.
std::string linesText(u64 l0, u64 l1) {
    u64 a = g.d->doc.lineStart(l0), b = g.d->doc.lineStart(l1 + 1);
    std::string t = g.d->doc.text(a, b - a);
    if (t.empty() || t.back() != '\n') t += eolStr();
    return t;
}

// Replaces [a, b) with `t` as one undo step, then selects anchor..caret.
void replaceRange(u64 a, u64 b, const std::string& t, u64 anchor, u64 caret) {
    g.d->doc.breakCoalesce();
    g.d->doc.beginGroup();
    if (b > a) g.d->doc.erase(a, b - a);
    if (!t.empty()) g.d->doc.insert(a, (const u8*)t.data(), t.size());
    g.d->doc.endGroup();
    g.d->doc.breakCoalesce();
    afterEdit(caret);
    g.d->anchor = std::min(anchor, g.d->doc.size());
    refresh(true);
}

void copyLines(bool cut) {
    u64 l0, l1;
    if (!selLines(l0, l1)) return;
    setClipboard(fromDocBytes(linesText(l0, l1)), true);
    if (!cut) return;
    u64 a = g.d->doc.lineStart(l0), b = g.d->doc.lineStart(l1 + 1);
    // Cutting the last line: take the line break before it instead.
    if (l1 == lastLine() && l0 > 0) a = g.d->doc.lineStart(l0) - (g.d->crlf && a >= 2 && g.d->doc.at(a - 2) == '\r' ? 2 : 1);
    replaceRange(a, b, "", a, a);
}

void doCopy() {
    if (!hasSel()) {  // no selection: copy the whole line, like Visual Studio
        copyLines(false);
        return;
    }
    u64 n = selMax() - selMin();
    if (n > kMaxClipboard) {
        MessageBoxW(g.frame, L"The selection is too large for the clipboard (limit 256 MB).", kAppName,
                    MB_OK | MB_ICONINFORMATION);
        return;
    }
    setClipboard(fromDocBytes(g.d->doc.text(selMin(), n)), false);
}

void doPaste() {
    if (!OpenClipboard(g.frame)) return;
    std::string s;
    HANDLE h = GetClipboardData(CF_UNICODETEXT);
    if (h) {
        const wchar_t* p = (const wchar_t*)GlobalLock(h);
        if (p) {
            s = toDocBytes(p, lstrlenW(p));
            GlobalUnlock(h);
        }
    }
    const bool lines = IsClipboardFormatAvailable(lineClipFormat()) != 0;
    CloseClipboard();
    if (s.empty()) return;
    normalizeEol(s, g.d->crlf);
    if (lines && !hasSel() && g.d->doc.linesKnown()) {
        // A whole-line copy pastes above the caret's line; the caret keeps its column.
        u64 ls = g.d->doc.lineStart(g.d->doc.lineOf(g.d->caret)), off = g.d->caret - ls;
        if (s.back() != '\n') s += eolStr();
        replaceRange(ls, ls, s, ls + s.size() + off, ls + s.size() + off);
        return;
    }
    g.d->doc.breakCoalesce();
    replaceSel((const u8*)s.data(), s.size(), false);
    g.d->doc.breakCoalesce();
}

u64 posFromPoint(int x, int y) {
    if (g.rows.empty()) return 0;
    LayoutCfg c = cfg();
    int ri = y / g.ch;
    if (y < 0) ri = 0;
    if (ri >= (int)g.rows.size()) ri = (int)g.rows.size() - 1;
    int col = g.d->hcol;
    if (x > g.gutterPx) col += (x - g.gutterPx + g.cw / 2) / g.cw;
    return posOfCol(g.d->doc, g.rows[(size_t)ri], col, c);
}

// ====================================================================== files

void startIndex(const u8* data, u64 len, bool forceSync) {
    std::unique_ptr<LineIndex> idx(new LineIndex());
    KillTimer(g.frame, kTimerIndex);
    g.d->indexing = false;
    unsigned gen = g.d->indexGen = ++g_gen;
    if (forceSync || len <= kSyncIndexMax) {
        idx->buildSync(data, len);
    } else {
        HWND frame = g.frame;
        idx->buildAsync(data, len, [frame, gen]() { PostMessageW(frame, WM_APP_INDEX, (WPARAM)gen, 0); });
        g.d->indexing = true;
        SetTimer(g.frame, kTimerIndex, 200, nullptr);
    }
    // Replacing the old index stops its worker; the caller must do this before
    // releasing the memory that worker was reading.
    g.d->index = std::move(idx);
}

void clearVcs() {
    g.d->vcs = VcsInfo();
    g.d->choices.clear();
    g.d->hunks.clear();
    g.d->baseHash.clear();
    g.d->baseDoc.reset(nullptr, 0, nullptr);
    g.d->baseText.clear();
    g.d->diffDocVersion = NPOS;
}

// Asks a worker thread which version control system manages the file and for
// its base revision. The answer arrives as WM_APP_VCS; opening never waits.
void startVcsQuery(bool samePath) {
    unsigned gen = g.d->vcsGen = ++g_gen;
    if (!samePath) {
        clearVcs();
        g.d->p4Asked = false;
    }
    g.d->vcsPending = !g.d->path.empty();
    if (g.d->path.empty()) return;
    std::wstring path = g.d->path;
    HWND frame = g.frame;
    std::thread([path, gen, frame]() {
        VcsInfo* r = new VcsInfo(queryVcs(path, (size_t)kDiffMax));
        if (!PostMessageW(frame, WM_APP_VCS, (WPARAM)gen, (LPARAM)r)) delete r;
    }).detach();
}

void onVcsResult(unsigned gen, VcsInfo* raw) {
    std::unique_ptr<VcsInfo> info(raw);
    // The answer may be for a tab in the background; apply it to that tab.
    DocState* t = tabByVcsGen(gen);
    if (!t) return;  // the file was closed or asked again since
    DocState* active = g.d;
    g.d = t;
    g.d->vcsPending = false;
    std::vector<u8> oldBase;
    oldBase.swap(g.d->baseText);
    g.d->baseDoc.reset(nullptr, 0, nullptr);
    g.d->baseHash.clear();
    g.d->vcs = std::move(*info);
    if (g.d->vcs.haveBase) {
        // Bring the base revision into the same form the document is held in.
        const std::vector<u8>& b = g.d->vcs.base;
        if (b.size() >= 2 && ((b[0] == 0xFF && b[1] == 0xFE) || (b[0] == 0xFE && b[1] == 0xFF)))
            utf16ToUtf8(b.data() + 2, b.size() - 2, b[0] == 0xFE, g.d->baseText);
        else if (b.size() >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF)
            g.d->baseText.assign(b.begin() + 3, b.end());
        else
            g.d->baseText = b;
        std::vector<u8>().swap(g.d->vcs.base);
        g.d->baseIdx.buildSync(g.d->baseText.data(), g.d->baseText.size());
        g.d->baseDoc.reset(g.d->baseText.data(), g.d->baseText.size(), &g.d->baseIdx);
        LineHasher h;
        h.feed(g.d->baseText.data(), g.d->baseText.size());
        h.finish();
        g.d->baseHash.swap(h.lines);
    }
    // Choices stay until the base revision changes (submit, revert, sync).
    if (g.d->baseText != oldBase) g.d->choices.clear();
    g.d->hlBase = DocState::HlCache();
    recomputeDiff();
    g.d = active;
    if (t == active) refresh(false);
    else invalidateTabs();
}

bool loadBacking(const std::wstring& path, Backing& b, u64& size, bool forceMap, DWORD& err) {
    b.file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                         nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (b.file == INVALID_HANDLE_VALUE) {
        err = GetLastError();
        return false;
    }
    LARGE_INTEGER li;
    if (!GetFileSizeEx(b.file, &li)) {
        err = GetLastError();
        b.close();
        return false;
    }
    size = (u64)li.QuadPart;
    if (size == 0) {
        b.close();
        return true;
    }
    if (size <= kHeapMax && !forceMap) {
        b.heap.resize((size_t)size);
        u64 done = 0;
        while (done < size) {
            DWORD got = 0;
            DWORD want = (DWORD)std::min<u64>(size - done, 1u << 30);
            if (!ReadFile(b.file, b.heap.data() + done, want, &got, nullptr) || got == 0) break;
            done += got;
        }
        if (done != size) {  // the file shrank while reading; keep what we got
            size = done;
            b.heap.resize((size_t)size);
        }
        CloseHandle(b.file);
        b.file = INVALID_HANDLE_VALUE;
        return true;
    }
    b.map = CreateFileMappingW(b.file, nullptr, PAGE_READONLY, 0, 0, nullptr);
    if (b.map) b.view = (const u8*)MapViewOfFile(b.map, FILE_MAP_READ, 0, 0, 0);
    if (!b.view) {
        err = GetLastError();
        b.close();
        return false;
    }
    b.mapped = true;
    return true;
}

void resetViewState() {
    g.d->top = g.d->caret = g.d->anchor = 0;
    g.d->hcol = 0;
    g.d->wantCol = -1;
}

bool openPath(const std::wstring& path, bool keepView) {
    Backing nb;
    u64 size = 0;
    DWORD err = 0;
    if (!loadBacking(path, nb, size, false, err)) {
        errorBox(L"Could not open the file.", path, err);
        return false;
    }
    const u8* data = nb.mapped ? nb.view : nb.heap.data();
    u64 skip = 0;
    FileEnc enc = FE_UTF8;
    if (size >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) {
        enc = FE_UTF8BOM;
        skip = 3;
    } else if (size >= 2 && data[0] == 0xFF && data[1] == 0xFE) {
        enc = FE_UTF16LE;
    } else if (size >= 2 && data[0] == 0xFE && data[1] == 0xFF) {
        enc = FE_UTF16BE;
    } else if (!looksUtf8(data, std::min<u64>(size, 1u << 16))) {
        enc = FE_ANSI;
        ensureAnsiTable();
    }
    if (enc == FE_UTF16LE || enc == FE_UTF16BE) {
        // UTF-16 is converted once on load and converted back on save.
        std::vector<u8> conv;
        utf16ToUtf8(data + 2, size - 2, enc == FE_UTF16BE, conv);
        nb.close();
        nb.heap.swap(conv);
        data = nb.heap.data();
        size = nb.heap.size();
    }
    const u8* text = size ? data + skip : nullptr;
    u64 len = size - skip;

    // Line ending style: whatever the first line uses.
    bool crlf = true;
    if (len) {
        const void* q = memchr(text, '\n', (size_t)std::min<u64>(len, 1u << 16));
        if (q) crlf = q != (const void*)text && ((const u8*)q)[-1] == '\r';
    }

    // Order matters: stop the old index worker, point the document at the new
    // data, and only then release the old file.
    g.d->index.reset();
    startIndex(text, len, false);
    g.d->doc.reset(text, len, g.d->index.get());
    g.d->back.take(nb);
    g.d->bomSkip = skip;
    const bool samePath = lstrcmpiW(path.c_str(), g.d->path.c_str()) == 0;
    g.d->path = path;
    g.d->enc = enc;
    g.d->crlf = crlf;
    if (!keepView) resetViewState();
    detectLang();
    startVcsQuery(samePath);
    refresh(false);
    return true;
}

void newDocument() {
    g.d->index.reset();
    startIndex(nullptr, 0, true);
    g.d->doc.reset(nullptr, 0, g.d->index.get());
    g.d->back.close();
    g.d->bomSkip = 0;
    g.d->path.clear();
    g.d->enc = FE_UTF8;
    g.d->crlf = true;
    resetViewState();
    startVcsQuery(false);
    refresh(false);
}

struct Writer {
    HANDLE h = INVALID_HANDLE_VALUE;
    std::vector<u8> buf;
    bool ok = true;
    DWORD err = 0;

    void raw(const u8* p, u64 n) {
        while (n && ok) {
            DWORD want = (DWORD)std::min<u64>(n, 64u << 20), put = 0;
            if (!WriteFile(h, p, want, &put, nullptr) || put == 0) {
                ok = false;
                err = GetLastError();
                return;
            }
            p += put;
            n -= put;
        }
    }
    void flush() {
        if (!buf.empty()) raw(buf.data(), buf.size());
        buf.clear();
    }
    void put(const u8* p, u64 n) {
        if (n >= (1u << 20)) {
            flush();
            raw(p, n);
            return;
        }
        buf.insert(buf.end(), p, p + n);
        if (buf.size() >= (1u << 20)) flush();
    }
};

bool writeDoc(const std::wstring& tmp, DWORD& err) {
    Writer w;
    w.h = CreateFileW(tmp.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (w.h == INVALID_HANDLE_VALUE) {
        err = GetLastError();
        return false;
    }
    w.buf.reserve(1u << 20);
    if (g.d->enc == FE_UTF8BOM) {
        const u8 bom[3] = {0xEF, 0xBB, 0xBF};
        w.put(bom, 3);
    }
    if (g.d->enc == FE_UTF16LE || g.d->enc == FE_UTF16BE) {
        const bool be = g.d->enc == FE_UTF16BE;
        const u8 bom[2] = {(u8)(be ? 0xFE : 0xFF), (u8)(be ? 0xFF : 0xFE)};
        w.put(bom, 2);
        Utf8ToUtf16 conv;
        std::vector<u16> out;
        auto emit = [&]() {
            if (be)
                for (u16& u : out) u = (u16)((u << 8) | (u >> 8));
            w.put((const u8*)out.data(), out.size() * 2);
            out.clear();
        };
        g.d->doc.forEachPiece([&](const u8* p, u64 n) {
            while (n && w.ok) {
                u64 k = std::min<u64>(n, 1u << 18);
                conv.feed(p, k, out);
                emit();
                p += k;
                n -= k;
            }
            return w.ok;
        });
        conv.finish(out);
        emit();
    } else {
        g.d->doc.forEachPiece([&](const u8* p, u64 n) {
            w.put(p, n);
            return w.ok;
        });
    }
    w.flush();
    if (w.ok && !FlushFileBuffers(w.h)) {
        w.ok = false;
        w.err = GetLastError();
    }
    CloseHandle(w.h);
    err = w.err;
    return w.ok;
}

bool replaceFile(const std::wstring& tmp, const std::wstring& path, DWORD& err) {
    if (GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES) {
        // Keeps the original file's attributes and permissions.
        if (ReplaceFileW(path.c_str(), tmp.c_str(), nullptr, REPLACEFILE_IGNORE_MERGE_ERRORS, nullptr, nullptr))
            return true;
    }
    if (MoveFileExW(tmp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED)) return true;
    err = GetLastError();
    return false;
}

bool saveTo(const std::wstring& path) {
    HCURSOR oldCur = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    std::wstring tmp = path + L".jn-tmp";
    DWORD err = 0;
    if (!writeDoc(tmp, err)) {
        DeleteFileW(tmp.c_str());
        SetCursor(oldCur);
        errorBox(L"Could not write the file.", path, err);
        return false;
    }
    const bool samePath = !g.d->path.empty() && lstrcmpiW(path.c_str(), g.d->path.c_str()) == 0;
    const bool wasMapped = g.d->back.mapped;
    const bool mustRelease = wasMapped && samePath;
    if (mustRelease) {
        // Windows will not replace a file that is memory-mapped, so let go of it first.
        g.d->index.reset();
        g.d->back.close();
    }
    if (!replaceFile(tmp, path, err)) {
        if (mustRelease) {
            // Put the mapping back; the document itself is untouched.
            Backing nb;
            u64 size = 0;
            DWORD e2 = 0;
            if (loadBacking(g.d->path, nb, size, true, e2) && nb.view && size >= g.d->bomSkip) {
                const u8* text = nb.view + g.d->bomSkip;
                startIndex(text, size - g.d->bomSkip, true);
                g.d->doc.rebaseOriginal(text, g.d->index.get());
                g.d->doc.indexReady();
                g.d->back.take(nb);
                DeleteFileW(tmp.c_str());
                SetCursor(oldCur);
                errorBox(L"Could not replace the file. Your changes are still open and unsaved.", path, err);
            } else {
                // The original is gone as well. The edited text is in the temp file.
                g.d->doc.reset(nullptr, 0, nullptr);
                SetCursor(oldCur);
                errorBox(L"Could not replace the file, and the original could not be reopened.\n"
                         L"Your edited text was written next to it with the extension .jn-tmp.",
                         path, err);
                openPath(tmp, true);
            }
        } else {
            DeleteFileW(tmp.c_str());
            SetCursor(oldCur);
            errorBox(L"Could not replace the file.", path, err);
        }
        refresh(false);
        return false;
    }
    if (wasMapped) {
        // Continue from the freshly written file (undo history starts over).
        // If the old mapping was released, the document must not point at it
        // while the new file is being opened.
        if (mustRelease) g.d->doc.reset(nullptr, 0, nullptr);
        if (!openPath(path, true)) {
            g.d->path = path;
            detectLang();
        }
    } else {
        g.d->path = path;
        detectLang();
        g.d->doc.markSaved();
        startVcsQuery(samePath);
    }
    SetCursor(oldCur);
    refresh(false);
    return true;
}

bool cmdSaveAs() {
    wchar_t file[32768];
    file[0] = 0;
    if (!g.d->path.empty()) lstrcpynW(file, g.d->path.c_str(), 32768);
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g.frame;
    ofn.lpstrFilter = L"All files (*.*)\0*.*\0Text files (*.txt)\0*.txt\0";
    ofn.lpstrFile = file;
    ofn.nMaxFile = 32768;
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR;
    if (!GetSaveFileNameW(&ofn)) return false;
    return saveTo(file);
}

bool cmdSave() {
    if (g.d->path.empty()) return cmdSaveAs();
    if (!g.d->doc.modified()) return true;
    return saveTo(g.d->path);
}

bool confirmDiscard() {
    if (!g.d->doc.modified()) return true;
    std::wstring msg = L"Do you want to save changes to " + fileNameOf(g.d->path) + L"?";
    int r = MessageBoxW(g.frame, msg.c_str(), kAppName, MB_YESNOCANCEL | MB_ICONWARNING);
    if (r == IDYES) return cmdSave();
    return r == IDNO;
}

void cmdOpen() {
    wchar_t file[32768];
    file[0] = 0;
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = g.frame;
    ofn.lpstrFilter = L"All files (*.*)\0*.*\0Text files (*.txt)\0*.txt\0";
    ofn.lpstrFile = file;
    ofn.nMaxFile = 32768;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR;
    if (GetOpenFileNameW(&ofn)) openInTab(file);
}

// ====================================================================== changes and diff view

void layoutFrame();
void setWrap(bool on);

void gotoChange(bool forward) {
    if (!g.d->doc.linesKnown()) return;
    if (g.d->diffDocVersion != g.d->doc.version()) recomputeDiff();
    std::vector<Slot> slots = buildSlots();
    u64 line = g.d->doc.lineOf(g.d->caret);
    int pick = -1;
    if (forward) {
        for (size_t i = 0; i < slots.size(); i++)
            if (slots[i].bStart > line) {
                pick = (int)i;
                break;
            }
    } else {
        for (size_t i = slots.size(); i-- > 0;)
            if (slots[i].bStart < line) {
                pick = (int)i;
                break;
            }
    }
    if (pick < 0) {
        MessageBeep(MB_OK);
        return;
    }
    const Slot& sl = slots[(size_t)pick];
    moveCaret(g.d->doc.lineStart(sl.bStart), false);
    // Bring the change to the middle of the window.
    g.d->top = g.d->doc.lineStart(sl.bStart);
    g.d->topFill = 0;
    scrollRows(-(g.visRows / 2));
    refresh(false);
}

// Puts the base revision's lines back for the change the caret is in.
// Applies one of the three choices to change `hi`.
// The text a slot's lines should hold for a given choice.
std::string choiceText(const DocState::Choice& c, int choice) {
    std::string base = g.d->baseDoc.text(g.d->baseDoc.lineStart(c.aStart),
                                      g.d->baseDoc.lineStart(c.aStart + c.aCount) - g.d->baseDoc.lineStart(c.aStart));
    normalizeEol(base, g.d->crlf);
    const char* eol = g.d->crlf ? "\r\n" : "\n";
    if (choice == kUseDepot) return base;
    if (choice == kKeepYours) return c.yours;
    if (!base.empty() && base.back() != '\n' && !c.yours.empty()) base += eol;
    return base + c.yours;
}

// Applies one of the three choices to slot `si` (as painted). Any choice can
// be changed again later: your own lines are remembered until you submit.
void applyChoice(int si, int choice) {
    if (!g.d->vcs.haveBase || !g.d->doc.linesKnown()) return;
    if (g.d->diffDocVersion != g.d->doc.version()) recomputeDiff();
    std::vector<Slot> slots = buildSlots();
    if (si < 0 || (size_t)si >= slots.size()) return;
    Slot sl = slots[(size_t)si];
    const u64 b0 = g.d->doc.lineStart(sl.bStart), b1 = g.d->doc.lineStart(sl.bStart + sl.bCount);
    const std::string now = g.d->doc.text(b0, b1 - b0);
    if (sl.rec < 0) {
        g.d->choices.push_back({sl.aStart, sl.aCount, now, kKeepYours});
        sl.rec = (int)g.d->choices.size() - 1;
    }
    DocState::Choice& rec = g.d->choices[(size_t)sl.rec];
    // If you edited these lines after choosing, your edit becomes "yours".
    if (now != choiceText(rec, rec.choice)) {
        rec.yours = now;
        rec.choice = kKeepYours;
    }
    if (choice == rec.choice) return;
    std::string t = choiceText(rec, choice);
    const char* eol = g.d->crlf ? "\r\n" : "\n";
    if (!t.empty() && t.back() != '\n' && b1 < g.d->doc.size()) t += eol;
    if (!t.empty() && b0 == g.d->doc.size() && b0 > 0 && g.d->doc.at(b0 - 1) != '\n') t = eol + t;
    rec.choice = choice;
    g.d->doc.breakCoalesce();
    g.d->doc.beginGroup();
    g.d->doc.erase(b0, b1 - b0);
    g.d->doc.insert(b0, (const u8*)t.data(), t.size());
    g.d->doc.endGroup();
    afterEdit(b0);
}

// Ctrl+R: use the depot version of the change at the caret.
// Ctrl+R: use the depot version of the change at the caret.
void revertChange() {
    if (!g.d->vcs.haveBase || !g.d->doc.linesKnown()) return;
    if (g.d->diffDocVersion != g.d->doc.version()) recomputeDiff();
    u32 line = (u32)g.d->doc.lineOf(g.d->caret);
    std::vector<Slot> slots = buildSlots();
    for (size_t i = 0; i < slots.size(); i++) {
        const Slot& sl = slots[i];
        if (line >= sl.bStart && line < sl.bStart + std::max<u32>(sl.bCount, 1)) {
            applyChoice((int)i, kUseDepot);
            return;
        }
    }
    MessageBeep(MB_OK);
}

void toggleDiff() {
    g.diffOn = !g.diffOn;
    if (g.diffOn) {
        // The two sides are matched line by line, so wrapping is off while comparing.
        g.wrapBeforeDiff = g.wrap;
        if (g.wrap) setWrap(false);
    }
    ShowWindow(g.diffPane, g.diffOn ? SW_SHOW : SW_HIDE);
    ShowWindow(g.strip, g.diffOn ? SW_SHOW : SW_HIDE);
    layoutFrame();
    if (!g.diffOn && g.wrapBeforeDiff) setWrap(true);
    refresh(true);
}

// Left pane of the diff view: the base revision, read-only, scrolled so that
// the line matching the middle of the editor sits at the same height.
// ---- left pane: the base revision

void paintDiff(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(hwnd, &ps);
    RECT rc;
    GetClientRect(hwnd, &rc);
    const int W = rc.right, H = rc.bottom;
    if (W <= 0 || H <= 0) {
        EndPaint(hwnd, &ps);
        return;
    }
    HDC dc = CreateCompatibleDC(wdc);
    HBITMAP bmp = CreateCompatibleBitmap(wdc, W, H);
    HGDIOBJ oldBmp = SelectObject(dc, bmp);
    const COLORREF bg = g.pal.bg, fg = g.pal.fg;
    const COLORREF dim = blend(fg, bg, 45);
    fillRect(dc, 0, 0, W, H, bg);
    SetBkMode(dc, TRANSPARENT);

    if (const wchar_t* note = diffNote()) {
        HGDIOBJ of = SelectObject(dc, g.uiFont);
        SetTextColor(dc, dim);
        RECT t = {MulDiv(12, g.dpi, 96), 0, W - MulDiv(12, g.dpi, 96), H};
        RECT m = t;
        DrawTextW(dc, note, -1, &m, DT_CENTER | DT_WORDBREAK | DT_NOPREFIX | DT_CALCRECT);
        t.top += std::max(0, (H - (int)(m.bottom - m.top)) / 2);
        DrawTextW(dc, note, -1, &t, DT_CENTER | DT_WORDBREAK | DT_NOPREFIX);
        SelectObject(dc, of);
    } else {
        HGDIOBJ of = SelectObject(dc, g.font);
        LayoutCfg c = cfg();
        c.wrapCols = 0;
        const u64 totalA = g.d->baseHash.size();
        int digits = 3;
        for (u64 v = totalA; v >= 1000; v /= 10) digits++;
        const int gut = (digits + 2) * g.cw;
        const int cols = std::max(1, (W - gut) / g.cw);
        const u64 top = alignedTop();
        auto yOfRow = [&](u64 row) { return (int)(((i64)row - (i64)top) * g.ch); };
        HlPainter hpBase;
        if (hlOn()) {
            hlSync(g.d->baseDoc, g.d->hlBase);
            hpBase.d = &g.d->baseDoc;
            hpBase.lang = g.d->lang;
            hpBase.cache = &g.d->hlBase;
        }
        // Change backgrounds first (line numbers included), so text draws over them.
        for (const Hunk& h : g.d->hunks) {
            const u64 r0 = changeTopRow(h.aStart, h.aCount, h.bStart, h.bCount);
            const int y0 = yOfRow(r0), yLines = yOfRow(r0 + h.aCount), y1 = yOfRow(r0 + std::max(h.aCount, h.bCount));
            if (y1 < 0 && y0 < 0) continue;
            if (y0 > H) continue;
            const COLORREF col = changeColor(h.aCount, h.bCount);
            fillRect(dc, 0, y0, W, yLines, tintOf(col));       // lines that changed
            fillRect(dc, 0, yLines, W, y1, fillTintOf(col));   // fillers: lines only on the right
            // Borders: the same lines the band in the strip continues.
            const int ew = edgeW();
            fillRect(dc, 0, y0, W, y0 + ew, edgeOf(col));
            if (h.aCount > 0) fillRect(dc, 0, yLines - ew, W, yLines, edgeOf(col));
        }
        // Binary search for the first base line at or below the top row.
        u64 lo = 0, hi = totalA;
        while (lo < hi) {
            u64 mid = (lo + hi) / 2;
            if (alignedA(mid) < top) lo = mid + 1;
            else hi = mid;
        }
        u64 row = top;
        for (u64 line = lo;; line++) {
            const u64 at = line < totalA ? alignedA(line) : alignedA(totalA);
            // Rows before this line that have no line on this side: fillers.
            row = std::max(row, at);  // filler rows were painted with their change above
            if (line >= totalA) break;
            const int y = yOfRow(at);
            if (y >= H) break;
            row = at + 1;
            wchar_t nb[32];
            int k = swprintf(nb, 32, L"%llu", (unsigned long long)line + 1);
            RECT grc = {0, y, gut - g.cw, y + g.ch};
            SetTextColor(dc, dim);
            DrawTextW(dc, nb, k, &grc, DT_RIGHT | DT_SINGLELINE | DT_NOPREFIX | DT_NOCLIP);
            Row r = rowAt(g.d->baseDoc, g.d->baseDoc.lineStart(line), c);
            // What changed within the line, against the matching current line.
            int hi = hunkAtA(g.d->hunks, (u32)line);
            if (hi >= 0) {
                const Hunk& h = g.d->hunks[(size_t)hi];
                if (h.aCount && h.bCount && line - h.aStart < h.bCount) {
                    u64 x0, x1;
                    if (inlineSpan(g.d->baseDoc, r.start, r.end, g.d->doc, h.bStart + (line - h.aStart), x0, x1)) {
                        int c0 = colOf(g.d->baseDoc, r, x0, c), c1 = colOf(g.d->baseDoc, r, x1, c);
                        int px0 = std::max(gut, gut + (c0 - g.d->hcol) * g.cw);
                        int px1 = std::min(W, gut + (c1 - g.d->hcol) * g.cw);
                        const int ew = edgeW();
                        if (px1 > px0) fillRect(dc, px0, y + ew, px1, y + g.ch - ew, inlineOf(changeColor(h.aCount, h.bCount)));
                    }
                }
            }
            SetTextColor(dc, fg);
            drawRowText(dc, g.d->baseDoc, r, gut, W, y, g.d->hcol, cols, c, hlOn() ? hpBase.row(line, r.start, r.end) : nullptr);
        }
        SelectObject(dc, of);
    }
    BitBlt(wdc, 0, 0, W, H, dc, 0, 0, SRCCOPY);
    SelectObject(dc, oldBmp);
    DeleteObject(bmp);
    DeleteDC(dc);
    EndPaint(hwnd, &ps);
}

LRESULT CALLBACK diffProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paintDiff(hwnd); return 0;
    case WM_MOUSEWHEEL:
    case WM_MOUSEHWHEEL: return SendMessageW(g.view, msg, wp, lp);
    case WM_LBUTTONDOWN: SetFocus(g.view); return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ---- middle strip: connecting bands and the three choices per change




void paintStrip(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(hwnd, &ps);
    RECT rc;
    GetClientRect(hwnd, &rc);
    const int W = rc.right, H = rc.bottom;
    g.icons.clear();
    if (W <= 0 || H <= 0) {
        EndPaint(hwnd, &ps);
        return;
    }
    HDC dc = CreateCompatibleDC(wdc);
    HBITMAP bmp = CreateCompatibleBitmap(wdc, W, H);
    HGDIOBJ oldBmp = SelectObject(dc, bmp);
    const COLORREF face = g.pal.face, bg = g.pal.bg,
                   fg = g.pal.fg;
    // Same background as the panes, so the bands flow straight into them.
    fillRect(dc, 0, 0, W, H, bg);
    (void)face;

    if (!diffNote()) {
        const u64 top = alignedTop();
        const int icon = MulDiv(13, g.dpi, 96), gap = MulDiv(3, g.dpi, 96);
        g_slots = buildSlots();
        for (size_t i = 0; i < g_slots.size(); i++) {
            const Slot& sl = g_slots[i];
            const u64 ra = changeTopRow(sl.aStart, sl.aCount, sl.bStart, sl.bCount), rb = ra;
            auto y = [&](u64 row) { return (double)((i64)row - (i64)top) * g.ch; };
            // The band starts and ends exactly where the tinted rows are on
            // each side; a side with no lines is a thin line at its position.
            const double a0 = y(ra), a1 = y(ra + sl.aCount), b0 = y(rb), b1 = y(rb + sl.bCount);
            if (std::max(a1, b1) < -g.ch || std::min(a0, b0) > H + g.ch) continue;
            const bool settled = sl.choice == kUseDepot;
            const COLORREF col = settled ? blend(fg, bg, 45) : sl.choice == kUseBoth ? kAdded
                                                                : changeColor(sl.aCount, sl.bCount);
            drawDiffBand(dc, W, (float)a0, (float)a1, (float)b0, (float)b1, settled ? fillTintOf(col) : tintOf(col),
                         edgeOf(col), (float)edgeW());

            // The three choices as 3D buttons, centred on the band. The one in
            // effect stays pushed in, like a toggle.
            const int bs = MulDiv(18, g.dpi, 96), bgap = MulDiv(4, g.dpi, 96);
            int cy = (int)std::lround((std::min(a0, b0) + std::max(a1, b1)) / 2);
            if (cy < bs || cy > H - bs) continue;
            const int total = 3 * bs + 2 * bgap;
            int x = (W - total) / 2;
            const int order[3] = {kUseDepot, kUseBoth, kKeepYours};
            const Glyph glyphs[3] = {GLYPH_LEFT, GLYPH_BOTH, GLYPH_RIGHT};
            const COLORREF tints[3] = {kDepotCol, kAdded, kYoursCol};
            for (int k = 0; k < 3; k++) {
                const int choice = order[k];
                RECT r = {x, cy - bs / 2, x + bs, cy - bs / 2 + bs};
                const bool hot = g.hoverHunk == (int)i && g.hoverChoice == choice;
                const bool held = hot && g.pressHunk == (int)i && g.pressChoice == choice;
                const bool on = sl.choice == choice;
                ButtonState st = held ? BTN_PRESSED : on ? (hot ? BTN_LATCHED_HOT : BTN_LATCHED) : hot ? BTN_HOT : BTN_NORMAL;
                drawButton3D(dc, r, st, g.pal, tints[k]);
                drawGlyph(dc, r, glyphs[k], choice == kKeepYours ? kYoursCol : kDepotCol, kYoursCol,
                          held || on);
                g.icons.push_back({r, (int)i, choice});
                x += bs + bgap;
            }
        }
    }
    BitBlt(wdc, 0, 0, W, H, dc, 0, 0, SRCCOPY);
    SelectObject(dc, oldBmp);
    DeleteObject(bmp);
    DeleteDC(dc);
    EndPaint(hwnd, &ps);
}

const wchar_t* choiceHint(int choice) {
    switch (choice) {
    case kUseDepot: return L"Use the depot version here (you can switch back until you submit)";
    case kUseBoth: return L"Use both: the depot lines followed by yours";
    default: return L"Use your version here";
    }
}

void applyChoice(int hi, int choice);

void setHover(int hunk, int choice) {
    if (hunk == g.hoverHunk && choice == g.hoverChoice) return;
    g.hoverHunk = hunk;
    g.hoverChoice = choice;
    g.hint = hunk >= 0 ? choiceHint(choice) : L"";
    InvalidateRect(g.strip, nullptr, FALSE);
    updateStatus();
}

LRESULT CALLBACK stripProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    auto hit = [&](POINT pt, int& hunk, int& choice) {
        hunk = choice = -1;
        for (const auto& ic : g.icons)
            if (PtInRect(&ic.rc, pt)) {
                hunk = ic.hunk;
                choice = ic.choice;
            }
    };
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paintStrip(hwnd); return 0;
    case WM_MOUSEWHEEL:
    case WM_MOUSEHWHEEL: return SendMessageW(g.view, msg, wp, lp);
    case WM_MOUSEMOVE: {
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        int hunk, choice;
        hit(pt, hunk, choice);
        setHover(hunk, choice);
        TRACKMOUSEEVENT tme = {sizeof(tme), TME_LEAVE, hwnd, 0};
        TrackMouseEvent(&tme);
        return 0;
    }
    case WM_MOUSELEAVE:
        if (GetCapture() != hwnd) setHover(-1, -1);
        return 0;
    case WM_SETCURSOR:
        if (g.hoverHunk >= 0) {
            SetCursor(LoadCursorW(nullptr, IDC_HAND));
            return TRUE;
        }
        break;
    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK: {
        // Pushed in while held; the choice applies on release, over the same button.
        if (g.hoverHunk >= 0) {
            g.pressHunk = g.hoverHunk;
            g.pressChoice = g.hoverChoice;
            SetCapture(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
        } else {
            SetFocus(g.view);
        }
        return 0;
    }
    case WM_LBUTTONUP: {
        if (GetCapture() != hwnd) return 0;
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        int hunk, choice;
        hit(pt, hunk, choice);
        const int ph = g.pressHunk, pc = g.pressChoice;
        g.pressHunk = g.pressChoice = -1;
        ReleaseCapture();
        SetFocus(g.view);
        if (hunk == ph && choice == pc && ph >= 0) applyChoice(ph, pc);
        InvalidateRect(hwnd, nullptr, FALSE);
        return 0;
    }
    case WM_CAPTURECHANGED:
        if (g.pressHunk >= 0) {
            g.pressHunk = g.pressChoice = -1;
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ====================================================================== version control actions

bool cmdSave();

std::wstring quoted(const std::wstring& s) { return L"\"" + s + L"\""; }

// Runs a git / p4 command in the file's folder and reports failures.
bool runVcs(const std::wstring& cmd) {
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    RunResult r = runProcess(cmd, dirOf(g.d->path), 30000, true, 1u << 20);
    SetCursor(old);
    if (!r.started) {
        std::wstring msg = L"Could not run:\n\n" + cmd + L"\n\nMake sure the tool is installed and on PATH.";
        MessageBoxW(g.frame, msg.c_str(), kAppName, MB_OK | MB_ICONERROR);
        return false;
    }
    if (r.timedOut || r.exitCode != 0) {
        std::wstring msg = cmd + L"\n\n" + (r.timedOut ? L"The command timed out.\n\n" : L"") + widenUtf8(r.out);
        MessageBoxW(g.frame, msg.c_str(), kAppName, MB_OK | MB_ICONERROR);
        return false;
    }
    return true;
}

bool vcsReady(bool mustBeSaved) {
    if (g.d->vcs.kind == VcsKind::None || g.d->vcs.toolMissing || g.d->path.empty()) return false;
    if (mustBeSaved && g.d->doc.modified() && !cmdSave()) return false;
    return true;
}

// Git: stage the file. Perforce: check it out.
void vcsPrepare() {
    const bool git = g.d->vcs.kind == VcsKind::Git;
    if (!vcsReady(git)) return;
    bool ok = git ? runVcs(gitCmd() + L" add -- " + quoted(nameOf(g.d->path))) : runVcs(p4Cmd() + L" edit " + quoted(g.d->path));
    if (ok) startVcsQuery(true);
}

// What the submit dialog shows and returns.
struct SubmitDlg {
    std::wstring title, message;
    std::vector<std::wstring> files;  // as shown
    std::vector<bool> checked;        // in / out, parallel to files
};

INT_PTR CALLBACK commitProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) {
    if (INT_PTR br = dialogTheme(dlg, msg, wp, lp)) return br;
    switch (msg) {
    case WM_INITDIALOG: {
        SetWindowLongPtrW(dlg, DWLP_USER, lp);
        SubmitDlg* d = (SubmitDlg*)lp;
        SetWindowTextW(dlg, d->title.c_str());
        HWND lv = GetDlgItem(dlg, IDC_FILES);
        ListView_SetExtendedListViewStyle(lv, LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);
        RECT rc;
        GetClientRect(lv, &rc);
        LVCOLUMNW col = {};
        col.mask = LVCF_WIDTH;
        col.cx = rc.right - GetSystemMetrics(SM_CXVSCROLL);
        ListView_InsertColumn(lv, 0, &col);
        for (size_t i = 0; i < d->files.size(); i++) {
            LVITEMW it = {};
            it.mask = LVIF_TEXT;
            it.iItem = (int)i;
            it.pszText = (LPWSTR)d->files[i].c_str();
            ListView_InsertItem(lv, &it);
            ListView_SetCheckState(lv, (int)i, d->checked[i] ? TRUE : FALSE);
        }
        SetDlgItemTextW(dlg, IDC_MSG, d->message.c_str());
        if (d->title.rfind(L"Commit", 0) == 0) SetDlgItemTextW(dlg, IDOK, L"&Commit");
        SetFocus(GetDlgItem(dlg, IDC_MSG));
        SendDlgItemMessageW(dlg, IDC_MSG, EM_SETSEL, 0, -1);
        return FALSE;
    }
    case WM_COMMAND:
        if (LOWORD(wp) == IDOK) {
            SubmitDlg* d = (SubmitDlg*)GetWindowLongPtrW(dlg, DWLP_USER);
            HWND lv = GetDlgItem(dlg, IDC_FILES);
            bool any = false;
            for (size_t i = 0; i < d->files.size(); i++) {
                d->checked[i] = ListView_GetCheckState(lv, (int)i) != 0;
                any = any || d->checked[i];
            }
            if (!any) {
                MessageBoxW(dlg, L"Tick at least one file to submit.", kAppName, MB_OK | MB_ICONINFORMATION);
                return TRUE;
            }
            int len = GetWindowTextLengthW(GetDlgItem(dlg, IDC_MSG));
            std::wstring m((size_t)len + 1, L'\0');
            GetDlgItemTextW(dlg, IDC_MSG, &m[0], len + 1);
            m.resize((size_t)len);
            if (m.find_first_not_of(L" \t\r\n") == std::wstring::npos) {
                MessageBoxW(dlg, L"Please enter a description.", kAppName, MB_OK | MB_ICONINFORMATION);
                return TRUE;
            }
            d->message = m;
            EndDialog(dlg, 1);
            return TRUE;
        }
        if (LOWORD(wp) == IDCANCEL) {
            EndDialog(dlg, 0);
            return TRUE;
        }
        break;
    }
    return FALSE;
}

// Runs a command quietly and returns its output, or "" on failure.
std::string vcsOutput(const std::wstring& cmd) {
    RunResult r = runProcess(cmd, dirOf(g.d->path), 20000, false, 1u << 20);
    return r.started && !r.timedOut && r.exitCode == 0 ? r.out : std::string();
}

std::wstring crlfText(const std::string& s) {
    std::wstring w = widenUtf8(s), out;
    for (wchar_t c : w) {
        if (c == L'\r') continue;
        if (c == L'\n') out += L'\r';
        out += c;
    }
    while (!out.empty() && (out.back() == L'\n' || out.back() == L'\r')) out.pop_back();
    return out;
}

std::string toUtf8(const std::wstring& w) {
    int k = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string s((size_t)std::max(k, 0), '\0');
    if (k > 0) WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), &s[0], k, nullptr, nullptr);
    return s;
}

// The Description block of a `p4 change -o` spec, as plain text.
std::string specDescription(const std::string& spec) {
    std::string out;
    size_t p = spec.find("\nDescription:");
    if (p == std::string::npos) return out;
    p = spec.find('\n', p + 1);
    while (p != std::string::npos && p + 1 < spec.size() && spec[p + 1] == '\t') {
        size_t e = spec.find('\n', p + 1);
        std::string line = spec.substr(p + 2, (e == std::string::npos ? spec.size() : e) - p - 2);
        if (!line.empty() && line.back() == '\r') line.pop_back();
        out += line + "\n";
        p = e;
    }
    while (!out.empty() && out.back() == '\n') out.pop_back();
    return out;
}

// The same spec with its Description replaced.
std::string specWithDescription(const std::string& spec, const std::wstring& desc) {
    std::string body = "Description:\n";
    std::string d = toUtf8(desc);
    size_t i = 0;
    while (i <= d.size()) {
        size_t e = d.find('\n', i);
        std::string line = d.substr(i, (e == std::string::npos ? d.size() : e) - i);
        if (!line.empty() && line.back() == '\r') line.pop_back();
        body += "\t" + line + "\n";
        if (e == std::string::npos) break;
        i = e + 1;
    }
    size_t p = spec.find("\nDescription:");
    if (p == std::string::npos) return spec;
    size_t q = spec.find('\n', p + 1);
    while (q != std::string::npos && q + 1 < spec.size() && spec[q + 1] == '\t') q = spec.find('\n', q + 1);
    return spec.substr(0, p + 1) + body + (q == std::string::npos ? "" : spec.substr(q + 1));
}

// The file lines ("\t//depot/a.txt\t# edit") of a change spec's Files block.
std::vector<std::string> specFiles(const std::string& spec) {
    std::vector<std::string> v;
    size_t p = spec.find("\nFiles:");
    if (p == std::string::npos) return v;
    p = spec.find('\n', p + 1);
    while (p != std::string::npos && p + 1 < spec.size() && spec[p + 1] == '\t') {
        size_t e = spec.find('\n', p + 1);
        std::string line = spec.substr(p + 1, (e == std::string::npos ? spec.size() : e) - p - 1);
        if (!line.empty() && line.back() == '\r') line.pop_back();
        v.push_back(line);
        p = e;
    }
    return v;
}

// The same spec with only the given file lines.
std::string specWithFiles(const std::string& spec, const std::vector<std::string>& keep) {
    size_t p = spec.find("\nFiles:");
    if (p == std::string::npos) return spec;
    size_t q = spec.find('\n', p + 1);
    while (q != std::string::npos && q + 1 < spec.size() && spec[q + 1] == '\t') q = spec.find('\n', q + 1);
    std::string body = "Files:\n";
    for (const std::string& f : keep) body += f + "\n";
    return spec.substr(0, p + 1) + body + (q == std::string::npos ? "" : spec.substr(q + 1));
}

// "\t//depot/a.txt\t# edit" -> "//depot/a.txt"
std::string specFilePath(const std::string& line) {
    size_t a = line.find_first_not_of('\t');
    size_t b = line.find('\t', a == std::string::npos ? 0 : a);
    return a == std::string::npos ? "" : line.substr(a, (b == std::string::npos ? line.size() : b) - a);
}

// Git: top folder of the repository holding the open file.
std::wstring gitTop() {
    std::string o = vcsOutput(gitCmd() + L" rev-parse --show-toplevel");
    while (!o.empty() && (o.back() == '\n' || o.back() == '\r')) o.pop_back();
    std::wstring w = widenUtf8(o);
    for (wchar_t& c : w)
        if (c == L'/') c = L'\\';
    return w;
}

// Git: changed tracked files as "XY path" lines and their paths.
void gitChanges(std::vector<std::wstring>& shown, std::vector<std::wstring>& paths, bool untracked) {
    std::string st = vcsOutput(gitCmd() + L" status --porcelain" + (untracked ? L"" : L" --untracked-files=no"));
    size_t pos = 0;
    while (pos < st.size()) {
        size_t e = st.find('\n', pos);
        if (e == std::string::npos) e = st.size();
        std::string line = st.substr(pos, e - pos);
        pos = e + 1;
        if (!line.empty() && line.back() == '\r') line.pop_back();
        if (line.size() < 4) continue;
        std::string path = line.substr(3);
        size_t arrow = path.find(" -> ");
        if (arrow != std::string::npos) path = path.substr(arrow + 4);
        if (path.size() >= 2 && path.front() == '"' && path.back() == '"') path = path.substr(1, path.size() - 2);
        shown.push_back(widenUtf8(line));
        paths.push_back(widenUtf8(path));
    }
}

// Perforce: submit the changelist this file is open in (every file in it).
// Git: commit every changed tracked file in the repository.
void vcsCommit() {
    const bool git = g.d->vcs.kind == VcsKind::Git;
    if (!vcsReady(true)) return;
    SubmitDlg d;
    if (git) {
        d.title = L"Commit Changes";
        std::vector<std::wstring> paths;
        gitChanges(d.files, paths, false);
        if (d.files.empty()) {
            MessageBoxW(g.frame, L"There are no changes to commit.", kAppName, MB_OK | MB_ICONINFORMATION);
            return;
        }
        d.checked.assign(d.files.size(), true);
        if (DialogBoxParamW(g.inst, MAKEINTRESOURCEW(IDD_COMMIT), g.frame, commitProc, (LPARAM)&d) != 1) return;
        // The message goes through a temp file so quotes and newlines survive.
        wchar_t dir[MAX_PATH], file[MAX_PATH];
        if (!GetTempPathW(MAX_PATH, dir) || !GetTempFileNameW(dir, L"jn", 0, file)) return;
        std::string bytes = toUtf8(d.message);
        HANDLE h = CreateFileW(file, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE) return;
        DWORD put = 0;
        WriteFile(h, bytes.data(), (DWORD)bytes.size(), &put, nullptr);
        CloseHandle(h);
        std::wstring cmd = gitCmd() + L" commit -F " + quoted(file) + L" --";
        for (size_t i = 0; i < paths.size(); i++)
            if (d.checked[i]) cmd += L" " + quoted(paths[i]);
        std::wstring top = gitTop();
        HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
        RunResult r = runProcess(cmd, top.empty() ? dirOf(g.d->path) : top, 30000, true, 1u << 20);
        SetCursor(old);
        DeleteFileW(file);
        if (!r.started || r.timedOut || r.exitCode != 0) {
            MessageBoxW(g.frame, (L"Commit failed.\n\n" + widenUtf8(r.out)).c_str(), kAppName, MB_OK | MB_ICONERROR);
            return;
        }
        g.d->choices.clear();
        startVcsQuery(true);
        return;
    }

    const std::wstring change = g.d->vcs.change.empty() ? L"default" : g.d->vcs.change;
    const bool numbered = change != L"default";
    std::string spec = vcsOutput(p4Cmd() + L" change -o" + (numbered ? L" " + change : L""));
    if (spec.empty()) {
        runVcs(p4Cmd() + L" change -o " + change);  // shows the error
        return;
    }
    std::vector<std::string> lines = specFiles(spec);
    if (lines.empty()) {
        MessageBoxW(g.frame, L"This changelist has no files.", kAppName, MB_OK | MB_ICONINFORMATION);
        return;
    }
    d.title = numbered ? L"Submit Changelist " + change : L"Submit Default Changelist";
    for (const std::string& l : lines) {
        std::string path = specFilePath(l);
        size_t hash = l.find("# ");
        std::string action = hash == std::string::npos ? "" : l.substr(hash + 2);
        d.files.push_back(widenUtf8(path) + (action.empty() ? L"" : L"   (" + widenUtf8(action) + L")"));
    }
    d.checked.assign(lines.size(), true);
    if (numbered) {
        std::string desc = specDescription(spec);
        if (desc != "<enter description here>") d.message = crlfText(desc);
    }
    if (DialogBoxParamW(g.inst, MAKEINTRESOURCEW(IDD_COMMIT), g.frame, commitProc, (LPARAM)&d) != 1) return;

    std::vector<std::string> keep, leave;
    for (size_t i = 0; i < lines.size(); i++) (d.checked[i] ? keep : leave).push_back(lines[i]);
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    // Files left out of a numbered changelist go back to the default one, so
    // they stay checked out but are not submitted.
    if (numbered && !leave.empty()) {
        std::wstring cmd = p4Cmd() + L" reopen -c default";
        for (const std::string& l : leave) cmd += L" " + quoted(widenUtf8(specFilePath(l)));
        if (!runVcs(cmd)) {
            SetCursor(old);
            return;
        }
        spec = vcsOutput(p4Cmd() + L" change -o " + change);
    }
    // Write the description (and, for the default changelist, the chosen
    // files) through `p4 change -i`; quotes and newlines survive that way.
    std::string newSpec = specWithDescription(numbered ? spec : specWithFiles(spec, keep), d.message);
    RunResult r = runProcess(p4Cmd() + L" change -i", dirOf(g.d->path), 30000, true, 1u << 20, newSpec);
    SetCursor(old);
    if (!r.started || r.timedOut || r.exitCode != 0) {
        MessageBoxW(g.frame, (L"Could not save the changelist.\n\n" + widenUtf8(r.out)).c_str(), kAppName,
                    MB_OK | MB_ICONERROR);
        return;
    }
    std::wstring num = change;
    if (!numbered) {  // "Change 1234 created with 2 open file(s)."
        num.clear();
        size_t p = r.out.find("Change ");
        for (size_t i = p == std::string::npos ? r.out.size() : p + 7; i < r.out.size() && isdigit((unsigned char)r.out[i]); i++)
            num += (wchar_t)r.out[i];
        if (num.empty()) {
            MessageBoxW(g.frame, (L"Unexpected answer from p4 change:\n\n" + widenUtf8(r.out)).c_str(), kAppName,
                        MB_OK | MB_ICONERROR);
            return;
        }
    }
    if (!runVcs(p4Cmd() + L" submit -c " + num)) return;
    g.d->choices.clear();
    // Submitting makes the file read-only again and may rewrite keywords.
    openPath(g.d->path, true);
}

// Lists every file in the changelist of the open file (Git: every changed
// file) so one can be opened here, or all of them at once.
// Opens every file in the changelist of the open file (Git: every changed
// file) as tabs.
void cmdChangelistFiles() {
    if (g.d->vcs.kind == VcsKind::None || g.d->vcs.toolMissing || g.d->path.empty()) return;
    std::vector<std::wstring> paths;
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    if (g.d->vcs.kind == VcsKind::Git) {
        std::vector<std::wstring> shown, rel;
        gitChanges(shown, rel, true);
        std::wstring top = gitTop();
        for (std::wstring& p : rel) {
            for (wchar_t& c : p)
                if (c == L'/') c = L'\\';
            paths.push_back(top + L"\\" + p);
        }
    } else {
        const std::wstring change = g.d->vcs.change.empty() ? L"default" : g.d->vcs.change;
        std::vector<std::string> depot = ztagAll(vcsOutput(p4Cmd() + L" -ztag opened -c " + change), "depotFile");
        if (!depot.empty()) {
            std::wstring cmd = p4Cmd() + L" -ztag where";
            for (const std::string& f : depot) cmd += L" " + quoted(widenUtf8(f));
            for (const std::string& p : ztagAll(vcsOutput(cmd), "path")) paths.push_back(widenUtf8(p));
        }
    }
    SetCursor(old);
    if (paths.empty()) {
        MessageBoxW(g.frame, L"There are no other files in this changelist.", kAppName, MB_OK | MB_ICONINFORMATION);
        return;
    }
    DocState* start = g.d;
    for (const std::wstring& p : paths)
        if (GetFileAttributesW(p.c_str()) != INVALID_FILE_ATTRIBUTES) openInTab(p);
    activateTab(tabIndex(start));  // stay on the file you were in
}

void vcsRevertFile() {
    const bool git = g.d->vcs.kind == VcsKind::Git;
    if (!vcsReady(false)) return;
    std::wstring msg = L"Revert " + fileNameOf(g.d->path) + L"?\n\nEvery change to it, saved or not, is discarded and the " +
                       (git ? L"committed version" : L"depot version") + L" is restored" +
                       (git ? L"." : L" (the file is no longer checked out).") + L"\n\nThis cannot be undone.";
    if (MessageBoxW(g.frame, msg.c_str(), kAppName, MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2) != IDYES) return;
    bool ok = git ? runVcs(gitCmd() + L" checkout HEAD -- " + quoted(nameOf(g.d->path))) : runVcs(p4Cmd() + L" revert " + quoted(g.d->path));
    if (ok) {
        g.d->choices.clear();
        openPath(g.d->path, true);
    }
}

void onP4Ask() {
    if (g.d->vcs.kind != VcsKind::Perforce || g.d->vcs.opened || g.d->vcs.toolMissing || g.d->checkingOut) return;
    if (!vcsConnection().autoCheckout) {
        int r = MessageBoxW(g.frame, L"This file is not checked out from Perforce.\n\nCheck it out now?", kAppName,
                            MB_YESNO | MB_ICONQUESTION);
        if (r == IDYES && runVcs(p4Cmd() + L" edit " + quoted(g.d->path))) startVcsQuery(true);
        return;
    }
    // Automatic: `p4 edit` runs in the background so typing never waits for
    // the server. The answer arrives as WM_APP_P4EDITED.
    g.d->checkingOut = true;
    updateStatus();
    std::wstring cmd = p4Cmd() + L" edit " + quoted(g.d->path), dir = dirOf(g.d->path);
    unsigned gen = g.d->vcsGen;
    HWND frame = g.frame;
    std::thread([cmd, dir, gen, frame]() {
        RunResult* r = new RunResult(runProcess(cmd, dir, 30000, true, 1u << 20));
        if (!PostMessageW(frame, WM_APP_P4EDITED, (WPARAM)gen, (LPARAM)r)) delete r;
    }).detach();
}

void onP4Edited(unsigned gen, RunResult* raw) {
    std::unique_ptr<RunResult> r(raw);
    DocState* t = tabByVcsGen(gen);
    if (!t) return;  // closed, or a different file is open in it now
    t->checkingOut = false;
    DocState* active = g.d;
    g.d = t;
    if (!r->started || r->timedOut || r->exitCode != 0) {
        std::wstring msg = L"Could not check out " + fileNameOf(g.d->path) + L" from Perforce.\n\n" +
                           (!r->started ? std::wstring(L"p4.exe could not be started.") : r->timedOut ? std::wstring(L"The command timed out.") : widenUtf8(r->out)) +
                           L"\n\nUse Source Control > Refresh Status to try again.";
        g.d = active;
        MessageBoxW(g.frame, msg.c_str(), kAppName, MB_OK | MB_ICONWARNING);
        updateStatus();
        return;
    }
    startVcsQuery(true);
    g.d = active;
}
// ---- connection dialog

void initConnectDialog(HWND dlg, const VcsConnection& c) {
    HWND kind = GetDlgItem(dlg, IDC_VCS_KIND);
    SendMessageW(kind, CB_ADDSTRING, 0, (LPARAM)L"Detect automatically (Git, then Perforce)");
    SendMessageW(kind, CB_ADDSTRING, 0, (LPARAM)L"Git only");
    SendMessageW(kind, CB_ADDSTRING, 0, (LPARAM)L"Perforce only");
    SendMessageW(kind, CB_SETCURSEL, c.only == VcsKind::Git ? 1 : c.only == VcsKind::Perforce ? 2 : 0, 0);
    SetDlgItemTextW(dlg, IDC_GIT_EXE, c.gitExe.c_str());
    SetDlgItemTextW(dlg, IDC_P4_EXE, c.p4Exe.c_str());
    SetDlgItemTextW(dlg, IDC_P4_PORT, c.p4Port.c_str());
    SetDlgItemTextW(dlg, IDC_P4_USER, c.p4User.c_str());
    SetDlgItemTextW(dlg, IDC_P4_CLIENT, c.p4Client.c_str());
    // Empty means "detect": show what was found as the field's grey hint.
    std::wstring git = findGit(), p4 = findP4();
    std::wstring gitHint = git.empty() ? L"not found - enter the path" : L"detected: " + git;
    std::wstring p4Hint = p4.empty() ? L"not found - enter the path" : L"detected: " + p4;
    SendDlgItemMessageW(dlg, IDC_GIT_EXE, EM_SETCUEBANNER, TRUE, (LPARAM)gitHint.c_str());
    SendDlgItemMessageW(dlg, IDC_P4_EXE, EM_SETCUEBANNER, TRUE, (LPARAM)p4Hint.c_str());
    CheckDlgButton(dlg, IDC_AUTOCHECKOUT, c.autoCheckout ? BST_CHECKED : BST_UNCHECKED);
}

std::wstring dlgText(HWND dlg, int id) {
    int len = GetWindowTextLengthW(GetDlgItem(dlg, id));
    std::wstring s((size_t)len + 1, L'\0');
    GetDlgItemTextW(dlg, id, &s[0], len + 1);
    s.resize((size_t)len);
    size_t a = s.find_first_not_of(L" \t"), b = s.find_last_not_of(L" \t");
    return a == std::wstring::npos ? L"" : s.substr(a, b - a + 1);
}

VcsConnection readConnectDialog(HWND dlg) {
    VcsConnection c;
    LRESULT k = SendDlgItemMessageW(dlg, IDC_VCS_KIND, CB_GETCURSEL, 0, 0);
    c.only = k == 1 ? VcsKind::Git : k == 2 ? VcsKind::Perforce : VcsKind::None;
    c.gitExe = dlgText(dlg, IDC_GIT_EXE);
    c.p4Exe = dlgText(dlg, IDC_P4_EXE);
    c.p4Port = dlgText(dlg, IDC_P4_PORT);
    c.p4User = dlgText(dlg, IDC_P4_USER);
    c.p4Client = dlgText(dlg, IDC_P4_CLIENT);
    c.autoCheckout = IsDlgButtonChecked(dlg, IDC_AUTOCHECKOUT) == BST_CHECKED;
    return c;
}

// Tries the settings in the dialog without keeping them.
void testConnection(HWND dlg) {
    VcsConnection saved = vcsConnection(), c = readConnectDialog(dlg);
    setVcsConnection(c);
    std::wstring dir = g.d->path.empty() ? L"" : dirOf(g.d->path);
    std::wstring report;
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    if (c.only != VcsKind::Perforce) {
        RunResult r = runProcess(gitCmd() + L" --version", dir, 5000, true, 1u << 16);
        std::wstring exe = c.gitExe.empty() ? findGit() : c.gitExe;
        report += L"Git: " + (exe.empty() ? std::wstring(L"git.exe not found") : exe) + L"\n" +
                  (r.started && r.exitCode == 0 ? widenUtf8(r.out) : std::wstring(L"not available\n"));
    }
    if (c.only != VcsKind::Git) {
        RunResult r = runProcess(p4Cmd() + L" info", dir, 10000, true, 1u << 16);
        std::wstring exe = c.p4Exe.empty() ? findP4() : c.p4Exe;
        report += L"\nPerforce: " + (exe.empty() ? std::wstring(L"p4.exe not found") : exe) + L"\n";
        if (!r.started) report += L"p4.exe not available";
        else if (r.timedOut) report += L"the server did not answer";
        else {
            std::wstring out = widenUtf8(r.out);
            if (out.size() > 1500) out.resize(1500);
            report += (r.exitCode == 0 ? L"connected\n\n" : L"failed\n\n") + out;
        }
    }
    SetCursor(old);
    setVcsConnection(saved);
    MessageBoxW(dlg, report.c_str(), L"Test Connection", MB_OK | MB_ICONINFORMATION);
}

INT_PTR CALLBACK connectProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) {
    if (INT_PTR br = dialogTheme(dlg, msg, wp, lp)) return br;
    switch (msg) {
    case WM_INITDIALOG: initConnectDialog(dlg, vcsConnection()); return TRUE;
    case WM_COMMAND:
        switch (LOWORD(wp)) {
        case IDC_VCS_TEST: testConnection(dlg); return TRUE;
        case IDOK:
            setVcsConnection(readConnectDialog(dlg));
            EndDialog(dlg, 1);
            return TRUE;
        case IDCANCEL: EndDialog(dlg, 0); return TRUE;
        }
        break;
    }
    return FALSE;
}

void saveSettings();

void cmdConnect() {
    if (DialogBoxParamW(g.inst, MAKEINTRESOURCEW(IDD_VCSCONNECT), g.frame, connectProc, 0) != 1) return;
    saveSettings();
    startVcsQuery(false);  // look at the open file again with the new settings
}

// ---- updates (same flow as v1: Help > Check for Updates, GitHub latest release)

void cmdCheckUpdates() {
    if (Updater_IsBusy()) {
        MessageBoxW(g.frame, L"An update check is already running.", kAppName, MB_OK | MB_ICONINFORMATION);
        return;
    }
    SetCursor(LoadCursorW(nullptr, IDC_APPSTARTING));
    Updater_CheckForUpdates(g.frame, false);
}

bool confirmDiscard();

void onUpdateReady(wchar_t* zip) {
    if (!zip) {
        Updater_Reset();
        return;
    }
    int r = MessageBoxW(g.frame,
                        L"A new version has been downloaded.\r\n\r\n"
                        L"JustNotepad will close, install the update and start again.\r\n"
                        L"Continue?",
                        L"Update Ready", MB_OKCANCEL | MB_ICONQUESTION);
    if (r == IDOK && confirmDiscard()) {
        if (Updater_LaunchApplier(zip, g.frame)) {
            Updater_FreeString(zip);
            saveSettings();
            DestroyWindow(g.frame);
            return;
        }
        MessageBoxW(g.frame, L"The update could not be started.", kAppName, MB_OK | MB_ICONWARNING);
    }
    Updater_FreeString(zip);
    Updater_Reset();
}

// ====================================================================== find / replace

bool isWordByte(u8 b) {
    return b >= 0x80 || (b >= '0' && b <= '9') || (b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z') || b == '_';
}

bool wholeWordAt(u64 pos, u64 n) {
    if (pos > 0 && isWordByte(g.d->doc.at(pos - 1))) return false;
    if (pos + n < g.d->doc.size() && isWordByte(g.d->doc.at(pos + n))) return false;
    return true;
}

bool searchProgress(u64) { return (GetAsyncKeyState(VK_ESCAPE) & 0x8000) == 0; }

u64 findFrom(const std::string& nd, u64 from, bool down, bool matchCase, bool word) {
    const u64 size = g.d->doc.size();
    for (;;) {
        u64 hit = down ? g.d->doc.searchFwd(nd, from, size, matchCase, searchProgress)
                       : g.d->doc.searchBack(nd, 0, from, matchCase, searchProgress);
        if (hit == NPOS || !word || wholeWordAt(hit, nd.size())) return hit;
        if (down) {
            from = hit + 1;
        } else {
            if (hit + nd.size() <= 1) return NPOS;
            from = hit + nd.size() - 1;
        }
    }
}

bool findNext(bool down) {
    std::string nd = toDocBytes(g.findBuf, lstrlenW(g.findBuf));
    if (nd.empty()) return false;
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    u64 hit = findFrom(nd, down ? selMax() : selMin(), down, g.findCase, g.findWord);
    if (hit == NPOS) hit = findFrom(nd, down ? 0 : g.d->doc.size(), down, g.findCase, g.findWord);  // wrap around
    SetCursor(old);
    if (hit == NPOS) {
        std::wstring msg = std::wstring(L"Cannot find \"") + g.findBuf + L"\"";
        MessageBoxW(g.findDlg ? g.findDlg : g.frame, msg.c_str(), kAppName, MB_OK | MB_ICONINFORMATION);
        return false;
    }
    g.d->anchor = hit;
    g.d->caret = hit + nd.size();
    g.d->wantCol = -1;
    g.d->doc.breakCoalesce();
    refresh(true);
    return true;
}

bool selectionMatches(const std::string& nd) {
    if (!hasSel() || selMax() - selMin() != nd.size()) return false;
    return g.d->doc.searchFwd(nd, selMin(), selMax(), g.findCase) == selMin();
}

void replaceOne() {
    std::string nd = toDocBytes(g.findBuf, lstrlenW(g.findBuf));
    std::string rp = toDocBytes(g.replBuf, lstrlenW(g.replBuf));
    if (nd.empty()) return;
    if (selectionMatches(nd)) {
        g.d->doc.breakCoalesce();
        replaceSel((const u8*)rp.data(), rp.size(), false);
    }
    findNext(true);
}

void replaceAll() {
    std::string nd = toDocBytes(g.findBuf, lstrlenW(g.findBuf));
    std::string rp = toDocBytes(g.replBuf, lstrlenW(g.replBuf));
    if (nd.empty()) return;
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    u64 pos = 0, count = 0;
    g.d->doc.beginGroup();
    for (;;) {
        u64 hit = findFrom(nd, pos, true, g.findCase, g.findWord);
        if (hit == NPOS) break;
        g.d->doc.erase(hit, nd.size());
        g.d->doc.insert(hit, (const u8*)rp.data(), rp.size());
        pos = hit + rp.size();
        count++;
    }
    g.d->doc.endGroup();
    SetCursor(old);
    afterEdit(std::min(g.d->caret, g.d->doc.size()));
    wchar_t b[96];
    swprintf(b, 96, L"Replaced %llu occurrence(s).", (unsigned long long)count);
    MessageBoxW(g.findDlg ? g.findDlg : g.frame, b, kAppName, MB_OK | MB_ICONINFORMATION);
}

void showFind(bool replace) {
    if (g.findDlg) {
        // Switching between Find and Replace needs a new dialog.
        DestroyWindow(g.findDlg);
        g.findDlg = nullptr;
    }
    // Seed with the selection when it is short and on one line.
    if (hasSel() && selMax() - selMin() < 200) {
        std::string s = g.d->doc.text(selMin(), selMax() - selMin());
        if (s.find('\n') == std::string::npos && s.find('\r') == std::string::npos) {
            std::wstring w = fromDocBytes(s);
            lstrcpynW(g.findBuf, w.c_str(), 256);
        }
    }
    g.fr = FINDREPLACEW{};
    g.fr.lStructSize = sizeof(g.fr);
    g.fr.hwndOwner = g.frame;
    g.fr.Flags = FR_DOWN | (g.findCase ? FR_MATCHCASE : 0) | (g.findWord ? FR_WHOLEWORD : 0);
    g.fr.lpstrFindWhat = g.findBuf;
    g.fr.wFindWhatLen = (WORD)sizeof(g.findBuf);
    g.fr.lpstrReplaceWith = g.replBuf;
    g.fr.wReplaceWithLen = (WORD)sizeof(g.replBuf);
    g.findDlg = replace ? ReplaceTextW(&g.fr) : FindTextW(&g.fr);
}

void onFindMessage(LPFINDREPLACEW p) {
    if (p->Flags & FR_DIALOGTERM) {
        g.findDlg = nullptr;
        return;
    }
    g.findCase = (p->Flags & FR_MATCHCASE) != 0;
    g.findWord = (p->Flags & FR_WHOLEWORD) != 0;
    if (p->Flags & FR_FINDNEXT) findNext((p->Flags & FR_DOWN) != 0);
    else if (p->Flags & FR_REPLACE) replaceOne();
    else if (p->Flags & FR_REPLACEALL) replaceAll();
}

INT_PTR CALLBACK gotoProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) {
    if (INT_PTR br = dialogTheme(dlg, msg, wp, lp)) return br;
    switch (msg) {
    case WM_INITDIALOG:
        if (g.d->doc.linesKnown()) {
            wchar_t b[32];
            swprintf(b, 32, L"%llu", (unsigned long long)g.d->doc.lineOf(g.d->caret) + 1);
            SetDlgItemTextW(dlg, IDC_LINE, b);
            SendDlgItemMessageW(dlg, IDC_LINE, EM_SETSEL, 0, -1);
        }
        return TRUE;
    case WM_COMMAND:
        if (LOWORD(wp) == IDOK) {
            wchar_t b[32] = L"";
            GetDlgItemTextW(dlg, IDC_LINE, b, 32);
            g.gotoLine = wcstoull(b, nullptr, 10);
            EndDialog(dlg, 1);
            return TRUE;
        }
        if (LOWORD(wp) == IDCANCEL) {
            EndDialog(dlg, 0);
            return TRUE;
        }
        break;
    }
    return FALSE;
}

void cmdGoto() {
    if (!g.d->doc.linesKnown()) {
        MessageBoxW(g.frame, L"Lines are still being counted for this file. Try again in a moment.", kAppName,
                    MB_OK | MB_ICONINFORMATION);
        return;
    }
    if (DialogBoxParamW(g.inst, MAKEINTRESOURCEW(IDD_GOTO), g.frame, gotoProc, 0) != 1) return;
    u64 total = g.d->doc.newlineCount() + 1;
    u64 line = g.gotoLine < 1 ? 1 : (g.gotoLine > total ? total : g.gotoLine);
    moveCaret(g.d->doc.lineStart(line - 1), false);
}

// ====================================================================== settings

std::wstring settingsPath(bool create) {
    wchar_t dir[MAX_PATH];
    DWORD n = GetEnvironmentVariableW(L"APPDATA", dir, MAX_PATH);
    if (n == 0 || n >= MAX_PATH) return L"";
    std::wstring d = std::wstring(dir) + L"\\JustNotepad";
    if (create) CreateDirectoryW(d.c_str(), nullptr);
    return d + L"\\settings.ini";
}

RECT g_savedRect = {CW_USEDEFAULT, CW_USEDEFAULT, 0, 0};
bool g_savedMax = false;

void loadSettings() {
    std::wstring p = settingsPath(false);
    if (p.empty() || GetFileAttributesW(p.c_str()) == INVALID_FILE_ATTRIBUTES) return;
    const wchar_t* f = p.c_str();
    wchar_t face[LF_FACESIZE];
    GetPrivateProfileStringW(L"view", L"font", L"Consolas", face, LF_FACESIZE, f);
    g.fontFace = face;
    g.fontPt = (int)GetPrivateProfileIntW(L"view", L"size", 11, f);
    if (g.fontPt < 6 || g.fontPt > 72) g.fontPt = 11;
    g.wrap = GetPrivateProfileIntW(L"view", L"wrap", 0, f) != 0;
    g.lineNumbers = GetPrivateProfileIntW(L"view", L"lineNumbers", 1, f) != 0;
    g.themeMode = (int)GetPrivateProfileIntW(L"view", L"theme", 0, f);
    g.highlight = GetPrivateProfileIntW(L"view", L"highlight", 1, f) != 0;
    if (g.themeMode < 0 || g.themeMode > 2) g.themeMode = 0;
    int w = (int)GetPrivateProfileIntW(L"window", L"w", 0, f);
    int h = (int)GetPrivateProfileIntW(L"window", L"h", 0, f);
    if (w > 100 && h > 100) {
        g_savedRect.left = (int)GetPrivateProfileIntW(L"window", L"x", 0, f) - 100000;
        g_savedRect.top = (int)GetPrivateProfileIntW(L"window", L"y", 0, f) - 100000;
        g_savedRect.right = w;
        g_savedRect.bottom = h;
    }
    g_savedMax = GetPrivateProfileIntW(L"window", L"max", 0, f) != 0;

    VcsConnection c;
    wchar_t buf[1024];
    auto str = [&](const wchar_t* key) {
        GetPrivateProfileStringW(L"vcs", key, L"", buf, 1024, f);
        return std::wstring(buf);
    };
    std::wstring only = str(L"use");
    c.only = only == L"git" ? VcsKind::Git : only == L"perforce" ? VcsKind::Perforce : VcsKind::None;
    c.gitExe = str(L"gitExe");
    c.p4Exe = str(L"p4Exe");
    c.p4Port = str(L"p4Port");
    c.p4User = str(L"p4User");
    c.p4Client = str(L"p4Client");
    c.autoCheckout = GetPrivateProfileIntW(L"vcs", L"autoCheckout", 1, f) != 0;
    setVcsConnection(c);
}

void putInt(const wchar_t* sec, const wchar_t* key, int v, const wchar_t* file) {
    wchar_t b[32];
    swprintf(b, 32, L"%d", v);
    WritePrivateProfileStringW(sec, key, b, file);
}

void saveSettings() {
    std::wstring p = settingsPath(true);
    if (p.empty()) return;
    const wchar_t* f = p.c_str();
    WritePrivateProfileStringW(L"view", L"font", g.fontFace.c_str(), f);
    putInt(L"view", L"size", g.fontPt, f);
    putInt(L"view", L"wrap", (g.diffOn ? g.wrapBeforeDiff : g.wrap) ? 1 : 0, f);
    putInt(L"view", L"lineNumbers", g.lineNumbers ? 1 : 0, f);
    putInt(L"view", L"theme", g.themeMode, f);
    putInt(L"view", L"highlight", g.highlight ? 1 : 0, f);
    WINDOWPLACEMENT wp = {};
    wp.length = sizeof(wp);
    if (GetWindowPlacement(g.frame, &wp)) {
        // Coordinates can be negative on multi-monitor setups; store them offset.
        putInt(L"window", L"x", wp.rcNormalPosition.left + 100000, f);
        putInt(L"window", L"y", wp.rcNormalPosition.top + 100000, f);
        putInt(L"window", L"w", wp.rcNormalPosition.right - wp.rcNormalPosition.left, f);
        putInt(L"window", L"h", wp.rcNormalPosition.bottom - wp.rcNormalPosition.top, f);
        putInt(L"window", L"max", wp.showCmd == SW_SHOWMAXIMIZED ? 1 : 0, f);
    }
    VcsConnection c = vcsConnection();
    WritePrivateProfileStringW(L"vcs", L"use",
                               c.only == VcsKind::Git ? L"git" : c.only == VcsKind::Perforce ? L"perforce" : L"auto", f);
    WritePrivateProfileStringW(L"vcs", L"gitExe", c.gitExe.c_str(), f);
    WritePrivateProfileStringW(L"vcs", L"p4Exe", c.p4Exe.c_str(), f);
    WritePrivateProfileStringW(L"vcs", L"p4Port", c.p4Port.c_str(), f);
    WritePrivateProfileStringW(L"vcs", L"p4User", c.p4User.c_str(), f);
    WritePrivateProfileStringW(L"vcs", L"p4Client", c.p4Client.c_str(), f);
    putInt(L"vcs", L"autoCheckout", c.autoCheckout ? 1 : 0, f);
}

// ====================================================================== line editing (Visual Studio keys)

void deleteLines() {
    u64 l0, l1;
    if (!selLines(l0, l1)) return;
    u64 a = g.d->doc.lineStart(l0), b = g.d->doc.lineStart(l1 + 1);
    if (l1 == lastLine() && l0 > 0) a = g.d->doc.lineStart(l0) - (g.d->crlf && a >= 2 && g.d->doc.at(a - 2) == '\r' ? 2 : 1);
    replaceRange(a, b, "", a, a);
}

void moveLines(bool up) {
    u64 l0, l1;
    if (!selLines(l0, l1)) return;
    if (up ? l0 == 0 : l1 >= lastLine()) return;
    const bool caretFirst = g.d->caret <= g.d->anchor;
    u64 sa = selMin(), sb = selMax();
    std::string block = linesText(l0, l1);
    if (up) {
        u64 p0 = g.d->doc.lineStart(l0 - 1), p1 = g.d->doc.lineStart(l0);
        std::string prev = g.d->doc.text(p0, p1 - p0);
        u64 end = g.d->doc.lineStart(l1 + 1);
        if (l1 == lastLine()) prev.erase(prev.size() - (g.d->crlf ? 2 : 1));  // block gained the break
        u64 shift = p1 - p0;
        u64 na = sa - shift, nb = sb - shift;
        replaceRange(p0, end, block + prev, caretFirst ? nb : na, caretFirst ? na : nb);
    } else {
        u64 b0 = g.d->doc.lineStart(l0), n0 = g.d->doc.lineStart(l1 + 1), n1 = g.d->doc.lineStart(l1 + 2);
        std::string next = g.d->doc.text(n0, n1 - n0);
        if (l1 + 1 == lastLine()) {  // next line is the last one: it gets the break, the block loses it
            next += eolStr();
            block.erase(block.size() - (g.d->crlf ? 2 : 1));
        }
        u64 shift = next.size();
        u64 na = sa + shift, nb = sb + shift;
        replaceRange(b0, n1, next + block, caretFirst ? nb : na, caretFirst ? na : nb);
    }
}

void duplicate() {
    if (hasSel()) {  // duplicate the selection right after itself and select the copy
        u64 a = selMin(), b = selMax();
        std::string t = g.d->doc.text(a, b - a);
        replaceRange(b, b, t, b, b + t.size());
        return;
    }
    u64 l0, l1;
    if (!selLines(l0, l1)) return;
    u64 ls = g.d->doc.lineStart(l0), off = g.d->caret - ls;
    std::string t = linesText(l0, l0);
    u64 at = g.d->doc.lineStart(l0 + 1);
    if (l0 == lastLine()) {  // last line has no break: put one before the copy
        std::string body = t.substr(0, t.size() - (g.d->crlf ? 2 : 1));
        t = eolStr() + body;
        replaceRange(at, at, t, at + strlen(eolStr()) + off, at + strlen(eolStr()) + off);
        return;
    }
    replaceRange(at, at, t, at + off, at + off);
}

// Ctrl+Enter: new line above, Ctrl+Shift+Enter: below; both keep the indent.
void openLine(bool below) {
    if (!g.d->doc.linesKnown()) return;
    u64 line = g.d->doc.lineOf(g.d->caret), ls = g.d->doc.lineStart(line);
    std::string indent;
    for (u64 p = ls; p < g.d->doc.size() && (g.d->doc.at(p) == ' ' || g.d->doc.at(p) == '\t'); p++) indent += (char)g.d->doc.at(p);
    if (below) {
        u64 next = g.d->doc.lineStart(line + 1);
        if (line == lastLine()) {
            std::string t = eolStr() + indent;
            replaceRange(g.d->doc.size(), g.d->doc.size(), t, g.d->doc.size() + t.size(), g.d->doc.size() + t.size());
        } else {
            std::string t = indent + eolStr();
            replaceRange(next, next, t, next + indent.size(), next + indent.size());
        }
    } else {
        std::string t = indent + eolStr();
        replaceRange(ls, ls, t, ls + indent.size(), ls + indent.size());
    }
}

// Tab with several lines selected indents them; Shift+Tab outdents.
void indentLines(bool out) {
    u64 l0, l1;
    if (!selLines(l0, l1)) return;
    u64 a = g.d->doc.lineStart(l0), b = g.d->doc.lineStart(l1 + 1);
    std::string t = g.d->doc.text(a, b - a), r;
    bool atStart = true;
    for (size_t i = 0; i < t.size(); i++) {
        if (atStart) {
            atStart = false;
            if (!out) {
                if (t[i] != '\r' && t[i] != '\n') r += '\t';
            } else if (t[i] == '\t') {
                continue;
            } else if (t[i] == ' ') {
                size_t k = 0;
                while (k < 4 && i + k < t.size() && t[i + k] == ' ') k++;
                i += k - 1;
                continue;
            }
        }
        r += t[i];
        if (t[i] == '\n') atStart = true;
    }
    if (r == t) return;
    // Keep the same lines selected.
    u64 end = a + r.size();
    if (l1 == lastLine()) replaceRange(a, b, r, a, end);
    else replaceRange(a, b, r, a, end);
}

void changeCase(bool upper) {
    if (!hasSel()) {
        MessageBeep(MB_OK);
        return;
    }
    u64 a = selMin(), b = selMax();
    std::wstring w = fromDocBytes(g.d->doc.text(a, b - a));
    if (upper) CharUpperBuffW(&w[0], (DWORD)w.size());
    else CharLowerBuffW(&w[0], (DWORD)w.size());
    std::string t = toDocBytes(w.c_str(), (int)w.size());
    const bool caretFirst = g.d->caret < g.d->anchor;
    replaceRange(a, b, t, caretFirst ? a + t.size() : a, caretFirst ? a : a + t.size());
}

// The word at the caret, as [a, b).
bool wordAt(u64& a, u64& b) {
    a = b = g.d->caret;
    const u64 n = g.d->doc.size();
    while (a > 0 && g.d->caret - a < 4096 && isWordByte(g.d->doc.at(a - 1))) a--;
    while (b < n && b - g.d->caret < 4096 && isWordByte(g.d->doc.at(b))) b++;
    return b > a;
}

void selectWord() {
    u64 a, b;
    if (!wordAt(a, b)) return;
    g.d->anchor = a;
    g.d->caret = b;
    g.d->wantCol = -1;
    refresh(true);
}

// Ctrl+F3 / Ctrl+Shift+F3: find the next / previous place the selection (or
// the word at the caret) occurs.
void findWord(bool down) {
    std::string what;
    if (hasSel() && selMax() - selMin() < 255) {
        what = g.d->doc.text(selMin(), selMax() - selMin());
    } else {
        u64 a, b;
        if (!wordAt(a, b)) return;
        what = g.d->doc.text(a, b - a);
        g.d->anchor = a;
        g.d->caret = b;
    }
    std::wstring w = fromDocBytes(what);
    if (w.empty() || w.size() >= 255 || w.find(L'\n') != std::wstring::npos) return;
    wcscpy_s(g.findBuf, w.c_str());
    findNext(down);
}

// Ctrl+]: jump to the matching bracket; with Shift, select up to it.
void gotoBracket(bool select) {
    static const char open[] = "([{", close[] = ")]}";
    const u64 n = g.d->doc.size();
    u64 at = NPOS;
    if (g.d->caret < n && strchr("()[]{}", g.d->doc.at(g.d->caret)) && g.d->doc.at(g.d->caret)) at = g.d->caret;
    else if (g.d->caret > 0 && strchr("()[]{}", g.d->doc.at(g.d->caret - 1)) && g.d->doc.at(g.d->caret - 1)) at = g.d->caret - 1;
    if (at == NPOS) {
        MessageBeep(MB_OK);
        return;
    }
    char c = (char)g.d->doc.at(at);
    const char* o = strchr(open, c);
    int idx = o ? (int)(o - open) : (int)(strchr(close, c) - close);
    const char me = open[idx], other = close[idx];
    const bool forward = o != nullptr;
    int depth = 0;
    const u64 limit = 16ull << 20;  // give up after 16 MB
    u64 p = at, steps = 0;
    for (;;) {
        u8 ch = g.d->doc.at(p);
        if (ch == (u8)me) depth += forward ? 1 : -1;
        else if (ch == (u8)other) depth += forward ? -1 : 1;
        if (depth == 0) break;
        if (forward ? ++p >= n : p-- == 0) {
            MessageBeep(MB_OK);
            return;
        }
        if (++steps > limit) return;
    }
    // Land just after a closing bracket, on an opening one.
    u64 target = forward ? p + 1 : p;
    if (select) {
        u64 from = forward ? at : at + 1;
        g.d->anchor = std::min(from, target);
        g.d->caret = std::max(from, target);
        if (!forward) std::swap(g.d->anchor, g.d->caret);
    } else {
        g.d->caret = g.d->anchor = target;
    }
    g.d->wantCol = -1;
    refresh(true);
}

// ====================================================================== theme

bool wantDark() { return g.themeMode == 2 || (g.themeMode == 0 && systemUsesDarkMode()); }

void applyTheme() {
    const bool dark = wantDark();
    g.pal = makePalette(dark);
    themeSetAppDark(dark);
    if (!g.frame) return;
    themeWindow(g.frame, dark);
    // Scroll bars of the editor follow the theme too.
    SetWindowTheme(g.view, dark ? L"DarkMode_Explorer" : nullptr, nullptr);
    SendMessageW(g.view, WM_THEMECHANGED, 0, 0);
    themeMenuColors(g.menu, g.pal);
    DrawMenuBar(g.frame);
    RedrawWindow(g.frame, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
}

// Shared by our dialogs: dark controls on WM_INITDIALOG, colours on WM_CTLCOLOR*.
INT_PTR dialogTheme(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_INITDIALOG) themeDialog(dlg, g.pal);
    if (msg == WM_DRAWITEM && themeDrawDialogButton((DRAWITEMSTRUCT*)lp, g.pal)) return TRUE;
    if (msg >= WM_CTLCOLORMSGBOX && msg <= WM_CTLCOLORSTATIC) return themeCtlColor(msg, wp, g.pal);
    return 0;
}

// ====================================================================== session
//
// Like Notepad: closing the window never asks about unsaved work. Every tab
// (files, edits not yet saved, Untitled text) is written to
// %APPDATA%\JustNotepad\session and comes back on the next start. The same
// is written every few seconds while you work, so a crash loses little.

std::wstring sessionDir(bool create) {
    std::wstring sp = settingsPath(create);
    if (sp.empty()) return L"";
    std::wstring d = dirOf(sp) + L"\\session";
    if (create) CreateDirectoryW(d.c_str(), nullptr);
    return d;
}

bool needsBackup(const DocState* t) { return t->doc.modified() || (t->path.empty() && t->doc.size() > 0); }

bool writeBackup(DocState* t, const std::wstring& file) {
    std::wstring tmp = file + L".tmp";
    HANDLE h = CreateFileW(tmp.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;
    bool ok = true;
    t->doc.forEachPiece([&](const u8* p, u64 n) {
        while (n && ok) {
            DWORD chunk = (DWORD)std::min<u64>(n, 1u << 30), put = 0;
            ok = WriteFile(h, p, chunk, &put, nullptr) && put == chunk;
            p += chunk;
            n -= chunk;
        }
        return ok;
    });
    CloseHandle(h);
    if (!ok || !MoveFileExW(tmp.c_str(), file.c_str(), MOVEFILE_REPLACE_EXISTING)) {
        DeleteFileW(tmp.c_str());
        return false;
    }
    return true;
}

void putU64(const wchar_t* sec, const wchar_t* key, u64 v, const wchar_t* file) {
    wchar_t b[32];
    swprintf(b, 32, L"%llu", (unsigned long long)v);
    WritePrivateProfileStringW(sec, key, b, file);
}

u64 getU64(const wchar_t* sec, const wchar_t* key, const wchar_t* file) {
    wchar_t b[32];
    GetPrivateProfileStringW(sec, key, L"0", b, 32, file);
    return wcstoull(b, nullptr, 10);
}

void saveSession() {
    std::wstring dir = sessionDir(true);
    if (dir.empty()) return;
    const std::wstring ini = dir + L"\\session.ini";
    std::wstring tmpIni = ini + L".tmp";
    DeleteFileW(tmpIni.c_str());
    const wchar_t* f = tmpIni.c_str();
    std::vector<std::wstring> keep;
    int n = 0, active = 0;
    for (auto& tp : g.tabs) {
        DocState* t = tp.get();
        if (t->path.empty() && !needsBackup(t)) continue;  // empty Untitled: nothing to keep
        wchar_t sec[16];
        swprintf(sec, 16, L"tab%d", n);
        WritePrivateProfileStringW(sec, L"path", t->path.c_str(), f);
        if (needsBackup(t)) {
            if (t->backupName.empty()) {
                wchar_t b[64];
                swprintf(b, 64, L"%lu-%llu-%u.txt", GetCurrentProcessId(), (unsigned long long)GetTickCount64(), ++g_gen);
                t->backupName = b;
            }
            // Rewrite only when the text changed since the last backup.
            std::wstring file = dir + L"\\" + t->backupName;
            if (t->backupVersion != t->doc.version() || GetFileAttributesW(file.c_str()) == INVALID_FILE_ATTRIBUTES) {
                if (writeBackup(t, file)) t->backupVersion = t->doc.version();
            }
            WritePrivateProfileStringW(sec, L"backup", t->backupName.c_str(), f);
            keep.push_back(t->backupName);
        }
        putU64(sec, L"caret", t == g.d ? g.d->caret : t->caret, f);
        putU64(sec, L"anchor", t->anchor, f);
        putU64(sec, L"top", t->top, f);
        putInt(sec, L"enc", (int)t->enc, f);
        putInt(sec, L"crlf", t->crlf ? 1 : 0, f);
        if (t == g.d) active = n;
        n++;
    }
    putInt(L"session", L"count", n, f);
    putInt(L"session", L"active", active, f);
    MoveFileExW(tmpIni.c_str(), ini.c_str(), MOVEFILE_REPLACE_EXISTING);
    // Backups that no tab needs any more.
    WIN32_FIND_DATAW fd;
    HANDLE fh = FindFirstFileW((dir + L"\\*.txt").c_str(), &fd);
    if (fh != INVALID_HANDLE_VALUE) {
        do {
            if (std::find(keep.begin(), keep.end(), std::wstring(fd.cFileName)) == keep.end())
                DeleteFileW((dir + L"\\" + fd.cFileName).c_str());
        } while (FindNextFileW(fh, &fd));
        FindClose(fh);
    }
}

// A cheap fingerprint of everything saveSession writes, to skip idle saves.
u64 sessionStamp() {
    u64 k = 1469598103934665603ull;
    auto mix = [&](u64 v) { k = (k ^ v) * 1099511628211ull; };
    mix(g.tabs.size());
    for (auto& t : g.tabs) {
        mix(t->doc.version());
        mix(t->doc.modified());
        mix(t->caret);
        mix(t->top);
        mix(std::hash<std::wstring>()(t->path));
    }
    mix((u64)(uintptr_t)g.d);
    return k;
}
u64 g_lastSessionStamp = 0;

void autosaveSession() {
    u64 st = sessionStamp();
    if (st == g_lastSessionStamp) return;
    g_lastSessionStamp = st;
    saveSession();
}

bool readFileBytes(const std::wstring& file, std::string& out) {
    HANDLE h = CreateFileW(file.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;
    LARGE_INTEGER sz;
    bool ok = GetFileSizeEx(h, &sz) && sz.QuadPart < (1ll << 40);
    if (ok) {
        out.resize((size_t)sz.QuadPart);
        u64 done = 0;
        while (ok && done < out.size()) {
            DWORD chunk = (DWORD)std::min<u64>(out.size() - done, 1u << 30), got = 0;
            ok = ReadFile(h, &out[done], chunk, &got, nullptr) && got > 0;
            done += got;
        }
    }
    CloseHandle(h);
    return ok;
}

// Reopens the tabs of the last session. Returns true if any came back.
bool restoreSession() {
    std::wstring dir = sessionDir(false);
    if (dir.empty()) return false;
    const std::wstring ini = dir + L"\\session.ini";
    const wchar_t* f = ini.c_str();
    int n = (int)GetPrivateProfileIntW(L"session", L"count", 0, f);
    int active = (int)GetPrivateProfileIntW(L"session", L"active", 0, f);
    int activeTab = -1, restored = 0;
    for (int i = 0; i < n; i++) {
        wchar_t sec[16], path[32768], bak[128];
        swprintf(sec, 16, L"tab%d", i);
        GetPrivateProfileStringW(sec, L"path", L"", path, 32768, f);
        GetPrivateProfileStringW(sec, L"backup", L"", bak, 128, f);
        const bool exists = path[0] && GetFileAttributesW(path) != INVALID_FILE_ATTRIBUTES;
        std::string text;
        const bool haveBak = bak[0] && readFileBytes(dir + L"\\" + bak, text);
        if (!exists && !haveBak) continue;  // a file that is gone, with nothing unsaved
        if (exists) {
            if (!openInTab(path)) continue;
        } else {
            if (!isPristine(g.d)) addTab();
            g.d->path = path;  // Untitled, or a file deleted since: Save recreates it
            g.d->enc = (FileEnc)std::clamp((int)GetPrivateProfileIntW(sec, L"enc", 0, f), 0, (int)FE_ANSI);
            g.d->crlf = GetPrivateProfileIntW(sec, L"crlf", 1, f) != 0;
        }
        if (haveBak) {
            // Put the unsaved text back as one edit, so it shows as modified
            // and Undo returns to what is on disk.
            g.d->doc.breakCoalesce();
            g.d->doc.beginGroup();
            if (g.d->doc.size()) g.d->doc.erase(0, g.d->doc.size());
            if (!text.empty()) g.d->doc.insert(0, (const u8*)text.data(), text.size());
            g.d->doc.endGroup();
            g.d->doc.breakCoalesce();
            g.d->backupName = bak;
            detectLang();  // Untitled tabs may have a "#!" line
            g.d->backupVersion = g.d->doc.version();
        }
        const u64 size = g.d->doc.size();
        g.d->caret = std::min(getU64(sec, L"caret", f), size);
        g.d->anchor = std::min(getU64(sec, L"anchor", f), size);
        g.d->top = std::min(getU64(sec, L"top", f), size);
        if (i == active) activeTab = tabIndex(g.d);
        restored++;
    }
    if (activeTab >= 0) activateTab(activeTab);
    if (restored) {
        recomputeDiff();
        refresh(false);
        invalidateTabs();
    }
    return restored > 0;
}

// ====================================================================== menu and commands

HMENU buildMenu() {
    HMENU bar = CreateMenu();
    HMENU file = CreatePopupMenu();
    AppendMenuW(file, MF_STRING, IDM_NEW, L"&New\tCtrl+N");
    AppendMenuW(file, MF_STRING, IDM_OPEN, L"&Open...\tCtrl+O");
    AppendMenuW(file, MF_STRING, IDM_SAVE, L"&Save\tCtrl+S");
    AppendMenuW(file, MF_STRING, IDM_SAVEAS, L"Save &As...\tCtrl+Shift+S");
    AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
    HMENU tabsMenu = CreatePopupMenu();
    AppendMenuW(tabsMenu, MF_STRING, IDM_NEWTAB, L"&New Tab\tCtrl+T");
    AppendMenuW(tabsMenu, MF_STRING, IDM_CLOSETAB, L"&Close Tab\tCtrl+W");
    AppendMenuW(tabsMenu, MF_STRING, IDM_CLOSEOTHERS, L"Close &Other Tabs");
    AppendMenuW(tabsMenu, MF_STRING, IDM_REOPENTAB, L"&Reopen Closed Tab\tCtrl+Shift+T");
    AppendMenuW(tabsMenu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(tabsMenu, MF_STRING, IDM_NEXTTAB, L"Ne&xt Tab\tCtrl+Tab");
    AppendMenuW(tabsMenu, MF_STRING, IDM_PREVTAB, L"&Previous Tab\tCtrl+Shift+Tab");
    AppendMenuW(tabsMenu, MF_STRING, IDM_MOVETABLEFT, L"Move Tab &Left\tCtrl+Shift+PgUp");
    AppendMenuW(tabsMenu, MF_STRING, IDM_MOVETABRIGHT, L"Move Tab Ri&ght\tCtrl+Shift+PgDn");
    AppendMenuW(file, MF_POPUP, (UINT_PTR)tabsMenu, L"&Tabs");
    AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(file, MF_STRING, IDM_EXIT, L"E&xit");
    HMENU edit = CreatePopupMenu();
    AppendMenuW(edit, MF_STRING, IDM_UNDO, L"&Undo\tCtrl+Z");
    AppendMenuW(edit, MF_STRING, IDM_REDO, L"&Redo\tCtrl+Y");
    AppendMenuW(edit, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(edit, MF_STRING, IDM_CUT, L"Cu&t\tCtrl+X");
    AppendMenuW(edit, MF_STRING, IDM_COPY, L"&Copy\tCtrl+C");
    AppendMenuW(edit, MF_STRING, IDM_PASTE, L"&Paste\tCtrl+V");
    AppendMenuW(edit, MF_STRING, IDM_DELETE, L"De&lete\tDel");
    AppendMenuW(edit, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(edit, MF_STRING, IDM_FIND, L"&Find...\tCtrl+F");
    AppendMenuW(edit, MF_STRING, IDM_FINDNEXT, L"Find &Next\tF3");
    AppendMenuW(edit, MF_STRING, IDM_FINDPREV, L"Find Pre&vious\tShift+F3");
    AppendMenuW(edit, MF_STRING, IDM_REPLACE, L"R&eplace...\tCtrl+H");
    AppendMenuW(edit, MF_STRING, IDM_GOTO, L"&Go To Line...\tCtrl+G");
    AppendMenuW(edit, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(edit, MF_STRING, IDM_SELECTALL, L"Select &All\tCtrl+A");
    AppendMenuW(edit, MF_STRING, IDM_SELECTWORD, L"Select &Word");
    AppendMenuW(edit, MF_SEPARATOR, 0, nullptr);
    HMENU line = CreatePopupMenu();
    AppendMenuW(line, MF_STRING, IDM_LINE_CUT, L"Cu&t Line\tCtrl+L");
    AppendMenuW(line, MF_STRING, IDM_LINE_DELETE, L"&Delete Line\tCtrl+Shift+L");
    AppendMenuW(line, MF_STRING, IDM_LINE_DUP, L"D&uplicate\tCtrl+D");
    AppendMenuW(line, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(line, MF_STRING, IDM_LINE_UP, L"Move &Up\tAlt+Up");
    AppendMenuW(line, MF_STRING, IDM_LINE_DOWN, L"Move Dow&n\tAlt+Down");
    AppendMenuW(line, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(line, MF_STRING, IDM_LINE_ABOVE, L"Insert Line &Above\tCtrl+Enter");
    AppendMenuW(line, MF_STRING, IDM_LINE_BELOW, L"Insert Line &Below\tCtrl+Shift+Enter");
    AppendMenuW(line, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(line, MF_STRING, IDM_OUTDENT, L"&Outdent\tShift+Tab");
    AppendMenuW(edit, MF_POPUP, (UINT_PTR)line, L"&Line");
    HMENU adv = CreatePopupMenu();
    AppendMenuW(adv, MF_STRING, IDM_UPPER, L"Make &Uppercase\tCtrl+Shift+U");
    AppendMenuW(adv, MF_STRING, IDM_LOWER, L"Make &Lowercase\tCtrl+U");
    AppendMenuW(adv, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(adv, MF_STRING, IDM_BRACKET, L"Go to Matching &Bracket\tCtrl+]");
    AppendMenuW(adv, MF_STRING, IDM_BRACKET_SEL, L"&Select to Matching Bracket\tCtrl+Shift+]");
    AppendMenuW(adv, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(adv, MF_STRING, IDM_FINDWORD, L"Find &Next Occurrence\tCtrl+F3");
    AppendMenuW(adv, MF_STRING, IDM_FINDWORD_PREV, L"Find &Previous Occurrence\tCtrl+Shift+F3");
    AppendMenuW(edit, MF_POPUP, (UINT_PTR)adv, L"Ad&vanced");
    HMENU view = CreatePopupMenu();
    AppendMenuW(view, MF_STRING, IDM_WRAP, L"&Word Wrap\tAlt+Z");
    AppendMenuW(view, MF_STRING, IDM_LINENUMBERS, L"&Line Numbers");
    AppendMenuW(view, MF_STRING, IDM_HIGHLIGHT, L"Syntax &Highlighting");
    AppendMenuW(view, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(view, MF_STRING, IDM_FONT, L"&Font...");
    AppendMenuW(view, MF_STRING, IDM_ZOOMIN, L"Zoom &In\tCtrl+Plus");
    AppendMenuW(view, MF_STRING, IDM_ZOOMOUT, L"Zoom &Out\tCtrl+Minus");
    AppendMenuW(view, MF_STRING, IDM_ZOOMRESET, L"&Reset Zoom\tCtrl+0");
    AppendMenuW(view, MF_SEPARATOR, 0, nullptr);
    HMENU theme = CreatePopupMenu();
    AppendMenuW(theme, MF_STRING, IDM_THEME_SYSTEM, L"Use &System Setting");
    AppendMenuW(theme, MF_STRING, IDM_THEME_LIGHT, L"&Light");
    AppendMenuW(theme, MF_STRING, IDM_THEME_DARK, L"&Dark");
    AppendMenuW(view, MF_POPUP, (UINT_PTR)theme, L"&Theme");
    HMENU sc = CreatePopupMenu();
    AppendMenuW(sc, MF_STRING, IDM_DIFF, L"&Diff\tF8");
    AppendMenuW(sc, MF_STRING, IDM_VCS_FILES, L"&Open Changelist Files\tCtrl+Shift+O");
    AppendMenuW(sc, MF_STRING, IDM_VCS_COMMIT, L"&Submit Current Changelist...");
    AppendMenuW(sc, MF_STRING, IDM_VCS_REVERTFILE, L"&Revert Changes...");
    AppendMenuW(sc, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(sc, MF_STRING, IDM_VCS_REFRESH, L"Re&fresh Status\tF5");
    AppendMenuW(sc, MF_STRING, IDM_VCS_CONNECT, L"&Connect...");
    HMENU help = CreatePopupMenu();
    AppendMenuW(help, MF_STRING, IDM_CHECKUPDATES, L"Check for &Updates...");
    AppendMenuW(help, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(help, MF_STRING, IDM_ABOUT, L"&About JustNotepad");
    AppendMenuW(bar, MF_POPUP, (UINT_PTR)file, L"&File");
    AppendMenuW(bar, MF_POPUP, (UINT_PTR)edit, L"&Edit");
    AppendMenuW(bar, MF_POPUP, (UINT_PTR)view, L"&View");
    AppendMenuW(bar, MF_POPUP, (UINT_PTR)sc, L"&Source Control");
    AppendMenuW(bar, MF_POPUP, (UINT_PTR)help, L"&Help");
    return bar;
}

HACCEL buildAccelerators() {
    static ACCEL a[] = {
        {FVIRTKEY | FCONTROL, 'N', IDM_NEW},
        {FVIRTKEY | FCONTROL, 'O', IDM_OPEN},
        {FVIRTKEY | FCONTROL, VK_F4, IDM_CLOSETAB},
        {FVIRTKEY | FCONTROL, VK_TAB, IDM_NEXTTAB},
        {FVIRTKEY | FCONTROL | FSHIFT, VK_TAB, IDM_PREVTAB},
        {FVIRTKEY | FCONTROL | FSHIFT, 'O', IDM_VCS_FILES},
        {FVIRTKEY | FCONTROL, 'S', IDM_SAVE},
        {FVIRTKEY | FCONTROL | FSHIFT, 'S', IDM_SAVEAS},
        {FVIRTKEY | FCONTROL, 'Z', IDM_UNDO},
        {FVIRTKEY | FCONTROL, 'Y', IDM_REDO},
        {FVIRTKEY | FCONTROL | FSHIFT, 'Z', IDM_REDO},
        {FVIRTKEY | FCONTROL, 'X', IDM_CUT},
        {FVIRTKEY | FCONTROL, 'C', IDM_COPY},
        {FVIRTKEY | FCONTROL, 'V', IDM_PASTE},
        {FVIRTKEY | FCONTROL, 'A', IDM_SELECTALL},
        {FVIRTKEY | FCONTROL, 'L', IDM_LINE_CUT},
        {FVIRTKEY | FCONTROL | FSHIFT, 'L', IDM_LINE_DELETE},
        {FVIRTKEY | FCONTROL, 'D', IDM_LINE_DUP},
        {FVIRTKEY | FALT, VK_UP, IDM_LINE_UP},
        {FVIRTKEY | FALT, VK_DOWN, IDM_LINE_DOWN},
        {FVIRTKEY | FCONTROL, VK_RETURN, IDM_LINE_ABOVE},
        {FVIRTKEY | FCONTROL | FSHIFT, VK_RETURN, IDM_LINE_BELOW},
        {FVIRTKEY | FCONTROL | FSHIFT, 'U', IDM_UPPER},
        {FVIRTKEY | FCONTROL, 'U', IDM_LOWER},
        {FVIRTKEY | FCONTROL, 'W', IDM_CLOSETAB},
        {FVIRTKEY | FCONTROL, 'T', IDM_NEWTAB},
        {FVIRTKEY | FCONTROL | FSHIFT, 'T', IDM_REOPENTAB},
        {FVIRTKEY | FCONTROL, VK_PRIOR, IDM_PREVTAB},
        {FVIRTKEY | FCONTROL, VK_NEXT, IDM_NEXTTAB},
        {FVIRTKEY | FCONTROL | FSHIFT, VK_PRIOR, IDM_MOVETABLEFT},
        {FVIRTKEY | FCONTROL | FSHIFT, VK_NEXT, IDM_MOVETABRIGHT},
        {FVIRTKEY | FCONTROL, '1', IDM_TAB1 + 0},
        {FVIRTKEY | FCONTROL, '2', IDM_TAB1 + 1},
        {FVIRTKEY | FCONTROL, '3', IDM_TAB1 + 2},
        {FVIRTKEY | FCONTROL, '4', IDM_TAB1 + 3},
        {FVIRTKEY | FCONTROL, '5', IDM_TAB1 + 4},
        {FVIRTKEY | FCONTROL, '6', IDM_TAB1 + 5},
        {FVIRTKEY | FCONTROL, '7', IDM_TAB1 + 6},
        {FVIRTKEY | FCONTROL, '8', IDM_TAB1 + 7},
        {FVIRTKEY | FCONTROL, '9', IDM_TAB1 + 8},
        {FVIRTKEY | FCONTROL, VK_OEM_6, IDM_BRACKET},
        {FVIRTKEY | FCONTROL | FSHIFT, VK_OEM_6, IDM_BRACKET_SEL},
        {FVIRTKEY | FCONTROL, VK_F3, IDM_FINDWORD},
        {FVIRTKEY | FCONTROL | FSHIFT, VK_F3, IDM_FINDWORD_PREV},
        {FVIRTKEY | FCONTROL, 'F', IDM_FIND},
        {FVIRTKEY | FCONTROL, 'H', IDM_REPLACE},
        {FVIRTKEY | FCONTROL, 'G', IDM_GOTO},
        {FVIRTKEY, VK_F3, IDM_FINDNEXT},
        {FVIRTKEY | FSHIFT, VK_F3, IDM_FINDPREV},
        {FVIRTKEY | FALT, 'Z', IDM_WRAP},
        {FVIRTKEY | FCONTROL, VK_OEM_PLUS, IDM_ZOOMIN},
        {FVIRTKEY | FCONTROL, VK_ADD, IDM_ZOOMIN},
        {FVIRTKEY | FCONTROL, VK_OEM_MINUS, IDM_ZOOMOUT},
        {FVIRTKEY | FCONTROL, VK_SUBTRACT, IDM_ZOOMOUT},
        {FVIRTKEY | FCONTROL, '0', IDM_ZOOMRESET},
        {FVIRTKEY, VK_F8, IDM_DIFF},
        {FVIRTKEY, VK_F7, IDM_NEXTCHANGE},
        {FVIRTKEY | FSHIFT, VK_F7, IDM_PREVCHANGE},
        {FVIRTKEY | FCONTROL, 'R', IDM_REVERTCHANGE},
        {FVIRTKEY, VK_F5, IDM_VCS_REFRESH},
    };
    return CreateAcceleratorTableW(a, (int)(sizeof(a) / sizeof(a[0])));
}

void layoutFrame() {
    RECT rc;
    GetClientRect(g.frame, &rc);
    int h = std::max<int>(0, rc.bottom - g.statusH);
    int top = tabBarH() + headerH();
    int left = 0;
    if (g.diffOn) {
        int sw = stripW();
        left = std::max(0, ((int)rc.right - sw) / 2);
        MoveWindow(g.diffPane, 0, top, left, h - top, TRUE);
        MoveWindow(g.strip, left, top, sw, h - top, TRUE);
        left += sw;
    }
    MoveWindow(g.view, left, top, rc.right - left, h - top, TRUE);
    RECT s = {0, h, rc.right, rc.bottom};
    InvalidateRect(g.frame, &s, FALSE);
    RECT hd = {0, 0, rc.right, top};
    InvalidateRect(g.frame, &hd, FALSE);
}

// ====================================================================== tabs

DocState* tabByVcsGen(unsigned gen) {
    for (auto& t : g.tabs)
        if (t->vcsGen == gen) return t.get();
    return nullptr;
}

int tabIndex(const DocState* d) {
    for (size_t i = 0; i < g.tabs.size(); i++)
        if (g.tabs[i].get() == d) return (int)i;
    return -1;
}

void invalidateTabs() {
    if (!g.frame) return;
    RECT r = {0, 0, 32767, tabBarH()};
    InvalidateRect(g.frame, &r, FALSE);
}

std::wstring tabTitle(const DocState* d) {
    std::wstring t = d->path.empty() ? L"Untitled" : fileNameOf(d->path);
    return d->doc.modified() ? t + L" \x2022" : t;
}

// Shows tab `i`: the view state lives in each tab, so this is just a pointer swap.
void activateTab(int i) {
    if (i < 0 || i >= (int)g.tabs.size()) return;
    if (g.tabs[(size_t)i].get() == g.d) return;
    if (g.findDlg) g.d->doc.breakCoalesce();
    g.d = g.tabs[(size_t)i].get();
    KillTimer(g.frame, kTimerIndex);
    if (g.d->indexing) SetTimer(g.frame, kTimerIndex, 200, nullptr);
    g.hoverHunk = g.hoverChoice = -1;
    g.hint.clear();
    recomputeDiff();
    relayout();
    refresh(false);
    invalidateTabs();
    InvalidateRect(g.diffPane, nullptr, FALSE);
    InvalidateRect(g.strip, nullptr, FALSE);
    SetFocus(g.view);
}

// A new empty tab, made active.
void addTab() {
    g.tabs.push_back(std::make_unique<DocState>());
    g.d = g.tabs.back().get();
    newDocument();
    invalidateTabs();
}

bool isPristine(const DocState* d) { return d->path.empty() && !d->doc.modified() && d->doc.size() == 0; }

// Opens `path` in a tab: switches to it if it is already open, reuses an
// empty Untitled tab, otherwise adds one.
bool openInTab(const std::wstring& path) {
    for (size_t i = 0; i < g.tabs.size(); i++)
        if (_wcsicmp(g.tabs[i]->path.c_str(), path.c_str()) == 0) {
            activateTab((int)i);
            return true;
        }
    DocState* prev = g.d;
    bool added = false;
    if (!isPristine(g.d)) {
        g.tabs.push_back(std::make_unique<DocState>());
        g.d = g.tabs.back().get();
        newDocument();
        added = true;
    }
    if (!openPath(path, false)) {
        if (added) {
            g.tabs.pop_back();
            g.d = prev;
            refresh(false);
        }
        invalidateTabs();
        return false;
    }
    invalidateTabs();
    InvalidateRect(g.diffPane, nullptr, FALSE);
    return true;
}

// Closes tab `i`, asking to save first. Returns false if cancelled.
bool closeTab(int i) {
    if (i < 0 || i >= (int)g.tabs.size()) return true;
    DocState* keep = g.d;
    g.d = g.tabs[(size_t)i].get();
    const bool ok = confirmDiscard();
    if (!ok) {
        g.d = keep;
        return false;
    }
    const bool wasActive = keep == g.d;
    g.d = keep;
    if (g.tabs.size() == 1) {  // the last tab becomes an empty Untitled
        newDocument();
        invalidateTabs();
        return true;
    }
    if (!g.tabs[(size_t)i]->path.empty()) {
        const DocState* t = g.tabs[(size_t)i].get();
        g.closed.push_back({t->path, t->caret, t->top});
        if (g.closed.size() > 30) g.closed.erase(g.closed.begin());
    }
    std::unique_ptr<DocState> dead = std::move(g.tabs[(size_t)i]);
    g.tabs.erase(g.tabs.begin() + i);
    if (wasActive) {
        g.d = nullptr;
        int next = std::min(i, (int)g.tabs.size() - 1);
        g.d = g.tabs[(size_t)next].get();
        KillTimer(g.frame, kTimerIndex);
        if (g.d->indexing) SetTimer(g.frame, kTimerIndex, 200, nullptr);
        recomputeDiff();
        refresh(false);
    }
    dead->index.reset();  // stop its worker before the file it reads is released
    dead.reset();
    invalidateTabs();
    InvalidateRect(g.diffPane, nullptr, FALSE);
    InvalidateRect(g.strip, nullptr, FALSE);
    return true;
}

bool closeAllTabs() {
    // Ask about every modified tab, showing each one; stop at the first Cancel.
    for (size_t i = 0; i < g.tabs.size(); i++)
        if (g.tabs[i]->doc.modified()) {
            activateTab((int)i);
            if (!confirmDiscard()) return false;
        }
    return true;
}

void paintTabs(HDC dc) {
    RECT rc;
    GetClientRect(g.frame, &rc);
    const int H = tabBarH(), pad = MulDiv(12, g.dpi, 96), xw = MulDiv(16, g.dpi, 96), gap = MulDiv(1, g.dpi, 96);
    const COLORREF face = g.pal.face, text = g.pal.faceText, bg = g.pal.bg;
    fillRect(dc, 0, 0, rc.right, H, blend(face, text, 92));
    g.tabRects.assign(g.tabs.size(), RECT{});
    g.tabCloseRects.assign(g.tabs.size(), RECT{});
    HGDIOBJ of = SelectObject(dc, g.uiFont);
    SetBkMode(dc, TRANSPARENT);
    int x = 0;
    const int maxW = MulDiv(240, g.dpi, 96), minW = MulDiv(90, g.dpi, 96);
    for (size_t i = 0; i < g.tabs.size(); i++) {
        const DocState* t = g.tabs[i].get();
        const bool active = t == g.d, hot = (int)i == g.tabHover;
        std::wstring title = tabTitle(t);
        SIZE sz = {};
        GetTextExtentPoint32W(dc, title.c_str(), (int)title.size(), &sz);
        int w = std::clamp((int)sz.cx + 2 * pad + xw, minW, maxW);
        RECT r = {x, 0, x + w, H};
        fillRect(dc, r.left, r.top, r.right, r.bottom, active ? bg : hot ? blend(face, bg, 50) : blend(face, text, 92));
        if (active) fillRect(dc, r.left, 0, r.right, MulDiv(2, g.dpi, 96), g.pal.accent);  // the open one
        RECT tr = {r.left + pad, 0, r.right - pad / 2 - xw, H};
        SetTextColor(dc, active ? g.pal.fg : blend(text, face, 75));
        DrawTextW(dc, title.c_str(), (int)title.size(), &tr, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
        // Close button: shown on the open tab and the one under the mouse.
        RECT xr = {r.right - pad / 2 - xw, (H - xw) / 2, r.right - pad / 2, (H - xw) / 2 + xw};
        if (active || hot) {
            // A flat × that rises into a 3D button when hovered, and goes in when pressed.
            const bool xhot = (int)i == g.tabCloseHover, xheld = xhot && g.tabClosePress == (int)i;
            if (xhot) drawButton3D(dc, xr, xheld ? BTN_PRESSED : BTN_HOT, g.pal, 0);
            drawGlyph(dc, xr, GLYPH_CLOSE, xhot ? g.pal.fg : blend(text, face, 70), 0, xheld);
        }
        g.tabRects[i] = r;
        g.tabCloseRects[i] = xr;
        x += w + gap;
    }
    SelectObject(dc, of);
}

int tabAt(POINT pt, bool* onClose) {
    for (size_t i = 0; i < g.tabRects.size(); i++)
        if (PtInRect(&g.tabRects[i], pt)) {
            if (onClose) *onClose = PtInRect(&g.tabCloseRects[i], pt) != 0;
            return (int)i;
        }
    return -1;
}

void setZoom(int pt) {
    pt = std::max(6, std::min(72, pt));
    if (pt == g.fontPt) return;
    g.fontPt = pt;
    makeFonts();
    refresh(true);
}

void setWrap(bool on) {
    g.wrap = on;
    g.d->hcol = 0;
    ShowScrollBar(g.view, SB_HORZ, on ? FALSE : TRUE);
    refresh(true);
}

void cmdFont() {
    LOGFONTW lf = {};
    lf.lfHeight = -MulDiv(g.fontPt, g.dpi, 72);
    lstrcpynW(lf.lfFaceName, g.fontFace.c_str(), LF_FACESIZE);
    CHOOSEFONTW cf = {};
    cf.lStructSize = sizeof(cf);
    cf.hwndOwner = g.frame;
    cf.lpLogFont = &lf;
    cf.Flags = CF_SCREENFONTS | CF_FIXEDPITCHONLY | CF_INITTOLOGFONTSTRUCT | CF_NOSCRIPTSEL | CF_NOVERTFONTS;
    if (!ChooseFontW(&cf)) return;
    g.fontFace = lf.lfFaceName;
    g.fontPt = std::max(6, std::min(72, cf.iPointSize / 10));
    makeFonts();
    refresh(true);
}

INT_PTR CALLBACK aboutProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) {
    if (INT_PTR br = dialogTheme(dlg, msg, wp, lp)) return br;
    static HFONT titleFont = nullptr;
    switch (msg) {
    case WM_INITDIALOG: {
        // The name in a larger, bold version of the dialog font.
        LOGFONTW lf = {};
        GetObjectW((HFONT)SendMessageW(dlg, WM_GETFONT, 0, 0), sizeof(lf), &lf);
        lf.lfHeight = lf.lfHeight * 16 / 10;
        lf.lfWeight = FW_BOLD;
        if (titleFont) DeleteObject(titleFont);
        titleFont = CreateFontIndirectW(&lf);
        SendDlgItemMessageW(dlg, IDC_ABOUT_TITLE, WM_SETFONT, (WPARAM)titleFont, TRUE);
        // A big, sharp icon.
        int sz = MulDiv(32, (int)GetDpiForWindow(dlg), 96);
        HICON icon = nullptr;
        LoadIconWithScaleDown(g.inst, MAKEINTRESOURCEW(IDI_APP), sz, sz, &icon);
        if (icon) SendDlgItemMessageW(dlg, IDC_ABOUT_ICON, STM_SETICON, (WPARAM)icon, 0);
        SetDlgItemTextW(dlg, IDC_ABOUT_VERSION, L"Version " JN_WIDEN(APP_VERSION_STRING));
        return TRUE;
    }
    case WM_NOTIFY: {
        NMHDR* nm = (NMHDR*)lp;
        if (nm->idFrom == IDC_ABOUT_LINK && (nm->code == NM_CLICK || nm->code == NM_RETURN)) {
            ShellExecuteW(nullptr, L"open", ((PNMLINK)lp)->item.szUrl, nullptr, nullptr, SW_SHOWNORMAL);
            return TRUE;
        }
        break;
    }
    case WM_COMMAND:
        if (LOWORD(wp) == IDC_ABOUT_UPDATE) {
            EndDialog(dlg, IDOK);
            PostMessageW(g.frame, WM_COMMAND, IDM_CHECKUPDATES, 0);
            return TRUE;
        }
        if (LOWORD(wp) == IDOK || LOWORD(wp) == IDCANCEL) {
            EndDialog(dlg, LOWORD(wp));
            return TRUE;
        }
        break;
    }
    return FALSE;
}

void onCommand(int id) {
    if (id >= IDM_TAB1 && id < IDM_TAB1 + 9) {
        // Ctrl+1..8: that tab; Ctrl+9: the last one, as in browsers.
        int i = id - IDM_TAB1;
        activateTab(i == 8 ? (int)g.tabs.size() - 1 : i);
        return;
    }
    switch (id) {
    case IDM_NEW:
        if (!isPristine(g.d)) addTab();
        break;
    case IDM_OPEN: cmdOpen(); break;
    case IDM_CLOSETAB: closeTab(tabIndex(g.d)); break;
    case IDM_NEWTAB: addTab(); break;
    case IDM_CLOSEOTHERS:
        for (int i = (int)g.tabs.size() - 1; i >= 0; i--)
            if (g.tabs[(size_t)i].get() != g.d && !closeTab(i)) break;
        break;
    case IDM_REOPENTAB:
        // The most recently closed file that still exists.
        while (!g.closed.empty()) {
            ClosedTab c = g.closed.back();
            g.closed.pop_back();
            if (GetFileAttributesW(c.path.c_str()) == INVALID_FILE_ATTRIBUTES) continue;
            if (openInTab(c.path)) {
                g.d->caret = g.d->anchor = std::min(c.caret, g.d->doc.size());
                g.d->top = std::min(c.top, g.d->doc.size());
                refresh(false);
            }
            break;
        }
        break;
    case IDM_MOVETABLEFT:
    case IDM_MOVETABRIGHT: {
        int i = tabIndex(g.d), j = i + (id == IDM_MOVETABLEFT ? -1 : 1);
        if (i >= 0 && j >= 0 && j < (int)g.tabs.size()) {
            std::swap(g.tabs[(size_t)i], g.tabs[(size_t)j]);
            invalidateTabs();
        }
        break;
    }
    case IDM_NEXTTAB:
    case IDM_PREVTAB: {
        int n = (int)g.tabs.size(), i = tabIndex(g.d);
        if (n > 1) activateTab((i + (id == IDM_NEXTTAB ? 1 : n - 1)) % n);
        break;
    }
    case IDM_SAVE: cmdSave(); break;
    case IDM_SAVEAS: cmdSaveAs(); break;
    case IDM_EXIT: SendMessageW(g.frame, WM_CLOSE, 0, 0); break;
    case IDM_UNDO: {
        u64 c;
        if (g.d->doc.undo(c)) afterEdit(c);
        break;
    }
    case IDM_REDO: {
        u64 c;
        if (g.d->doc.redo(c)) afterEdit(c);
        break;
    }
    case IDM_CUT:
        if (!hasSel()) {  // no selection: cut the whole line, like Visual Studio
            copyLines(true);
            break;
        }
        doCopy();
        if (hasSel() && selMax() - selMin() <= kMaxClipboard) deleteSel();
        break;
    case IDM_LINE_CUT: copyLines(true); break;
    case IDM_LINE_DELETE: deleteLines(); break;
    case IDM_LINE_DUP: duplicate(); break;
    case IDM_LINE_UP: moveLines(true); break;
    case IDM_LINE_DOWN: moveLines(false); break;
    case IDM_LINE_ABOVE: openLine(false); break;
    case IDM_LINE_BELOW: openLine(true); break;
    case IDM_OUTDENT: indentLines(true); break;
    case IDM_UPPER: changeCase(true); break;
    case IDM_LOWER: changeCase(false); break;
    case IDM_SELECTWORD: selectWord(); break;
    case IDM_BRACKET: gotoBracket(false); break;
    case IDM_BRACKET_SEL: gotoBracket(true); break;
    case IDM_FINDWORD: findWord(true); break;
    case IDM_FINDWORD_PREV: findWord(false); break;
    case IDM_COPY: doCopy(); break;
    case IDM_PASTE: doPaste(); break;
    case IDM_DELETE: deleteSel(); break;
    case IDM_SELECTALL:
        g.d->anchor = 0;
        g.d->caret = g.d->doc.size();
        g.d->doc.breakCoalesce();
        refresh(false);
        break;
    case IDM_FIND: showFind(false); break;
    case IDM_REPLACE: showFind(true); break;
    case IDM_FINDNEXT:
        if (g.findBuf[0]) findNext(true);
        else showFind(false);
        break;
    case IDM_FINDPREV:
        if (g.findBuf[0]) findNext(false);
        else showFind(false);
        break;
    case IDM_GOTO: cmdGoto(); break;
    case IDM_WRAP:
        if (!g.diffOn) setWrap(!g.wrap);
        break;
    case IDM_DIFF: toggleDiff(); break;
    case IDM_NEXTCHANGE: gotoChange(true); break;
    case IDM_PREVCHANGE: gotoChange(false); break;
    case IDM_REVERTCHANGE: revertChange(); break;
    case IDM_VCS_PREPARE: vcsPrepare(); break;
    case IDM_VCS_COMMIT: vcsCommit(); break;
    case IDM_VCS_REVERTFILE: vcsRevertFile(); break;
    case IDM_VCS_CONNECT: cmdConnect(); break;
    case IDM_CHECKUPDATES: cmdCheckUpdates(); break;
    case IDM_VCS_REFRESH:
        g.d->p4Asked = false;  // a failed automatic checkout is tried again on the next edit
        if (!g.d->path.empty()) startVcsQuery(true);
        break;
    case IDM_HIGHLIGHT:
        g.highlight = !g.highlight;
        saveSettings();
        InvalidateRect(g.view, nullptr, FALSE);
        InvalidateRect(g.diffPane, nullptr, FALSE);
        updateStatus();
        break;
    case IDM_LINENUMBERS:
        g.lineNumbers = !g.lineNumbers;
        refresh(true);
        break;
    case IDM_FONT: cmdFont(); break;
    case IDM_ZOOMIN: setZoom(g.fontPt + 1); break;
    case IDM_ZOOMOUT: setZoom(g.fontPt - 1); break;
    case IDM_ZOOMRESET: setZoom(11); break;
    case IDM_THEME_SYSTEM:
    case IDM_THEME_LIGHT:
    case IDM_THEME_DARK:
        g.themeMode = id - IDM_THEME_SYSTEM;
        applyTheme();
        saveSettings();
        break;
    case IDM_VCS_FILES: cmdChangelistFiles(); break;
    case IDM_ABOUT: DialogBoxParamW(g.inst, MAKEINTRESOURCEW(IDD_ABOUT), g.frame, aboutProc, 0); break;
    }
}

void updateMenu() {
    auto en = [](UINT id, bool on) { EnableMenuItem(g.menu, id, MF_BYCOMMAND | (on ? MF_ENABLED : MF_GRAYED)); };
    en(IDM_UNDO, g.d->doc.canUndo());
    en(IDM_REDO, g.d->doc.canRedo());
    en(IDM_CUT, true);
    en(IDM_COPY, true);
    en(IDM_DELETE, hasSel());
    en(IDM_PASTE, IsClipboardFormatAvailable(CF_UNICODETEXT) != 0);
    CheckMenuItem(g.menu, IDM_WRAP, MF_BYCOMMAND | (g.wrap ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(g.menu, IDM_LINENUMBERS, MF_BYCOMMAND | (g.lineNumbers ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(g.menu, IDM_HIGHLIGHT, MF_BYCOMMAND | (g.highlight ? MF_CHECKED : MF_UNCHECKED));
    en(IDM_WRAP, !g.diffOn);
    CheckMenuItem(g.menu, IDM_DIFF, MF_BYCOMMAND | (g.diffOn ? MF_CHECKED : MF_UNCHECKED));
    const bool p4 = g.d->vcs.kind == VcsKind::Perforce;
    const bool vcs = g.d->vcs.kind != VcsKind::None && !g.d->vcs.toolMissing;
    themeSetMenuText(g.menu, IDM_VCS_COMMIT, p4 || !vcs ? L"&Submit Current Changelist..." : L"&Commit All Changes...");
    en(IDM_VCS_COMMIT, vcs && (!p4 || g.d->vcs.opened));
    en(IDM_VCS_REVERTFILE, vcs && g.d->vcs.tracked);
    en(IDM_VCS_REFRESH, !g.d->path.empty());
    en(IDM_VCS_FILES, vcs);
    for (int k = 0; k < 3; k++)
        CheckMenuItem(g.menu, IDM_THEME_SYSTEM + k, MF_BYCOMMAND | (g.themeMode == k ? MF_CHECKED : MF_UNCHECKED));
}

// ====================================================================== view window

void onKeyDown(WPARAM vk) {
    const bool ctrl = GetKeyState(VK_CONTROL) < 0;
    const bool shift = GetKeyState(VK_SHIFT) < 0;
    LayoutCfg c = cfg();
    switch (vk) {
    case VK_LEFT:
        if (hasSel() && !shift && !ctrl) moveCaret(selMin(), false);
        else moveCaret(ctrl ? wordLeft(g.d->doc, g.d->caret, c) : prevChar(g.d->doc, g.d->caret, c), shift);
        break;
    case VK_RIGHT:
        if (hasSel() && !shift && !ctrl) moveCaret(selMax(), false);
        else moveCaret(ctrl ? wordRight(g.d->doc, g.d->caret, c) : nextChar(g.d->doc, g.d->caret, c), shift);
        break;
    case VK_UP:
        if (ctrl) {
            scrollRows(-1);
            refresh(false);
        } else {
            moveVertical(-1, shift, false);
        }
        break;
    case VK_DOWN:
        if (ctrl) {
            scrollRows(1);
            refresh(false);
        } else {
            moveVertical(1, shift, false);
        }
        break;
    case VK_PRIOR: moveVertical(-std::max(1, g.visRows - 1), shift, true); break;
    case VK_NEXT: moveVertical(std::max(1, g.visRows - 1), shift, true); break;
    case VK_HOME:
        if (ctrl) {
            moveCaret(0, shift);
        } else {
            // First press goes to the first non-blank character, second to column 0.
            Row r = rowContaining(g.d->doc, g.d->caret, c);
            u64 p = r.start;
            while (p < r.end && (g.d->doc.at(p) == ' ' || g.d->doc.at(p) == '\t')) p++;
            moveCaret(g.d->caret == p ? r.start : p, shift);
        }
        break;
    case VK_END:
        if (ctrl) moveCaret(g.d->doc.size(), shift);
        else moveCaret(rowContaining(g.d->doc, g.d->caret, c).end, shift);
        break;
    case VK_BACK:
        if (hasSel()) {
            deleteSel();
        } else if (g.d->caret > 0) {
            u64 p = ctrl ? wordLeft(g.d->doc, g.d->caret, c) : prevChar(g.d->doc, g.d->caret, c);
            g.d->doc.erase(p, g.d->caret - p, !ctrl);
            afterEdit(p);
        }
        break;
    case VK_DELETE:
        if (hasSel()) {
            deleteSel();
        } else if (g.d->caret < g.d->doc.size()) {
            u64 p = ctrl ? wordRight(g.d->doc, g.d->caret, c) : nextChar(g.d->doc, g.d->caret, c);
            g.d->doc.erase(g.d->caret, p - g.d->caret, !ctrl);
            afterEdit(g.d->caret);
        }
        break;
    case VK_INSERT:
        if (ctrl) doCopy();
        else if (shift) doPaste();
        break;
    }
}

void onChar(wchar_t ch) {
    if (ch == L'\r') {
        typeNewline();
        return;
    }
    if (ch == L'\t') {
        // Shift+Tab outdents; Tab with a multi-line selection indents it.
        const bool shift = GetKeyState(VK_SHIFT) < 0;
        u64 l0, l1;
        if (shift || (hasSel() && selLines(l0, l1) && l1 > l0)) indentLines(shift);
        else typeChar('\t');
        return;
    }
    if (ch < 0x20 || ch == 0x7F) return;  // other control codes arrive via accelerators
    if (ch >= 0xD800 && ch <= 0xDBFF) {
        g.pendingHigh = ch;
        return;
    }
    u32 cp = ch;
    if (ch >= 0xDC00 && ch <= 0xDFFF) {
        if (!g.pendingHigh) return;
        cp = 0x10000 + (((u32)g.pendingHigh - 0xD800) << 10) + ((u32)ch - 0xDC00);
    }
    g.pendingHigh = 0;
    typeChar(cp);
}

void dragTo(int x, int y) {
    if (y < 0) scrollRows(-1);
    else if (y >= g.visRows * g.ch) scrollRows(1);
    relayout();
    g.d->caret = posFromPoint(x, y);
    g.d->wantCol = -1;
    refresh(false);
}

LRESULT CALLBACK viewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paintView(hwnd); return 0;
    case WM_SIZE:
        if (g.view) refresh(false);
        return 0;
    case WM_SETFOCUS: {
        g.focused = true;
        g.caretOn = true;
        UINT t = GetCaretBlinkTime();
        if (t != INFINITE && t != 0) SetTimer(hwnd, kTimerBlink, t, nullptr);
        InvalidateRect(hwnd, nullptr, FALSE);
        return 0;
    }
    case WM_KILLFOCUS:
        g.focused = false;
        KillTimer(hwnd, kTimerBlink);
        InvalidateRect(hwnd, nullptr, FALSE);
        return 0;
    case WM_TIMER:
        if (wp == kTimerBlink) {
            g.caretOn = !g.caretOn;
            InvalidateRect(hwnd, nullptr, FALSE);
        } else if (wp == kTimerDrag && g.selecting) {
            POINT pt;
            GetCursorPos(&pt);
            ScreenToClient(hwnd, &pt);
            if (pt.y < 0 || pt.y >= g.visRows * g.ch) dragTo(pt.x, pt.y);
        }
        return 0;
    case WM_KEYDOWN: onKeyDown(wp); return 0;
    case WM_CHAR: onChar((wchar_t)wp); return 0;
    case WM_LBUTTONDOWN: {
        SetFocus(hwnd);
        u64 p = posFromPoint(GET_X_LPARAM(lp), GET_Y_LPARAM(lp));
        g.selecting = true;
        SetCapture(hwnd);
        SetTimer(hwnd, kTimerDrag, 50, nullptr);
        moveCaret(p, (wp & MK_SHIFT) != 0);
        return 0;
    }
    case WM_LBUTTONDBLCLK: {
        u64 p = posFromPoint(GET_X_LPARAM(lp), GET_Y_LPARAM(lp));
        u64 a, b;
        wordAt(g.d->doc, p, a, b);
        g.d->anchor = a;
        g.d->caret = b;
        g.d->doc.breakCoalesce();
        refresh(false);
        return 0;
    }
    case WM_MOUSEMOVE:
        if (g.selecting && (wp & MK_LBUTTON)) dragTo(GET_X_LPARAM(lp), GET_Y_LPARAM(lp));
        return 0;
    case WM_LBUTTONUP:
    case WM_CAPTURECHANGED:
        if (g.selecting) {
            g.selecting = false;
            KillTimer(hwnd, kTimerDrag);
            if (msg == WM_LBUTTONUP) ReleaseCapture();
        }
        return 0;
    case WM_MOUSEWHEEL: {
        if (GET_KEYSTATE_WPARAM(wp) & MK_CONTROL) {
            setZoom(g.fontPt + (GET_WHEEL_DELTA_WPARAM(wp) > 0 ? 1 : -1));
            return 0;
        }
        g.wheelAcc += GET_WHEEL_DELTA_WPARAM(wp);
        UINT lines = 3;
        SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
        if (lines == WHEEL_PAGESCROLL || lines > 100) lines = (UINT)std::max(1, g.visRows - 1);
        // Accumulate, so fractional notches from precision touchpads add up.
        int rowsToScroll = (g.wheelAcc * (int)lines) / WHEEL_DELTA;
        if (rowsToScroll != 0) {
            g.wheelAcc -= rowsToScroll * WHEEL_DELTA / (int)lines;
            scrollRows(-rowsToScroll);
            refresh(false);
        }
        return 0;
    }
    case WM_VSCROLL: {
        switch (LOWORD(wp)) {
        case SB_LINEUP: scrollRows(-1); break;
        case SB_LINEDOWN: scrollRows(1); break;
        case SB_PAGEUP: scrollRows(-std::max(1, g.visRows - 1)); break;
        case SB_PAGEDOWN: scrollRows(std::max(1, g.visRows - 1)); break;
        case SB_TOP:
            g.d->top = 0;
            g.d->topFill = 1 << 30;  // clamped to the fillers above line 1
            break;
        case SB_BOTTOM:
            if (alignOn() && g.d->doc.linesKnown()) setTopRow(alignedTotal() > (u64)g.visRows ? alignedTotal() - g.visRows : 0);
            else g.d->top = g.d->doc.size();
            break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: {
            SCROLLINFO si = {};
            si.cbSize = sizeof(si);
            si.fMask = SIF_TRACKPOS;
            GetScrollInfo(hwnd, SB_VERT, &si);
            if (alignOn() && g.d->doc.linesKnown()) {
                setTopRow((u64)(unsigned)si.nTrackPos);
                break;
            }
            u64 off = (u64)(unsigned)si.nTrackPos << g.vshift;
            g.d->top = std::min(off, g.d->doc.size());
            g.d->topFill = 0;
            break;
        }
        default: return 0;
        }
        refresh(false);
        return 0;
    }
    case WM_HSCROLL: {
        if (g.wrap) return 0;
        switch (LOWORD(wp)) {
        case SB_LINELEFT: g.d->hcol -= 4; break;
        case SB_LINERIGHT: g.d->hcol += 4; break;
        case SB_PAGELEFT: g.d->hcol -= std::max(1, g.visCols - 4); break;
        case SB_PAGERIGHT: g.d->hcol += std::max(1, g.visCols - 4); break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: {
            SCROLLINFO si = {};
            si.cbSize = sizeof(si);
            si.fMask = SIF_TRACKPOS;
            GetScrollInfo(hwnd, SB_HORZ, &si);
            g.d->hcol = si.nTrackPos;
            break;
        }
        default: return 0;
        }
        if (g.d->hcol < 0) g.d->hcol = 0;
        refresh(false);
        return 0;
    }
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ====================================================================== frame window

LRESULT CALLBACK frameProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    LRESULT themed = 0;
    if (themeMenuBarMessage(hwnd, msg, wp, lp, &themed, g.pal)) return themed;
    if (msg == WM_MEASUREITEM && wp == 0) {
        themeMeasureMenuItem(hwnd, (MEASUREITEMSTRUCT*)lp, g.uiFont);
        return TRUE;
    }
    if (msg == WM_MENUCHAR && (HIWORD(wp) & MF_POPUP)) return themeMenuChar((HMENU)lp, (wchar_t)LOWORD(wp));
    if (msg == WM_DRAWITEM && wp == 0) {
        themeDrawMenuItem((DRAWITEMSTRUCT*)lp, g.uiFont, g.pal);
        return TRUE;
    }
    if (g.findMsg && msg == g.findMsg) {
        onFindMessage((LPFINDREPLACEW)lp);
        return 0;
    }
    switch (msg) {
    case WM_CREATE:
        g.frame = hwnd;
        g.dpi = (int)GetDpiForWindow(hwnd);
        makeFonts();
        g.view = CreateWindowExW(0, kViewClass, L"", WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL, 0, 0, 0, 0,
                                 hwnd, nullptr, g.inst, nullptr);
        if (g.wrap) ShowScrollBar(g.view, SB_HORZ, FALSE);
        g.diffPane = CreateWindowExW(0, kDiffClass, L"", WS_CHILD, 0, 0, 0, 0, hwnd, nullptr, g.inst, nullptr);
        g.strip = CreateWindowExW(0, kStripClass, L"", WS_CHILD, 0, 0, 0, 0, hwnd, nullptr, g.inst, nullptr);
        return 0;
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paintStatus(hwnd); return 0;
    case WM_SIZE:
        if (g.view) layoutFrame();
        return 0;
    case WM_SETFOCUS:
        if (g.view) SetFocus(g.view);
        return 0;
    case WM_DPICHANGED: {
        g.dpi = HIWORD(wp);
        makeFonts();
        const RECT* r = (const RECT*)lp;
        SetWindowPos(hwnd, nullptr, r->left, r->top, r->right - r->left, r->bottom - r->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
        layoutFrame();
        refresh(true);
        return 0;
    }
    case WM_INITMENUPOPUP: updateMenu(); return 0;
    case WM_MBUTTONUP: {
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        int i = tabAt(pt, nullptr);
        if (i >= 0) closeTab(i);
        return 0;
    }
    case WM_MOUSEMOVE: {
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        bool onX = false;
        int i = tabAt(pt, &onX);
        int x = onX ? i : -1;
        int nav = !g.diffOn ? 0 : PtInRect(&g.navPrev, pt) ? 1 : PtInRect(&g.navNext, pt) ? 2 : 0;
        if (i != g.tabHover || x != g.tabCloseHover || nav != g.navHover) {
            g.tabHover = i;
            g.tabCloseHover = x;
            g.navHover = nav;
            RECT r = {0, 0, 32767, tabBarH() + headerH()};
            InvalidateRect(hwnd, &r, FALSE);
            TRACKMOUSEEVENT tme = {sizeof(tme), TME_LEAVE, hwnd, 0};
            TrackMouseEvent(&tme);
        }
        return 0;
    }
    case WM_MOUSELEAVE:
        if (GetCapture() == hwnd) return 0;
        if (g.tabHover >= 0 || g.navHover) {
            g.tabHover = g.tabCloseHover = -1;
            g.navHover = 0;
            RECT r = {0, 0, 32767, tabBarH() + headerH()};
            InvalidateRect(hwnd, &r, FALSE);
        }
        return 0;
    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK: {
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        bool onX = false;
        int ti = tabAt(pt, &onX);
        if (ti >= 0 && !onX) {
            activateTab(ti);
            return 0;
        }
        // Buttons go in while held and act on release over the same button.
        if (ti >= 0 && onX) g.tabClosePress = ti;
        else if (g.diffOn && PtInRect(&g.navPrev, pt)) g.navPress = 1;
        else if (g.diffOn && PtInRect(&g.navNext, pt)) g.navPress = 2;
        else {
            SetFocus(g.view);
            return 0;
        }
        SetCapture(hwnd);
        RECT r = {0, 0, 32767, tabBarH() + headerH()};
        InvalidateRect(hwnd, &r, FALSE);
        return 0;
    }
    case WM_LBUTTONUP: {
        if (GetCapture() != hwnd) return 0;
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        const int closeTabIdx = g.tabClosePress, nav = g.navPress;
        g.tabClosePress = -1;
        g.navPress = 0;
        ReleaseCapture();
        bool onX = false;
        int ti = tabAt(pt, &onX);
        if (closeTabIdx >= 0 && ti == closeTabIdx && onX) closeTab(closeTabIdx);
        else if (nav == 1 && PtInRect(&g.navPrev, pt)) gotoChange(false);
        else if (nav == 2 && PtInRect(&g.navNext, pt)) gotoChange(true);
        SetFocus(g.view);
        RECT r = {0, 0, 32767, tabBarH() + headerH()};
        InvalidateRect(hwnd, &r, FALSE);
        return 0;
    }
    case WM_CAPTURECHANGED:
        if (g.tabClosePress >= 0 || g.navPress) {
            g.tabClosePress = -1;
            g.navPress = 0;
            RECT r = {0, 0, 32767, tabBarH() + headerH()};
            InvalidateRect(hwnd, &r, FALSE);
        }
        return 0;
    case WM_COMMAND:
        if (lp == 0 || HIWORD(wp) <= 1) onCommand(LOWORD(wp));
        return 0;
    case WM_DROPFILES: {
        HDROP drop = (HDROP)wp;
        wchar_t file[32768];
        std::vector<std::wstring> files;
        UINT n = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
        for (UINT k = 0; k < n; k++)
            if (DragQueryFileW(drop, k, file, 32768)) files.push_back(file);
        DragFinish(drop);
        SetForegroundWindow(hwnd);
        for (const std::wstring& f : files) openInTab(f);
        return 0;
    }
    case WM_APP_INDEX:
        for (auto& t : g.tabs)
            if ((unsigned)wp == t->indexGen && t->index && t->index->ready()) {
                t->indexing = false;
                t->doc.indexReady();
                if (t.get() == g.d) {
                    KillTimer(hwnd, kTimerIndex);
                    refresh(false);
                }
            }
        return 0;
    case WM_APP_VCS: onVcsResult((unsigned)wp, (VcsInfo*)lp); return 0;
    case WM_APP_P4ASK: onP4Ask(); return 0;
    case WM_APP_P4EDITED: onP4Edited((unsigned)wp, (RunResult*)lp); return 0;
    case WM_UPDATE_STATUS:
        g.updatePct = (int)(INT_PTR)wp;
        updateStatus();
        return 0;
    case WM_UPDATE_DONE: {
        wchar_t* m = (wchar_t*)lp;
        if (m) {
            MessageBoxW(hwnd, m, kAppName, MB_OK | (wp == UPDATE_FAILED ? MB_ICONWARNING : MB_ICONINFORMATION));
            Updater_FreeString(m);
        }
        Updater_Reset();
        return 0;
    }
    case WM_UPDATE_READY: onUpdateReady((wchar_t*)lp); return 0;
    case WM_TIMER:
        if (wp == kTimerSession) {
            autosaveSession();
            return 0;
        }
        if (wp == kTimerIndex) {
            if (g.d->indexing) updateStatus();
            else KillTimer(hwnd, kTimerIndex);
        } else if (wp == kTimerDiff) {
            KillTimer(hwnd, kTimerDiff);
            recomputeDiff();
            relayout();
            InvalidateRect(g.view, nullptr, FALSE);
            if (g.diffOn) {
                InvalidateRect(g.diffPane, nullptr, FALSE);
                InvalidateRect(g.strip, nullptr, FALSE);
                RECT hd = {0, 0, 32767, tabBarH() + headerH()};
                InvalidateRect(hwnd, &hd, FALSE);
            }
            updateStatus();
        }
        return 0;
    case WM_CLOSE:
        // Nothing to ask: unsaved work is kept in the session (see saveSession).
        saveSession();
        saveSettings();
        DestroyWindow(hwnd);
        return 0;
    case WM_SETTINGCHANGE:
        // Windows light/dark switch, when following the system setting.
        if (g.themeMode == 0 && lp && lstrcmpiW((LPCWSTR)lp, L"ImmersiveColorSet") == 0) applyTheme();
        return DefWindowProcW(hwnd, msg, wp, lp);
    case WM_DESTROY:
        for (auto& t : g.tabs) t->index.reset();
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

}  // namespace

int WINAPI wWinMain(HINSTANCE inst, HINSTANCE, PWSTR, int show) {
    g.inst = inst;
    loadSettings();
    INITCOMMONCONTROLSEX icc = {sizeof(icc), ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES | ICC_LINK_CLASS};
    InitCommonControlsEx(&icc);
    themeInit();
    gdiplusStart();
    g.pal = makePalette(wantDark());
    themeSetAppDark(g.pal.dark);

    WNDCLASSEXW fc = {};
    fc.cbSize = sizeof(fc);
    fc.lpfnWndProc = frameProc;
    fc.hInstance = inst;
    fc.hIcon = LoadIconW(inst, MAKEINTRESOURCEW(IDI_APP));
    fc.hIconSm = fc.hIcon;
    fc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    fc.lpszClassName = kFrameClass;
    RegisterClassExW(&fc);

    WNDCLASSEXW vc = {};
    vc.cbSize = sizeof(vc);
    vc.style = CS_DBLCLKS;
    vc.lpfnWndProc = viewProc;
    vc.hInstance = inst;
    vc.hCursor = LoadCursorW(nullptr, IDC_IBEAM);
    vc.lpszClassName = kViewClass;
    RegisterClassExW(&vc);

    WNDCLASSEXW dcl = {};
    dcl.cbSize = sizeof(dcl);
    dcl.lpfnWndProc = diffProc;
    dcl.hInstance = inst;
    dcl.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    dcl.lpszClassName = kDiffClass;
    RegisterClassExW(&dcl);
    dcl.lpfnWndProc = stripProc;
    dcl.lpszClassName = kStripClass;
    RegisterClassExW(&dcl);

    g.menu = buildMenu();
    themeOwnerDrawMenu(g.menu);
    for (int k = 0; k < 3; k++) themeSetMenuRadio(g.menu, IDM_THEME_SYSTEM + k);
    g.accel = buildAccelerators();
    g.findMsg = RegisterWindowMessageW(FINDMSGSTRINGW);

    int x = CW_USEDEFAULT, y = CW_USEDEFAULT, w = CW_USEDEFAULT, h = CW_USEDEFAULT;
    if (g_savedRect.right > 0) {
        x = g_savedRect.left;
        y = g_savedRect.top;
        w = g_savedRect.right;
        h = g_savedRect.bottom;
        // Ignore a saved position that is no longer on any monitor.
        RECT probe = {x, y, x + w, y + h};
        if (!MonitorFromRect(&probe, MONITOR_DEFAULTTONULL)) x = y = CW_USEDEFAULT;
    }
    CreateWindowExW(WS_EX_ACCEPTFILES, kFrameClass, kAppName, WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN, x, y, w, h,
                    nullptr, g.menu, inst, nullptr);
    if (!g.frame || !g.view) return 1;

    // Load the document before the window is shown, so the first frame the
    // user sees already has the text in it.
    addTab();
    restoreSession();
    for (int i = 1; i < __argc; i++)
        if (__wargv[i][0]) openInTab(__wargv[i]);

    applyTheme();
    g_lastSessionStamp = sessionStamp();
    SetTimer(g.frame, kTimerSession, 5000, nullptr);
    ShowWindow(g.frame, (g_savedMax && show == SW_SHOWNORMAL) ? SW_SHOWMAXIMIZED : show);
    UpdateWindow(g.frame);
    SetFocus(g.view);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        if (g.findDlg && IsDialogMessageW(g.findDlg, &msg)) continue;
        if (TranslateAcceleratorW(g.frame, g.accel, &msg)) continue;
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    gdiplusStop();
    return (int)msg.wParam;
}
