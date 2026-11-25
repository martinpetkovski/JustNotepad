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
#include <functional>
#include <cmath>
#include <ctime>
#include <ctype.h>
#include <memory>
#include <string>
#include <thread>
#include <vector>

#include "../core/diff.h"
#include "../core/csv.h"
#include "../core/datatree.h"
#include "../core/highlight.h"
#include "../core/markdown.h"
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

// Defined below, outside this file's private namespace (next to wWinMain).
void onOpenFromOtherCopy(const COPYDATASTRUCT* cd);

namespace {

const wchar_t* const kAppName = L"JustNotepad";
const wchar_t* const kFrameClass = L"JustNotepadFrame";
const wchar_t* const kViewClass = L"JustNotepadView";
const wchar_t* const kDiffClass = L"JustNotepadDiff";
const wchar_t* const kStripClass = L"JustNotepadDiffStrip";
const wchar_t* const kPreviewClass = L"JustNotepadPreview";
const wchar_t* const kSplitClass = L"JustNotepadSplitter";
const wchar_t* const kHexClass = L"JustNotepadHex";
const wchar_t* const kCsvClass = L"JustNotepadTable";

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
           WM_APP_P4EDITED = WM_APP + 4, WM_APP_BLAME = WM_APP + 5, WM_APP_DISKCHECK = WM_APP + 6;
const UINT_PTR kTimerBlink = 1, kTimerIndex = 2, kTimerDrag = 3, kTimerDiff = 4, kTimerSession = 5, kTimerSide = 6,
                kTimerDisk = 7;

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

    // Who changed what: blame of the latest revision (fetched when the diff
    // opens), and how the edited text differs from that revision.
    Blame blame;
    bool blameLoaded = false, blamePending = false;
    std::wstring blameError;
    unsigned blameGen = 0;
    std::vector<u64> headHash;   // per line of the blamed revision
    std::vector<Hunk> localHunks; // blamed revision -> current text
    u64 localVersion = NPOS;

    // Diff against a chosen revision (File History) instead of #have / HEAD.
    std::wstring baseSpec, baseLabel;

    bool binary = false;    // looks like binary data: the hex view opens with it
    bool forceHex = false;  // View > Show as Hex

    // Session backup of unsaved text (see saveSession).
    std::wstring backupName;
    u64 backupVersion = NPOS;

    // Extra carets / selections (Ctrl+Alt+click, Shift+Alt+., Alt+drag). The
    // main one stays in caret/anchor.
    // `pad`: the caret sits that many columns past the end of its line
    // (column selection over short lines); typing there fills in spaces first.
    struct Sel { u64 caret, anchor; int pad = 0; };
    std::vector<Sel> extra;
    int caretPad = 0;  // the same for the main caret

    // The file on disk as we last read or wrote it, to notice outside changes.
    FILETIME diskTime = {};
    u64 diskSize = 0;
    bool diskKnown = false;

    DocState() {
        // Keep extra carets on the same text while it changes.
        doc.onEdit = [this](u64 pos, u64 removed, u64 inserted) {
            auto map = [&](u64 x) { return x < pos ? x : x < pos + removed ? pos : x - removed + inserted; };
            for (Sel& e : extra) {
                e.caret = map(e.caret);
                e.anchor = map(e.anchor);
            }
        };
    }
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

    // side pane (data view / Markdown preview)
    bool sideOn = true;              // on by default: JSON, XML and Markdown open with it
    HWND sideTree = nullptr, sidePreview = nullptr;
    DataTree data;                   // what the tree shows
    std::vector<bool> dataFilled;    // nodes whose children are in the tree
    const DocState* dataDoc = nullptr;
    u64 dataVersion = NPOS;
    bool treeUpdating = false;
    int mdScroll = 0;
    u64 mdSyncedTop = NPOS;
    bool sideShown = false;          // what layoutFrame last laid out
    // Movable dividers: share of the width for the diff's left side and for
    // the side pane; dragging one (1 = diff strip, 2 = side divider).
    double diffSplit = 0.5, sideSplit = 0.42;
    int dragSplit = 0, dragOffset = 0;
    RECT sideSplitRect = {};
    HWND sideSplitter = nullptr;
    HWND sideHex = nullptr, sideCsv = nullptr;
    int hexShift = 0;
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

    // editor extras
    bool showWs = false;             // View > Show Whitespace
    bool boxActive = false, boxDrag = false, altClicked = false;
    u64 boxL0 = 0, boxL1 = 0;        // column (box) selection: lines and columns of its corners
    int boxC0 = 0, boxC1 = 0;
    std::vector<std::wstring> recent;  // File > Open Recent, newest first
    HMENU recentMenu = nullptr, encMenu = nullptr, eolMenu = nullptr;

    std::wstring statusLeft, statusRight, title;
    u64 gotoLine = 0;
};

App g;
unsigned g_gen = 0;  // source of unique ids for background work replies

INT_PTR dialogTheme(HWND dlg, UINT msg, WPARAM wp, LPARAM lp);
void scrollRows(int n);
void syncHexToEditor();
void hexScrollbar();
void syncCsvToEditor();
void saveSettings();
std::wstring formatTime(long long t);
bool findBarOpen();
void updateFindStatus(u64 hit);
void paintRowMatches(HDC dc, const Row& r, int y, const LayoutCfg& c);
void closeFindBar();
void showEditorMenu(POINT screen);
POINT caretPoint();
void onCommand(int id);
void layoutFrame();
bool sideVisible();
void updateSidePane();
void syncPreviewToEditor();
void themeSidePane();
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
bool multi();
void clearExtras();
void multiInsert(const std::string& t);
bool multiKey(WPARAM vk, bool ctrl, bool shift);
bool multiCopy();
void multiPaste(std::string s);
std::vector<std::pair<u64, u64>> selRanges();
void editEach(const std::function<void(DocState::Sel&, size_t)>& f);
void eraseSel(DocState::Sel& s);
std::string newlineText(u64 at, size_t& caretOff);
bool outdentCloser(u32 cp);
bool bracketPair(u64& a, u64& b);
void noteDiskState();
void addRecent(const std::wstring& path);
void rebuildRecentMenu();
void checkDisk(bool mayAsk);
void addCaretAt(u64 p);
void boxRebuild();
bool boxPoint(int x, int y, u64& line, int& col);
void showEncodingMenu(POINT screen);
void detectLang();
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
    if (multi()) {
        swprintf(b, 256, L"    %llu carets", (unsigned long long)g.d->extra.size() + 1);
        left += b;
    } else if (hasSel()) {
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
    std::wstring zoom;
    if (g.fontPt != 11) {
        swprintf(b, 256, L"%d%%      ", (g.fontPt * 100 + 5) / 11);
        zoom = b;
    }
    std::wstring right = upd + vcsText() + langText + zoom + humanSize(g.d->doc.size()) + L"      " + encNames[g.d->enc] + L"      " +
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
    if (g.sideTree) {
        const bool want = sideVisible();
        if (want != g.sideShown) {
            g.sideShown = want;
            layoutFrame();
            updateSidePane();
        } else if (want && g.dataDoc == g.d && g.dataVersion != g.d->doc.version()) {
            SetTimer(g.frame, kTimerSide, 400, nullptr);  // the tree follows once typing pauses
        }
    }
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
    if (g.sideShown && g.sidePreview && IsWindowVisible(g.sidePreview)) {
        InvalidateRect(g.sidePreview, nullptr, FALSE);  // laid out again if the text changed
        syncPreviewToEditor();
    }
    if (g.sideShown && g.sideHex && IsWindowVisible(g.sideHex)) {
        syncHexToEditor();
        hexScrollbar();
        InvalidateRect(g.sideHex, nullptr, FALSE);
    }
    if (g.sideShown && g.sideCsv && IsWindowVisible(g.sideCsv)) {
        syncCsvToEditor();
        InvalidateRect(g.sideCsv, nullptr, FALSE);  // re-read if the text changed
    }
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

// ---- blame: who made each change

void startBlame() {
    DocState* t = g.d;
    if (!g.diffOn || t->blameLoaded || t->blamePending || t->path.empty() || t->vcs.kind == VcsKind::None ||
        t->vcs.toolMissing || !t->vcs.tracked)
        return;
    t->blamePending = true;
    unsigned gen = t->blameGen = ++g_gen;
    std::wstring path = t->path;
    VcsKind kind = t->vcs.kind;
    HWND frame = g.frame;
    std::thread([path, kind, gen, frame]() {
        Blame* b = new Blame();
        std::wstring err;
        if (!blameFile(path, kind, *b, err)) {
            b->line.clear();
            b->infos.clear();
            BlameInfo e;
            e.summary = err.empty() ? L"blame failed" : err;  // carried to the UI in the first info
            b->infos.push_back(e);
        }
        if (!PostMessageW(frame, WM_APP_BLAME, (WPARAM)gen, (LPARAM)b)) delete b;
    }).detach();
}

void onBlame(unsigned gen, Blame* raw) {
    std::unique_ptr<Blame> b(raw);
    for (auto& tp : g.tabs) {
        DocState* t = tp.get();
        if (t->blameGen != gen) continue;
        t->blamePending = false;
        t->blameLoaded = true;
        t->blame = std::move(*b);
        t->blameError = t->blame.line.empty() && !t->blame.infos.empty() ? t->blame.infos[0].summary : L"";
        LineHasher h;
        h.feed((const u8*)t->blame.text.data(), t->blame.text.size());
        h.finish();
        t->headHash.swap(h.lines);
        t->localVersion = NPOS;
        if (t == g.d) InvalidateRect(g.view, nullptr, FALSE);
        return;
    }
}

std::wstring ago(long long t) {
    if (t <= 0) return L"";
    long long now = (long long)time(nullptr), d = now - t;
    wchar_t b[48];
    if (d < 3600) swprintf(b, 48, L"%lld min ago", std::max(1LL, d / 60));
    else if (d < 86400) swprintf(b, 48, L"%lld hour%ls ago", d / 3600, d / 3600 == 1 ? L"" : L"s");
    else if (d < 86400 * 45) swprintf(b, 48, L"%lld day%ls ago", d / 86400, d / 86400 == 1 ? L"" : L"s");
    else return formatTime(t).substr(0, 10);
    return b;
}

// Who made a change: name on one line, "change 1234 · 3 days ago" on the next.
struct Owner {
    std::wstring who, detail;
};

Owner changeOwner(const Hunk& h) {
    const bool p4 = g.d->vcs.kind == VcsKind::Perforce;
    const Owner yours = {L"You", p4 ? L"not submitted" : L"not committed"};
    // Against the latest revision every difference is your own edit.
    if (g.d->baseSpec.empty()) return yours;
    if (h.bCount == 0) return {L"Removed", L"after that revision"};
    if (!g.d->blameLoaded) return {L"\x2026", L""};
    if (g.d->blame.line.empty()) {  // say why, rather than nothing
        std::wstring e = g.d->blameError;
        size_t nl = e.find_first_of(L"\r\n");
        return {L"Unknown", nl == std::wstring::npos ? e : e.substr(0, nl)};
    }
    if (g.d->localVersion != g.d->doc.version() && g.d->diffDocVersion == g.d->doc.version()) {
        g.d->localHunks = diffLines(g.d->headHash, g.d->curHash);  // what you changed since the latest revision
        g.d->localVersion = g.d->doc.version();
    }
    // Count owners over the change's lines (a few thousand at most).
    std::vector<int> count(g.d->blame.infos.size() + 1, 0);  // last slot: you
    const u32 n = std::min<u32>(h.bCount, 4000);
    for (u32 k = 0; k < n; k++) {
        const u32 line = h.bStart + k;
        int hi = hunkAtB(g.d->localHunks, line);
        if (hi >= 0 && g.d->localHunks[(size_t)hi].bCount > 0) {
            count.back()++;
            continue;
        }
        i64 headLine = mapBtoA(g.d->localHunks, line);
        if (headLine >= 0 && (u64)headLine < g.d->blame.line.size()) {
            int who = g.d->blame.line[(size_t)headLine];
            if (who >= 0) count[(size_t)who]++;
        }
    }
    int best = -1, owners = 0;
    for (size_t i = 0; i < count.size(); i++) {
        if (!count[i]) continue;
        owners++;
        if (best < 0 || count[i] > count[(size_t)best]) best = (int)i;
    }
    if (best < 0) return {L"", L""};
    Owner o;
    if ((size_t)best == count.size() - 1) {
        o = yours;
    } else {
        const BlameInfo& bi = g.d->blame.infos[(size_t)best];
        o.who = (p4 ? L"#" + bi.change : bi.change.substr(0, 7)) + L" " + (bi.who.empty() ? L"?" : bi.who);
        o.detail = ago(bi.time);
    }
    if (owners > 1) {
        wchar_t b[32];
        swprintf(b, 32, L" +%d", owners - 1);
        o.who += b;
    }
    return o;
}


// View > Show Whitespace: a dot for each space, an arrow for each tab.
void paintWhitespace(HDC dc, const Row& r, int y, const LayoutCfg& c) {
    const u64 n = std::min<u64>(r.end - r.start, 16384);
    const std::string b = g.d->doc.text(r.start, n);
    SetTextColor(dc, blend(g.pal.fg, g.pal.bg, 30));
    const int maxCol = g.d->hcol + g.visCols + 1;
    int col = 0;
    u64 i = 0;
    while (i < n && col < maxCol) {
        u32 cp;
        const int len = decodeChar((const u8*)b.data() + i, n - i, c.ansi, cp);
        const int w = cellsFor(cp, col, c.tabw);
        if ((cp == ' ' || cp == '\t') && col >= g.d->hcol) {
            const int x = g.gutterPx + (col - g.d->hcol) * g.cw;
            RECT rc = {x, y, x + (cp == ' ' ? g.cw : std::max(g.cw, w * g.cw)), y + g.ch};
            DrawTextW(dc, cp == ' ' ? L"\x00B7" : L"\x2192", 1, &rc,
                      (cp == ' ' ? DT_CENTER : DT_LEFT) | DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_NOCLIP);
        }
        col += w;
        i += (u64)std::max(1, len);
    }
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

    const bool linesOk = g.d->doc.linesKnown();
    // Every selection and caret (more than one with multiple carets).
    std::vector<std::pair<u64, u64>> selList = multi() ? selRanges() : std::vector<std::pair<u64, u64>>{{selMin(), selMax()}};
    std::vector<std::pair<u64, int>> caretList{{g.d->caret, g.d->caretPad}};
    for (auto& x : g.d->extra) caretList.push_back({x.caret, x.pad});
    u64 brA = NPOS, brB = NPOS;  // the bracket at the caret and its partner
    if (!hasSel() && !multi() && !bracketPair(brA, brB)) brA = brB = NPOS;
    if (g.diffOn) startBlame();
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
        paintRowMatches(dc, r, y, c);
        for (const auto& se : selList) {
            const u64 s = se.first, e = se.second;
            if (!(s < e && e > r.start && s < r.next)) continue;
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
        // Column selection: the whole rectangle, the empty space past short lines included.
        if (g.boxActive && g.boxC0 != g.boxC1 && startsLine && linesOk) {
            const u64 ln = lineNo - 1;
            if (ln >= std::min(g.boxL0, g.boxL1) && ln <= std::max(g.boxL0, g.boxL1)) {
                const int x0 = std::max(g.gutterPx, g.gutterPx + (std::min(g.boxC0, g.boxC1) - g.d->hcol) * g.cw);
                const int x1 = std::min(g.cliW, g.gutterPx + (std::max(g.boxC0, g.boxC1) - g.d->hcol) * g.cw);
                if (x1 > x0) fillRect(dc, x0, y, x1, y + g.ch, selBg);
            }
        }
        // Matching brackets: a light square around each.
        for (u64 bp : {brA, brB}) {
            if (bp == NPOS || bp < r.start || bp >= r.end) continue;
            const int x = g.gutterPx + (colOf(g.d->doc, r, bp, c) - g.d->hcol) * g.cw;
            if (x < g.gutterPx || x >= g.cliW) continue;
            const COLORREF edge = blend(fg, bg, 45);
            fillRect(dc, x, y, x + g.cw, y + g.ch, blend(fg, bg, 12));
            fillRect(dc, x, y, x + g.cw, y + 1, edge);
            fillRect(dc, x, y + g.ch - 1, x + g.cw, y + g.ch, edge);
            fillRect(dc, x, y, x + 1, y + g.ch, edge);
            fillRect(dc, x + g.cw - 1, y, x + g.cw, y + g.ch, edge);
        }
        SetTextColor(dc, fg);
        drawRowText(dc, g.d->doc, r, g.gutterPx, g.cliW, y, g.d->hcol, g.visCols, c,
                    hlOn() && linesOk ? hp.row(lineNo - 1, r.start, r.end) : nullptr);
        if (g.showWs && r.end > r.start) paintWhitespace(dc, r, y, c);

        for (const auto& cp : caretList) {
            const u64 cpos = cp.first;
            if (!(g.focused && g.caretOn && rowHas(g.d->doc, r, cpos))) continue;
            int col = colOf(g.d->doc, r, cpos, c) + cp.second;
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
int stripW() { return MulDiv(112, g.dpi, 96); }
int splitW() { return MulDiv(8, g.dpi, 96); }

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
    // Follow the strip where it really is (the middle can be dragged).
    RECT strip;
    GetWindowRect(g.strip, &strip);
    MapWindowPoints(nullptr, g.frame, (POINT*)&strip, 2);
    RECT viewR;
    GetWindowRect(g.view, &viewR);
    MapWindowPoints(nullptr, g.frame, (POINT*)&viewR, 2);
    const int H = headerH(), sw = stripW(), left = std::max(0, (int)strip.left);
    const int editorRight = viewR.right;  // the side pane may follow the editor
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
    if (!g.d->baseLabel.empty()) base = g.d->baseLabel;  // a revision picked in File History
    else if (g.d->vcs.kind == VcsKind::Perforce) base = L"Depot  " + (g.d->vcs.baseFrom.empty() ? g.d->vcs.label : g.d->vcs.baseFrom);
    else if (g.d->vcs.kind == VcsKind::Git) base = L"Committed  HEAD (" + g.d->vcs.label + L")";
    else base = L"Base";
    side(0, left, kDepotCol, base);
    side(left + sw, editorRight, kYoursCol, std::wstring(L"Yours  ") + (g.d->doc.modified() ? L"(unsaved)  " : L"") + g.d->path);

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
    if (!IsRectEmpty(&g.sideSplitRect)) {
        const RECT& sr = g.sideSplitRect;
        fillRect(dc, sr.left, sr.top, sr.right, sr.bottom, g.pal.face);
        fillRect(dc, sr.left, sr.top, sr.left + 1, sr.bottom, g.pal.dark ? blend(RGB(255, 255, 255), g.pal.face, 18) : RGB(255, 255, 255));
        fillRect(dc, sr.right - 1, sr.top, sr.right, sr.bottom, g.pal.border);
        // A grip in the middle: three small raised squares.
        const int cx = (sr.left + sr.right) / 2, cy = (sr.top + sr.bottom) / 2, k = std::max(2, MulDiv(2, g.dpi, 96));
        for (int n = -1; n <= 1; n++) {
            int y = cy + n * 3 * k;
            fillRect(dc, cx - k / 2, y, cx - k / 2 + k, y + k, blend(g.pal.faceText, g.pal.face, 45));
        }
    }
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
    g.boxActive = false;
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

// Enter: a line break with the indent (see newlineText), at every caret.
void typeNewline() {
    editEach([](DocState::Sel& s, size_t) {
        s.pad = 0;
        eraseSel(s);
        const u64 at = s.caret;
        size_t off = 0;
        const std::string t = newlineText(at, off);
        g.d->doc.insert(at, (const u8*)t.data(), t.size());
        s.caret = s.anchor = at + off;
    });
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
    if (multi()) {
        multiInsert(std::string((const char*)b, (size_t)n));
        return;
    }
    if (outdentCloser(cp)) return;
    replaceSel(b, n, true);
}

void deleteSel() {
    if (multi()) {
        multiInsert("");
        return;
    }
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
    if (multi() && multiCopy()) return;
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
    if (multi()) {
        multiPaste(s);
        return;
    }
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
    std::wstring path = g.d->path, spec = g.d->baseSpec;
    HWND frame = g.frame;
    std::thread([path, spec, gen, frame]() {
        VcsInfo* r = new VcsInfo(queryVcs(path, (size_t)kDiffMax));
        if (!spec.empty() && r->kind != VcsKind::None) {
            // A revision picked in File History is the base instead.
            std::vector<u8> bytes;
            std::wstring err;
            if (fetchRevision(path, r->kind, spec, (size_t)kDiffMax, bytes, err)) {
                r->base.swap(bytes);
                r->haveBase = true;
                r->tracked = true;
                r->baseError.clear();
            } else {
                r->base.clear();
                r->haveBase = false;
                r->tracked = true;
                r->baseError = err;
            }
            r->baseFrom = spec;
        }
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
    // The latest revision may have moved (sync, submit): blame it again when needed.
    g.d->blameLoaded = false;
    g.d->localVersion = NPOS;
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
    // Binary: a NUL byte near the start (text files do not have them).
    g.d->binary = len && memchr(text, 0, (size_t)std::min<u64>(len, 64u << 10)) != nullptr;
    g.d->extra.clear();
    g.boxActive = false;
    noteDiskState();
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
    g.d->extra.clear();
    g.d->diskKnown = false;
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
    noteDiskState();
    addRecent(path);
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
    updateSidePane();
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
            // Who made the change, right above its buttons.
            if (g.d->vcs.kind != VcsKind::None) {
                const Owner o = changeOwner({sl.aStart, sl.aCount, sl.bStart, sl.bCount});
                if (!o.who.empty()) {
                    HGDIOBJ of = SelectObject(dc, g.uiFont);
                    TEXTMETRICW tm;
                    GetTextMetricsW(dc, &tm);
                    const int lh = tm.tmHeight, px = MulDiv(1, g.dpi, 96), lines = o.detail.empty() ? 1 : 2;
                    int ty = cy - bs / 2 - 3 * px - lines * lh;
                    SetBkMode(dc, TRANSPARENT);
                    RECT r1 = {2 * px, ty, W - 2 * px, ty + lh};
                    SetTextColor(dc, fg);
                    DrawTextW(dc, o.who.c_str(), -1, &r1, DT_CENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
                    if (lines == 2) {
                        RECT r2 = {2 * px, ty + lh, W - 2 * px, ty + 2 * lh};
                        SetTextColor(dc, blend(fg, bg, 60));
                        DrawTextW(dc, o.detail.c_str(), -1, &r2, DT_CENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
                    }
                    SelectObject(dc, of);
                }
            }
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
        if (g.dragSplit == 1) {  // moving the middle: where the strip's left edge goes
            POINT f = pt;
            MapWindowPoints(hwnd, g.frame, &f, 1);
            RECT rc;
            GetClientRect(g.frame, &rc);
            const int avail = (IsRectEmpty(&g.sideSplitRect) ? (int)rc.right : (int)g.sideSplitRect.left) - stripW();
            if (avail > 0) g.diffSplit = std::clamp((double)(f.x - g.dragOffset) / avail, 0.05, 0.95);
            layoutFrame();
            UpdateWindow(g.frame);
            return 0;
        }
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
        // Anywhere else the strip is the divider between the two sides.
        SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
        return TRUE;
    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK: {
        // Pushed in while held; the choice applies on release, over the same button.
        if (g.hoverHunk >= 0) {
            g.pressHunk = g.hoverHunk;
            g.pressChoice = g.hoverChoice;
            SetCapture(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
        } else if (msg == WM_LBUTTONDBLCLK) {  // double click: both sides equal again
            g.diffSplit = 0.5;
            layoutFrame();
            saveSettings();
        } else {
            g.dragSplit = 1;
            g.dragOffset = GET_X_LPARAM(lp);
            SetCapture(hwnd);
        }
        return 0;
    }
    case WM_LBUTTONUP: {
        if (GetCapture() != hwnd) return 0;
        if (g.dragSplit == 1) {
            g.dragSplit = 0;
            ReleaseCapture();
            saveSettings();
            SetFocus(g.view);
            return 0;
        }
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
        if (g.dragSplit == 1) g.dragSplit = 0;
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
// After a submit, commit or revert: every open tab asks again (the other
// files of the changelist are no longer checked out, revisions moved on).
void refreshAllTabsVcs() {
    DocState* keep = g.d;
    for (auto& t : g.tabs) {
        if (t->path.empty()) continue;
        g.d = t.get();
        g.d->baseSpec.clear();  // "diff against #3" made sense before the submit, not after
        g.d->baseLabel.clear();
        g.d->p4Asked = false;
        startVcsQuery(true);
    }
    g.d = keep;
    refresh(false);
    invalidateTabs();
}

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
        refreshAllTabsVcs();
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
    refreshAllTabsVcs();  // the other files of the changelist, too
}

// ---- File History: every revision, diff against any, get any

std::vector<Revision> g_history;

std::wstring formatTime(long long t) {
    if (t <= 0) return L"";
    ULARGE_INTEGER u;
    u.QuadPart = (ULONGLONG)t * 10000000ULL + 116444736000000000ULL;
    FILETIME ft = {u.LowPart, u.HighPart};
    SYSTEMTIME utc, local;
    FileTimeToSystemTime(&ft, &utc);
    SystemTimeToTzSpecificLocalTime(nullptr, &utc, &local);
    wchar_t b[32];
    swprintf(b, 32, L"%04u-%02u-%02u %02u:%02u", local.wYear, local.wMonth, local.wDay, local.wHour, local.wMinute);
    return b;
}

std::wstring firstLine(const std::wstring& s) {
    size_t n = s.find_first_of(L"\r\n");
    return n == std::wstring::npos ? s : s.substr(0, n);
}

int historySelection(HWND dlg) { return ListView_GetNextItem(GetDlgItem(dlg, IDC_H_LIST), -1, LVNI_SELECTED); }

void diffAgainstRevision(const Revision& r) {
    const bool p4 = g.d->vcs.kind == VcsKind::Perforce;
    g.d->baseSpec = r.spec;
    g.d->baseLabel = p4 ? L"Depot  #" + r.rev + L"  (change " + r.change + L", " + r.user + L", " + formatTime(r.time) + L")"
                        : L"Commit  " + r.rev + L"  " + firstLine(r.desc) + L"  (" + r.user + L", " + formatTime(r.time) + L")";
    startVcsQuery(true);
    if (!g.diffOn) toggleDiff();
    invalidateTabs();
}

void diffAgainstLatest() {
    g.d->baseSpec.clear();
    g.d->baseLabel.clear();
    startVcsQuery(true);
    if (!g.diffOn) toggleDiff();
    invalidateTabs();
}

// Puts revision `r` of the file on disk and reopens it.
// Perforce: p4 sync file#N (your workspace then has #N). Git: the file's
// content at that commit is written over the working file.
bool getRevision(HWND owner, const Revision& r) {
    const bool p4 = g.d->vcs.kind == VcsKind::Perforce;
    const std::wstring name = fileNameOf(g.d->path);
    std::wstring what = p4 ? L"#" + r.rev + L" (change " + r.change + L")" : L"commit " + r.rev;
    std::wstring msg = L"Replace " + name + L" with " + what + L"?";
    const bool localChanges = g.d->doc.modified() || (p4 ? g.d->vcs.opened : !g.d->hunks.empty() && g.d->baseSpec.empty());
    if (p4 && g.d->vcs.opened)
        msg += L"\n\nThe file is checked out: it is reverted first, so your changes to it are lost.";
    else if (localChanges)
        msg += L"\n\nYour changes to it are lost.";
    if (MessageBoxW(owner, msg.c_str(), L"Get Revision", MB_OKCANCEL | (localChanges ? MB_ICONWARNING : MB_ICONQUESTION)) != IDOK)
        return false;
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    bool ok = true;
    if (p4) {
        if (g.d->vcs.opened) ok = runVcs(p4Cmd() + L" revert " + quoted(g.d->path));
        if (ok) ok = runVcs(p4Cmd() + L" sync " + quoted(g.d->path + L"#" + r.rev));
    } else {
        std::vector<u8> bytes;
        std::wstring err;
        ok = fetchRevision(g.d->path, VcsKind::Git, r.spec, 1u << 30, bytes, err);
        if (ok) {
            HANDLE h = CreateFileW((g.d->path + L".jn-tmp").c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
            DWORD put = 0;
            ok = h != INVALID_HANDLE_VALUE && WriteFile(h, bytes.data(), (DWORD)bytes.size(), &put, nullptr) && put == bytes.size();
            if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
            DWORD e2 = 0;
            ok = ok && replaceFile(g.d->path + L".jn-tmp", g.d->path, e2);
            if (!ok) DeleteFileW((g.d->path + L".jn-tmp").c_str());
        }
        if (!ok) MessageBoxW(owner, (L"Could not get " + what + L".\n\n" + err).c_str(), kAppName, MB_OK | MB_ICONERROR);
    }
    SetCursor(old);
    if (!ok) return false;
    g.d->choices.clear();
    g.d->baseSpec.clear();
    g.d->baseLabel.clear();
    openPath(g.d->path, true);
    return true;
}

INT_PTR CALLBACK historyProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp) {
    if (INT_PTR br = dialogTheme(dlg, msg, wp, lp)) return br;
    switch (msg) {
    case WM_INITDIALOG: {
        const bool p4 = g.d->vcs.kind == VcsKind::Perforce;
        SetDlgItemTextW(dlg, IDC_H_TITLE, g.d->path.c_str());
        HWND lv = GetDlgItem(dlg, IDC_H_LIST);
        ListView_SetExtendedListViewStyle(lv, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);
        RECT rc;
        GetClientRect(lv, &rc);
        const int w = rc.right - GetSystemMetrics(SM_CXVSCROLL);
        struct Col { const wchar_t* name; int pct; };
        const Col cols[] = {{p4 ? L"Rev" : L"Commit", 9}, {p4 ? L"Change" : L"Author", p4 ? 9 : 16}, {L"Date", 17},
                            {p4 ? L"User" : L"", p4 ? 12 : 0}, {L"Action", 9}, {L"Description", p4 ? 44 : 49}};
        int ci = 0;
        for (const Col& c : cols) {
            if (!c.pct) continue;
            LVCOLUMNW col = {};
            col.mask = LVCF_TEXT | LVCF_WIDTH;
            col.pszText = (LPWSTR)c.name;
            col.cx = w * c.pct / 100;
            ListView_InsertColumn(lv, ci++, &col);
        }
        int sel = 0;
        for (size_t i = 0; i < g_history.size(); i++) {
            const Revision& r = g_history[i];
            std::wstring cells[] = {p4 ? L"#" + r.rev : r.rev, p4 ? r.change : r.user, formatTime(r.time), r.user, r.action,
                                    firstLine(r.desc)};
            LVITEMW it = {};
            it.mask = LVIF_TEXT;
            it.iItem = (int)i;
            it.pszText = (LPWSTR)cells[0].c_str();
            ListView_InsertItem(lv, &it);
            int sub = 1;
            for (int k = 1; k < 6; k++) {
                if (!p4 && k == 3) continue;  // Git: the author is already in column 2
                ListView_SetItemText(lv, (int)i, sub++, (LPWSTR)cells[k].c_str());
            }
            if (r.spec == g.d->baseSpec) sel = (int)i;  // the one being diffed against
        }
        if (!g_history.empty()) {
            ListView_SetItemState(lv, sel, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_EnsureVisible(lv, sel, FALSE);
        }
        EnableWindow(GetDlgItem(dlg, IDC_H_LATEST), !g.d->baseSpec.empty());
        SetFocus(lv);
        return FALSE;
    }
    case WM_NOTIFY: {
        NMHDR* nm = (NMHDR*)lp;
        if (nm->idFrom != IDC_H_LIST) break;
        if (nm->code == LVN_ITEMCHANGED) {
            int i = historySelection(dlg);
            std::wstring d;
            if (i >= 0 && (size_t)i < g_history.size()) {
                const Revision& r = g_history[(size_t)i];
                d = r.desc;
                // Edit controls need \r\n line breaks.
                std::wstring out;
                for (wchar_t c : d) {
                    if (c == L'\n') out += L'\r';
                    out += c;
                }
                d = out;
            }
            SetDlgItemTextW(dlg, IDC_H_DESC, d.c_str());
        } else if (nm->code == NM_DBLCLK || nm->code == LVN_ITEMACTIVATE) {
            PostMessageW(dlg, WM_COMMAND, IDC_H_DIFF, 0);
        }
        break;
    }
    case WM_COMMAND: {
        int i = historySelection(dlg);
        const bool have = i >= 0 && (size_t)i < g_history.size();
        switch (LOWORD(wp)) {
        case IDC_H_DIFF:
            if (have) {
                Revision r = g_history[(size_t)i];
                EndDialog(dlg, 1);
                diffAgainstRevision(r);
            }
            return TRUE;
        case IDC_H_GET:
            if (have && getRevision(dlg, g_history[(size_t)i])) EndDialog(dlg, 1);
            return TRUE;
        case IDC_H_LATEST:
            EndDialog(dlg, 1);
            diffAgainstLatest();
            return TRUE;
        case IDOK:
            PostMessageW(dlg, WM_COMMAND, IDC_H_DIFF, 0);
            return TRUE;
        case IDCANCEL: EndDialog(dlg, 0); return TRUE;
        }
        break;
    }
    }
    return FALSE;
}

void cmdHistory() {
    if (g.d->vcs.kind == VcsKind::None || g.d->vcs.toolMissing || g.d->path.empty()) {
        MessageBeep(MB_OK);
        return;
    }
    std::wstring err;
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    const bool ok = fileHistory(g.d->path, g.d->vcs.kind, g_history, err);
    SetCursor(old);
    if (!ok) {
        MessageBoxW(g.frame, (L"No history for " + fileNameOf(g.d->path) + L".\n\n" + err).c_str(), kAppName,
                    MB_OK | MB_ICONINFORMATION);
        return;
    }
    DialogBoxParamW(g.inst, MAKEINTRESOURCEW(IDD_HISTORY), g.frame, historyProc, 0);
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
        refreshAllTabsVcs();
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

// ====================================================================== bars (find bar, data view search)
//
// A strip with one or two text boxes and square 3D buttons, part of the main
// window: the editor's find / replace bar and the data view's search box.

struct BarButton {
    int id;
    const wchar_t* text;  // label, or nullptr for a glyph
    int glyph;            // Glyph when text is null
    bool toggle;          // stays pushed in while on
    bool on;
    int row;              // 0 first row, 1 replace row
    RECT rc;
    const wchar_t* tip;
};

struct Bar {
    HWND wnd = nullptr, edit = nullptr, edit2 = nullptr;
    bool twoRows = false;  // find bar: the replace row is open
    std::vector<BarButton> buttons;
    int hot = -1, press = -1;
    std::wstring status;   // "3 of 12", "No results", ...
    bool noMatch = false;
    const wchar_t* label1 = L"";
    const wchar_t* label2 = L"";
    std::function<void(int)> action;      // a button, or Enter (id 1) / Shift+Enter (id 2) / Esc (id 3)
    std::function<void()> changed;        // the first box's text changed
};

Bar g_findBar, g_dataBar;
const wchar_t* const kBarClass = L"JustNotepadBar";
enum BarIds { BAR_ENTER = 1, BAR_SHIFTENTER = 2, BAR_ESC = 3, BAR_CASE = 10, BAR_WORD, BAR_PREV, BAR_NEXT, BAR_CLOSE,
              BAR_REPLMODE, BAR_REPLACE, BAR_REPLACEALL };

int barRowH() { return g.statusH + MulDiv(10, g.dpi, 96); }
int barHeight(const Bar& b) { return b.wnd && IsWindowVisible(b.wnd) ? barRowH() * (b.twoRows ? 2 : 1) : 0; }

std::wstring editText(HWND e) {
    int n = GetWindowTextLengthW(e);
    std::wstring s((size_t)n + 1, L'\0');
    GetWindowTextW(e, &s[0], n + 1);
    s.resize((size_t)n);
    return s;
}

void barLayout(Bar& b) {
    RECT rc;
    GetClientRect(b.wnd, &rc);
    const int px = MulDiv(1, g.dpi, 96), rh = barRowH(), bs = rh - 10 * px, pad = 8 * px, gap = 4 * px;
    const int labelW = 58 * px;
    // Buttons from the right, per row.
    int xr[2] = {(int)rc.right - pad, (int)rc.right - pad};
    const int statusW = 86 * px;
    for (int k = (int)b.buttons.size() - 1; k >= 0; k--) {
        BarButton& bt = b.buttons[(size_t)k];
        HDC dc = GetDC(b.wnd);
        HGDIOBJ of = SelectObject(dc, g.uiFont);
        SIZE sz = {bs, bs};
        if (bt.text) {
            GetTextExtentPoint32W(dc, bt.text, lstrlenW(bt.text), &sz);
            sz.cx = std::max(bs, (int)sz.cx + 14 * px);
        } else {
            sz.cx = bs;
        }
        SelectObject(dc, of);
        ReleaseDC(b.wnd, dc);
        int y = bt.row * rh + (rh - bs) / 2;
        bt.rc = {xr[bt.row] - (int)sz.cx, y, xr[bt.row], y + bs};
        xr[bt.row] -= (int)sz.cx + gap;
    }
    xr[0] -= statusW;  // room for "3 of 12" between the text box and the buttons
    const int ex = pad + labelW, eh = bs - 4 * px;
    const int ew1 = std::max(60 * px, xr[0] - ex - gap);
    MoveWindow(b.edit, ex + 2 * px, (rh - eh) / 2, ew1 - 4 * px, eh, TRUE);
    if (b.edit2) {
        const int ew2 = std::max(60 * px, xr[1] - ex - gap);
        MoveWindow(b.edit2, ex + 2 * px, rh + (rh - eh) / 2, ew2 - 4 * px, eh, TRUE);
        ShowWindow(b.edit2, b.twoRows ? SW_SHOW : SW_HIDE);
    }
}

void paintBar(HWND hwnd) {
    Bar& b = *(Bar*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(hwnd, &ps);
    RECT rc;
    GetClientRect(hwnd, &rc);
    HDC dc = CreateCompatibleDC(wdc);
    HBITMAP bmp = CreateCompatibleBitmap(wdc, std::max(1, (int)rc.right), std::max(1, (int)rc.bottom));
    HGDIOBJ ob = SelectObject(dc, bmp);
    fillRect(dc, 0, 0, rc.right, rc.bottom, g.pal.face);
    fillRect(dc, 0, rc.bottom - 1, rc.right, rc.bottom, g.pal.border);
    HGDIOBJ of = SelectObject(dc, g.uiFont);
    SetBkMode(dc, TRANSPARENT);
    const int px = MulDiv(1, g.dpi, 96), rh = barRowH(), pad = 8 * px;
    // Labels, and a sunken 90s frame around each text box.
    for (int row = 0; row < (b.twoRows ? 2 : 1); row++) {
        HWND e = row == 0 ? b.edit : b.edit2;
        if (!e) continue;
        RECT er;
        GetWindowRect(e, &er);
        MapWindowPoints(nullptr, hwnd, (POINT*)&er, 2);
        InflateRect(&er, 2 * px, 2 * px);
        fillRect(dc, er.left, er.top, er.right, er.bottom, g.pal.bg);
        drawButton3D(dc, er, BTN_PRESSED, g.pal, 0);
        fillRect(dc, er.left + 2, er.top + 2, er.right - 2, er.bottom - 2, g.pal.bg);
        RECT lr = {pad, row * rh, er.left - 4 * px, (row + 1) * rh};
        SetTextColor(dc, g.pal.faceText);
        const wchar_t* lab = row == 0 ? b.label1 : b.label2;
        DrawTextW(dc, lab, -1, &lr, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    }
    // Status ("3 of 12"), right after the first text box.
    if (!b.status.empty()) {
        RECT er;
        GetWindowRect(b.edit, &er);
        MapWindowPoints(nullptr, hwnd, (POINT*)&er, 2);
        int left = er.right + 8 * px, right = left + 82 * px;
        for (const BarButton& bt : b.buttons)
            if (bt.row == 0 && bt.rc.left >= left) right = std::min(right, (int)bt.rc.left - 4 * px);
        RECT sr = {left, 0, right, rh};
        SetTextColor(dc, b.noMatch ? (g.pal.dark ? RGB(255, 120, 110) : RGB(190, 30, 30)) : blend(g.pal.faceText, g.pal.face, 70));
        DrawTextW(dc, b.status.c_str(), -1, &sr, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
    }
    for (size_t k = 0; k < b.buttons.size(); k++) {
        const BarButton& bt = b.buttons[k];
        if (bt.row == 1 && !b.twoRows) continue;
        const bool hot = b.hot == (int)k, held = hot && b.press == (int)k;
        ButtonState st = held ? BTN_PRESSED : bt.on ? (hot ? BTN_LATCHED_HOT : BTN_LATCHED) : hot ? BTN_HOT : BTN_NORMAL;
        drawButton3D(dc, bt.rc, st, g.pal, g.pal.accent);
        if (bt.text) {
            RECT tr = bt.rc;
            if (held || bt.on) OffsetRect(&tr, 1, 1);
            SetTextColor(dc, g.pal.faceText);
            DrawTextW(dc, bt.text, -1, &tr, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        } else {
            drawGlyph(dc, bt.rc, (Glyph)bt.glyph, g.pal.faceText, g.pal.faceText, held || bt.on);
        }
    }
    SelectObject(dc, of);
    BitBlt(wdc, 0, 0, rc.right, rc.bottom, dc, 0, 0, SRCCOPY);
    SelectObject(dc, ob);
    DeleteObject(bmp);
    DeleteDC(dc);
    EndPaint(hwnd, &ps);
}

int barButtonAt(Bar& b, POINT pt) {
    for (size_t k = 0; k < b.buttons.size(); k++) {
        if (b.buttons[k].row == 1 && !b.twoRows) continue;
        if (PtInRect(&b.buttons[k].rc, pt)) return (int)k;
    }
    return -1;
}

// Keys in the text boxes: Enter / Shift+Enter / Esc go to the bar's action,
// Tab moves between the boxes and back to the editor.
LRESULT CALLBACK barEditProc(HWND e, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR, DWORD_PTR ref) {
    Bar& b = *(Bar*)ref;
    if (msg == WM_KEYDOWN) {
        const bool shift = GetKeyState(VK_SHIFT) < 0;
        if (wp == VK_RETURN) {
            if (b.action) b.action(e == b.edit2 ? (shift ? BAR_REPLACEALL : BAR_REPLACE) : shift ? BAR_SHIFTENTER : BAR_ENTER);
            return 0;
        }
        if (wp == VK_ESCAPE) {
            if (b.action) b.action(BAR_ESC);
            return 0;
        }
        if (wp == VK_TAB) {
            if (e == b.edit && b.edit2 && b.twoRows && !shift) SetFocus(b.edit2);
            else if (e == b.edit2 && shift) SetFocus(b.edit);
            else SetFocus(g.view);
            return 0;
        }
    }
    if (msg == WM_CHAR && (wp == L'\r' || wp == 27 || wp == L'\t')) return 0;  // no beep
    if (msg == WM_CHAR && wp == 1) {  // Ctrl+A in a box: select all of it
        SendMessageW(e, EM_SETSEL, 0, -1);
        return 0;
    }
    return DefSubclassProc(e, msg, wp, lp);
}

LRESULT CALLBACK barProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    Bar* bp = (Bar*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    if (msg >= WM_CTLCOLORMSGBOX && msg <= WM_CTLCOLORSTATIC && bp) {
        INT_PTR r = themeCtlColor(msg, wp, g.pal);
        if (msg == WM_CTLCOLOREDIT && (HWND)lp == bp->edit && bp->noMatch)
            SetTextColor((HDC)wp, g.pal.dark ? RGB(255, 120, 110) : RGB(190, 30, 30));
        return r;
    }
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paintBar(hwnd); return 0;
    case WM_SIZE:
        if (bp) barLayout(*bp);
        return 0;
    case WM_COMMAND:
        if (bp && HIWORD(wp) == EN_CHANGE && (HWND)lp == bp->edit && bp->changed) bp->changed();
        return 0;
    case WM_MOUSEMOVE: {
        if (!bp) break;
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        int h = barButtonAt(*bp, pt);
        if (h != bp->hot) {
            bp->hot = h;
            InvalidateRect(hwnd, nullptr, FALSE);
            TRACKMOUSEEVENT tme = {sizeof(tme), TME_LEAVE, hwnd, 0};
            TrackMouseEvent(&tme);
        }
        return 0;
    }
    case WM_MOUSELEAVE:
        if (bp && GetCapture() != hwnd && bp->hot >= 0) {
            bp->hot = -1;
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK:
        if (bp && bp->hot >= 0) {
            bp->press = bp->hot;
            SetCapture(hwnd);
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    case WM_LBUTTONUP: {
        if (!bp || GetCapture() != hwnd) return 0;
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        int k = barButtonAt(*bp, pt), p = bp->press;
        bp->press = -1;
        ReleaseCapture();
        InvalidateRect(hwnd, nullptr, FALSE);
        if (k == p && k >= 0 && bp->action) bp->action(bp->buttons[(size_t)k].id);
        return 0;
    }
    case WM_CAPTURECHANGED:
        if (bp && bp->press >= 0) {
            bp->press = -1;
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void createBar(Bar& b, HWND parent, bool withSecondBox) {
    b.wnd = CreateWindowExW(0, kBarClass, L"", WS_CHILD | WS_CLIPCHILDREN, 0, 0, 0, 0, parent, nullptr, g.inst, nullptr);
    SetWindowLongPtrW(b.wnd, GWLP_USERDATA, (LONG_PTR)&b);
    auto box = [&](int id) {
        HWND e = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL, 0, 0, 0, 0, b.wnd,
                                 (HMENU)(INT_PTR)id, g.inst, nullptr);
        SendMessageW(e, WM_SETFONT, (WPARAM)g.uiFont, FALSE);
        SetWindowTheme(e, L"", L"");
        SetWindowSubclass(e, barEditProc, 1, (DWORD_PTR)&b);
        return e;
    };
    b.edit = box(100);
    if (withSecondBox) b.edit2 = box(101);
}

void themeBars() {
    for (Bar* b : {&g_findBar, &g_dataBar}) {
        if (!b->wnd) continue;
        if (b->edit) SendMessageW(b->edit, WM_SETFONT, (WPARAM)g.uiFont, TRUE);
        if (b->edit2) SendMessageW(b->edit2, WM_SETFONT, (WPARAM)g.uiFont, TRUE);
        barLayout(*b);
        InvalidateRect(b->wnd, nullptr, TRUE);
    }
}

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
    if (findBarOpen()) updateFindStatus(hit);
    if (hit == NPOS && findBarOpen()) return false;
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
    swprintf(b, 96, L"Replaced %llu", (unsigned long long)count);
    if (findBarOpen()) {
        g_findBar.status = b;
        g_findBar.noMatch = count == 0;
        InvalidateRect(g_findBar.wnd, nullptr, FALSE);
    } else {
        MessageBoxW(g.frame, b, kAppName, MB_OK | MB_ICONINFORMATION);
    }
}

// ---- the editor's find / replace bar

bool findBarOpen() { return g_findBar.wnd && IsWindowVisible(g_findBar.wnd); }

std::string findNeedle() { return toDocBytes(g.findBuf, lstrlenW(g.findBuf)); }

// "3 of 12": counts matches (files up to 16 MB; beyond that just found / not found).
void updateFindStatus(u64 hit) {
    std::string nd = findNeedle();
    g_findBar.noMatch = false;
    if (nd.empty()) {
        g_findBar.status.clear();
    } else if (hit == NPOS) {
        g_findBar.status = L"No results";
        g_findBar.noMatch = true;
    } else if (g.d->doc.size() <= (16u << 20)) {
        u64 total = 0, index = 0, pos = 0;
        for (;;) {
            u64 h = findFrom(nd, pos, true, g.findCase, g.findWord);
            if (h == NPOS || total > 999999) break;
            total++;
            if (h <= hit) index = total;
            pos = h + std::max<u64>(1, nd.size());
        }
        wchar_t b[64];
        swprintf(b, 64, L"%llu of %llu", (unsigned long long)index, (unsigned long long)total);
        g_findBar.status = b;
    } else {
        g_findBar.status = L"Found";
    }
    if (g_findBar.wnd) InvalidateRect(g_findBar.wnd, nullptr, FALSE);
    if (g_findBar.edit) InvalidateRect(g_findBar.edit, nullptr, TRUE);
}

// Searches from `from` (wrapping), selects the hit and keeps the focus where it is.
bool findAndSelect(u64 from, bool down) {
    std::string nd = findNeedle();
    if (nd.empty()) {
        updateFindStatus(NPOS);
        return false;
    }
    HCURSOR old = SetCursor(LoadCursorW(nullptr, IDC_WAIT));
    u64 hit = findFrom(nd, from, down, g.findCase, g.findWord);
    if (hit == NPOS) hit = findFrom(nd, down ? 0 : g.d->doc.size(), down, g.findCase, g.findWord);  // wrap around
    SetCursor(old);
    if (hit != NPOS) {
        g.d->anchor = hit;
        g.d->caret = hit + nd.size();
        g.d->wantCol = -1;
        g.d->doc.breakCoalesce();
        refresh(true);
    }
    updateFindStatus(hit);
    InvalidateRect(g.view, nullptr, FALSE);  // match highlights
    return hit != NPOS;
}

void openFindBar(bool replace) {
    // Start from the selection when it is short and on one line.
    if (hasSel() && selMax() - selMin() < 200) {
        std::string s = g.d->doc.text(selMin(), selMax() - selMin());
        if (s.find('\n') == std::string::npos && s.find('\r') == std::string::npos) {
            std::wstring w = fromDocBytes(s);
            lstrcpynW(g.findBuf, w.c_str(), 256);
            SetWindowTextW(g_findBar.edit, g.findBuf);
        }
    }
    g_findBar.twoRows = replace;
    for (BarButton& bt : g_findBar.buttons) {
        if (bt.id == BAR_CASE) bt.on = g.findCase;
        if (bt.id == BAR_WORD) bt.on = g.findWord;
        if (bt.id == BAR_REPLMODE) bt.on = replace;
    }
    ShowWindow(g_findBar.wnd, SW_SHOW);
    layoutFrame();
    barLayout(g_findBar);
    SetFocus(replace && GetWindowTextLengthW(g_findBar.edit) > 0 ? g_findBar.edit2 : g_findBar.edit);
    SendMessageW(GetFocus(), EM_SETSEL, 0, -1);
    updateFindStatus(lstrlenW(g.findBuf) && hasSel() && selectionMatches(findNeedle()) ? selMin() : NPOS);
    if (!lstrlenW(g.findBuf)) g_findBar.status.clear();
    InvalidateRect(g_findBar.wnd, nullptr, FALSE);
    InvalidateRect(g.view, nullptr, FALSE);
}

void closeFindBar() {
    if (!findBarOpen()) return;
    ShowWindow(g_findBar.wnd, SW_HIDE);
    layoutFrame();
    SetFocus(g.view);
    InvalidateRect(g.view, nullptr, FALSE);
}

void showFind(bool replace) { openFindBar(replace); }

void onFindBarAction(int id) {
    switch (id) {
    case BAR_ENTER:
    case BAR_NEXT: findAndSelect(selMax(), true); break;
    case BAR_SHIFTENTER:
    case BAR_PREV: findAndSelect(selMin(), false); break;
    case BAR_ESC:
    case BAR_CLOSE: closeFindBar(); break;
    case BAR_CASE:
    case BAR_WORD:
        for (BarButton& bt : g_findBar.buttons)
            if (bt.id == id) {
                bt.on = !bt.on;
                (id == BAR_CASE ? g.findCase : g.findWord) = bt.on;
            }
        findAndSelect(selMin(), true);  // the current match may no longer count
        break;
    case BAR_REPLMODE: openFindBar(!g_findBar.twoRows); break;
    case BAR_REPLACE: {
        std::wstring r = editText(g_findBar.edit2);
        lstrcpynW(g.replBuf, r.c_str(), 256);
        replaceOne();
        break;
    }
    case BAR_REPLACEALL: {
        std::wstring r = editText(g_findBar.edit2);
        lstrcpynW(g.replBuf, r.c_str(), 256);
        replaceAll();
        break;
    }
    }
    if (findBarOpen() && GetFocus() == g.view && id != BAR_ESC && id != BAR_CLOSE) SetFocus(g_findBar.edit);
}

// Typing in the find box searches straight away, from where the current match starts.
void onFindBarChanged() {
    std::wstring t = editText(g_findBar.edit);
    lstrcpynW(g.findBuf, t.c_str(), 256);
    findAndSelect(hasSel() ? selMin() : g.d->caret, true);
}

void setupFindBar() {
    createBar(g_findBar, g.frame, true);
    g_findBar.label1 = L"Find";
    g_findBar.label2 = L"Replace";
    g_findBar.buttons = {
        {BAR_CASE, L"Aa", 0, true, g.findCase, 0, {}, L"Match case"},
        {BAR_WORD, L"ab", 0, true, g.findWord, 0, {}, L"Whole word"},
        {BAR_PREV, nullptr, GLYPH_UP, false, false, 0, {}, L"Previous (Shift+Enter)"},
        {BAR_NEXT, nullptr, GLYPH_DOWN, false, false, 0, {}, L"Next (Enter)"},
        {BAR_REPLMODE, L"Replace", 0, true, false, 0, {}, L"Show replace (Ctrl+H)"},
        {BAR_CLOSE, nullptr, GLYPH_CLOSE, false, false, 0, {}, L"Close (Esc)"},
        {BAR_REPLACE, L"Replace", 0, false, false, 1, {}, L"Replace (Enter)"},
        {BAR_REPLACEALL, L"Replace All", 0, false, false, 1, {}, L"Replace all (Shift+Enter)"},
    };
    g_findBar.action = onFindBarAction;
    g_findBar.changed = onFindBarChanged;
}

// Match highlights for one editor row (visible rows only, so this is cheap).
void paintRowMatches(HDC dc, const Row& r, int y, const LayoutCfg& c) {
    if (!findBarOpen()) return;
    const std::string nd = findNeedle();
    if (nd.empty() || nd.size() > 255 || r.end <= r.start || r.end - r.start > (1u << 16)) return;
    // Search a little past the row so matches that cross its edge still show.
    const u64 from = r.start > nd.size() ? r.start - nd.size() + 1 : 0;
    u64 to = std::min(g.d->doc.size(), r.end + nd.size() - 1);
    u64 pos = from;
    while (pos < to) {
        u64 h = g.d->doc.searchFwd(nd, pos, to, g.findCase);
        if (h == NPOS) break;
        pos = h + 1;
        if (g.findWord && !wholeWordAt(h, nd.size())) continue;
        u64 a = std::max(h, r.start), b = std::min(h + nd.size(), r.end);
        if (b <= a) continue;
        int c0 = colOf(g.d->doc, r, a, c), c1 = colOf(g.d->doc, r, b, c);
        int x0 = std::max(g.gutterPx, g.gutterPx + (c0 - g.d->hcol) * g.cw);
        int x1 = std::min(g.cliW, g.gutterPx + (c1 - g.d->hcol) * g.cw);
        if (x1 > x0) fillRect(dc, x0, y, x1, y + g.ch, blend(RGB(255, 200, 40), g.pal.bg, g.pal.dark ? 35 : 45));
    }
}

void onFindMessage(LPFINDREPLACEW) {}

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
    g.sideOn = GetPrivateProfileIntW(L"view", L"sidePane", 1, f) != 0;
    g.showWs = GetPrivateProfileIntW(L"view", L"whitespace", 0, f) != 0;
    g.recent.clear();
    for (int i = 0; i < 12; i++) {
        wchar_t key[16], val[MAX_PATH * 4];
        swprintf(key, 16, L"f%d", i);
        GetPrivateProfileStringW(L"recent", key, L"", val, MAX_PATH * 4, f);
        if (val[0]) g.recent.push_back(val);
    }
    g.diffSplit = std::clamp((int)GetPrivateProfileIntW(L"view", L"diffSplit", 500, f), 50, 950) / 1000.0;
    g.sideSplit = std::clamp((int)GetPrivateProfileIntW(L"view", L"sideSplit", 420, f), 100, 900) / 1000.0;
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
    putInt(L"view", L"sidePane", g.sideOn ? 1 : 0, f);
    putInt(L"view", L"whitespace", g.showWs ? 1 : 0, f);

    putInt(L"view", L"diffSplit", (int)(g.diffSplit * 1000 + 0.5), f);
    putInt(L"view", L"sideSplit", (int)(g.sideSplit * 1000 + 0.5), f);
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

// ====================================================================== multiple carets
//
// The main caret is g.d->caret / anchor; any others are in g.d->extra. Edits
// apply to every one of them as a single undo step; the document's edit hook
// keeps all of them on the same text as it changes.

bool multi() { return !g.d->extra.empty(); }

void clearExtras() {
    g.boxActive = false;
    g.d->caretPad = 0;
    if (g.d->extra.empty()) return;
    g.d->extra.clear();
    InvalidateRect(g.view, nullptr, FALSE);
}

// Sorts the carets and merges ones that touch; the main caret stays main.
void normalizeSels() {
    if (g.d->extra.empty()) return;
    struct S { u64 caret, anchor; bool main; int pad; };
    std::vector<S> v;
    v.push_back({g.d->caret, g.d->anchor, true, g.d->caretPad});
    for (auto& e : g.d->extra) v.push_back({e.caret, e.anchor, false, e.pad});
    auto lo = [](const S& x) { return std::min(x.caret, x.anchor); };
    auto hi = [](const S& x) { return std::max(x.caret, x.anchor); };
    std::sort(v.begin(), v.end(), [&](const S& a, const S& b) { return lo(a) < lo(b); });
    std::vector<S> out;
    for (const S& x : v) {
        if (!out.empty() && (lo(x) < hi(out.back()) || (lo(x) == hi(out.back()) && (x.caret == x.anchor || out.back().caret == out.back().anchor)))) {
            S& m = out.back();
            const u64 a = std::min(lo(m), lo(x)), b = std::max(hi(m), hi(x));
            const bool fwd = m.caret >= m.anchor;
            m.anchor = fwd ? a : b;
            m.caret = fwd ? b : a;
            m.main = m.main || x.main;
            continue;
        }
        out.push_back(x);
    }
    g.d->extra.clear();
    for (const S& x : out) {
        if (x.main) {
            g.d->caret = x.caret;
            g.d->anchor = x.anchor;
            g.d->caretPad = x.pad;
        } else {
            g.d->extra.push_back({x.caret, x.anchor, x.pad});
        }
    }
}

// Runs `f` on every caret (the main one included), from the last in the text
// to the first, as one undo step. `f` gets the caret and its index in text
// order; it edits the document around it and may set where it ends up.
void editEach(const std::function<void(DocState::Sel&, size_t)>& f) {
    auto& ex = g.d->extra;
    ex.push_back({g.d->caret, g.d->anchor, g.d->caretPad});
    const size_t mainIdx = ex.size() - 1;
    std::vector<size_t> order(ex.size());
    for (size_t i = 0; i < order.size(); i++) order[i] = i;
    std::sort(order.begin(), order.end(), [&](size_t a, size_t b) {
        return std::min(ex[a].caret, ex[a].anchor) < std::min(ex[b].caret, ex[b].anchor);
    });
    g.d->doc.breakCoalesce();
    g.d->doc.beginGroup();
    for (size_t k = order.size(); k-- > 0;) f(ex[order[k]], k);
    g.d->doc.endGroup();
    g.d->doc.breakCoalesce();
    g.d->caret = ex[mainIdx].caret;
    g.d->anchor = ex[mainIdx].anchor;
    g.d->caretPad = ex[mainIdx].pad;
    ex.erase(ex.begin() + (std::ptrdiff_t)mainIdx);
    g.boxActive = false;  // the text moved under the box
    normalizeSels();
    g.d->wantCol = -1;
    g.caretOn = true;
    refresh(true);
}

// Moves every caret; `to` maps one caret to its new place.
void moveEach(bool extend, const std::function<u64(const DocState::Sel&)>& to) {
    auto one = [&](u64& caret, u64& anchor) {
        DocState::Sel s{caret, anchor};
        caret = std::min(to(s), g.d->doc.size());
        if (!extend) anchor = caret;
    };
    one(g.d->caret, g.d->anchor);
    for (auto& e : g.d->extra) {
        one(e.caret, e.anchor);
        e.pad = 0;
    }
    g.d->caretPad = 0;
    g.boxActive = false;
    normalizeSels();
    g.d->wantCol = -1;
    g.caretOn = true;
    g.d->doc.breakCoalesce();
    refresh(true);
}

void eraseSel(DocState::Sel& s) {
    const u64 a = std::min(s.caret, s.anchor), b = std::max(s.caret, s.anchor);
    if (b > a) g.d->doc.erase(a, b - a);
}

// Text for a caret: spaces first when it sits past the end of its line.
std::string padded(DocState::Sel& s, const std::string& t) {
    const int pad = s.caret == s.anchor ? s.pad : 0;
    s.pad = 0;
    if (t.empty() || pad <= 0) return t;
    return std::string((size_t)pad, ' ') + t;
}

void multiInsert(const std::string& t) {
    editEach([&](DocState::Sel& s, size_t) {
        const std::string u = padded(s, t);
        eraseSel(s);
        if (!u.empty()) g.d->doc.insert(s.caret, (const u8*)u.data(), u.size());
    });
}

// The arrow keys, Home, End, Backspace and Delete with several carets.
bool multiKey(WPARAM vk, bool ctrl, bool shift) {
    LayoutCfg c = cfg();
    Doc& d = g.d->doc;
    auto selLo = [](const DocState::Sel& s) { return std::min(s.caret, s.anchor); };
    auto selHi = [](const DocState::Sel& s) { return std::max(s.caret, s.anchor); };
    switch (vk) {
    case VK_LEFT:
        moveEach(shift, [&](const DocState::Sel& s) {
            if (s.caret != s.anchor && !shift && !ctrl) return selLo(s);
            return ctrl ? wordLeft(d, s.caret, c) : prevChar(d, s.caret, c);
        });
        return true;
    case VK_RIGHT:
        moveEach(shift, [&](const DocState::Sel& s) {
            if (s.caret != s.anchor && !shift && !ctrl) return selHi(s);
            return ctrl ? wordRight(d, s.caret, c) : nextChar(d, s.caret, c);
        });
        return true;
    case VK_UP:
    case VK_DOWN:
        if (ctrl) return false;
        moveEach(shift, [&](const DocState::Sel& s) {
            Row r = rowContaining(d, s.caret, c), o;
            const int col = colOf(d, r, s.caret, c);
            const bool ok = vk == VK_UP ? prevRow(d, r.start, c, o) : nextRow(d, r, c, o);
            return ok ? posOfCol(d, o, col, c) : s.caret;
        });
        return true;
    case VK_HOME:
        if (ctrl) return false;
        moveEach(shift, [&](const DocState::Sel& s) {
            Row r = rowContaining(d, s.caret, c);
            u64 p = r.start;
            while (p < r.end && (d.at(p) == ' ' || d.at(p) == '\t')) p++;
            return s.caret == p ? r.start : p;
        });
        return true;
    case VK_END:
        if (ctrl) return false;
        moveEach(shift, [&](const DocState::Sel& s) { return rowContaining(d, s.caret, c).end; });
        return true;
    case VK_BACK:
        editEach([&](DocState::Sel& s, size_t) {
            if (s.caret == s.anchor && s.pad > 0) {  // in the empty space past the line: just move back
                s.pad--;
                return;
            }
            if (s.caret != s.anchor) return eraseSel(s);
            if (s.caret == 0) return;
            const u64 p = ctrl ? wordLeft(d, s.caret, c) : prevChar(d, s.caret, c);
            d.erase(p, s.caret - p);
        });
        return true;
    case VK_DELETE:
        editEach([&](DocState::Sel& s, size_t) {
            if (s.caret == s.anchor && s.pad > 0) return;  // nothing to the right
            s.pad = 0;
            if (s.caret != s.anchor) return eraseSel(s);
            if (s.caret >= d.size()) return;
            const u64 p = ctrl ? wordRight(d, s.caret, c) : nextChar(d, s.caret, c);
            d.erase(s.caret, p - s.caret);
        });
        return true;
    }
    return false;
}

// All carets in text order, as [lo, hi) ranges.
std::vector<std::pair<u64, u64>> selRanges() {
    std::vector<std::pair<u64, u64>> v;
    v.push_back({selMin(), selMax()});
    for (auto& e : g.d->extra) v.push_back({std::min(e.caret, e.anchor), std::max(e.caret, e.anchor)});
    std::sort(v.begin(), v.end());
    return v;
}

// Copy with several carets: each selection on its own line.
bool multiCopy() {
    auto v = selRanges();
    std::string t;
    bool any = false;
    for (size_t i = 0; i < v.size(); i++) {
        if (i) t += eolStr();
        t += g.d->doc.text(v[i].first, v[i].second - v[i].first);
        any = any || v[i].second > v[i].first;
    }
    if (!any) return false;
    setClipboard(fromDocBytes(t), false);
    return true;
}

// Paste with several carets: one line each when the clipboard has exactly as
// many lines as there are carets, otherwise everything at every caret.
void multiPaste(std::string s) {
    normalizeEol(s, g.d->crlf);
    const std::string eol = eolStr();
    std::vector<std::string> parts;
    size_t at = 0;
    while (true) {
        size_t k = s.find(eol, at);
        if (k == std::string::npos) {
            parts.push_back(s.substr(at));
            break;
        }
        parts.push_back(s.substr(at, k - at));
        at = k + eol.size();
    }
    if (parts.size() > 1 && parts.back().empty()) parts.pop_back();
    const size_t n = g.d->extra.size() + 1;
    const bool spread = parts.size() == n;
    editEach([&](DocState::Sel& sel, size_t i) {
        const std::string t = padded(sel, spread ? parts[i] : s);
        eraseSel(sel);
        if (!t.empty()) g.d->doc.insert(sel.caret, (const u8*)t.data(), t.size());
    });
}

// Ctrl+Alt+click: another caret.
void addCaretAt(u64 p) {
    g.d->extra.push_back({g.d->caret, g.d->anchor, g.d->caretPad});
    g.d->caret = g.d->anchor = p;
    g.d->caretPad = 0;
    g.boxActive = false;
    normalizeSels();
    g.caretOn = true;
    g.d->doc.breakCoalesce();
    refresh(false);
}

// Shift+Alt+.: the next match of the selection gets a caret too (the first
// press selects the word at the caret). Shift+Alt+;: every match.
void addNextMatch(bool all) {
    Doc& d = g.d->doc;
    if (!hasSel()) {
        u64 a, b;
        wordAt(d, g.d->caret, a, b);
        if (a == b) {
            MessageBeep(MB_OK);
            return;
        }
        g.d->anchor = a;
        g.d->caret = b;
        if (!all) {
            refresh(true);
            return;
        }
    }
    const u64 len = selMax() - selMin();
    if (len > 4096) {
        MessageBeep(MB_OK);
        return;
    }
    const std::string needle = d.text(selMin(), len);
    if (all) {
        const u64 keepA = g.d->anchor, keepC = g.d->caret;
        g.d->extra.clear();
        u64 p = 0;
        for (int n = 0; n < 20000; n++) {
            u64 hit = d.searchFwd(needle, p, d.size(), true);
            if (hit == NPOS) break;
            if (hit != std::min(keepA, keepC)) g.d->extra.push_back({hit + len, hit});
            p = hit + std::max<u64>(len, 1);
        }
    } else {
        u64 from = selMax();
        for (auto& e : g.d->extra) from = std::max(from, std::max(e.caret, e.anchor));
        u64 hit = d.searchFwd(needle, from, d.size(), true);
        if (hit == NPOS) hit = d.searchFwd(needle, 0, d.size(), true);
        auto taken = [&](u64 h) {
            if (selMin() == h) return true;
            for (auto& e : g.d->extra)
                if (std::min(e.caret, e.anchor) == h) return true;
            return false;
        };
        if (hit == NPOS || taken(hit)) {
            MessageBeep(MB_OK);
            return;
        }
        g.d->extra.push_back({g.d->caret, g.d->anchor});
        g.d->anchor = hit;
        g.d->caret = hit + len;
    }
    g.boxActive = false;
    normalizeSels();
    g.d->doc.breakCoalesce();
    refresh(true);
}

// ---- column (box) selection: Alt+drag, Shift+Alt+arrows

void boxRebuild() {
    if (!g.d->doc.linesKnown()) return;
    LayoutCfg c = cfg();
    Doc& d = g.d->doc;
    const u64 last = d.newlineCount();
    g.boxL0 = std::min(g.boxL0, last);
    g.boxL1 = std::min(g.boxL1, last);
    const u64 l0 = std::min(g.boxL0, g.boxL1), l1 = std::max(g.boxL0, g.boxL1);
    g.d->extra.clear();
    const int cLo = std::min(g.boxC0, g.boxC1);
    for (u64 l = l0; l <= l1 && l - l0 < 20000; l++) {
        Row r = rowAt(d, d.lineStart(l), c);
        const int len = colOf(d, r, r.end, c);  // the line's width in columns
        const u64 a = posOfCol(d, r, g.boxC0, c), b = posOfCol(d, r, g.boxC1, c);
        // A box entirely past the end of this line: an empty caret in the
        // space after it, at the box's left edge.
        const int pad = cLo > len ? cLo - len : 0;
        if (l == g.boxL1) {
            g.d->anchor = a;
            g.d->caret = b;
            g.d->caretPad = pad;
        } else {
            g.d->extra.push_back({b, a, pad});
        }
    }
    g.boxActive = true;
    g.caretOn = true;
    g.d->doc.breakCoalesce();
}

// The box starts from the current selection: anchor and caret are its corners.
void boxStartAtCaret() {
    if (g.boxActive || !g.d->doc.linesKnown()) return;
    LayoutCfg c = cfg();
    Doc& d = g.d->doc;
    g.d->extra.clear();
    g.boxL0 = d.lineOf(g.d->anchor);
    g.boxL1 = d.lineOf(g.d->caret);
    g.boxC0 = colOf(d, rowContaining(d, g.d->anchor, c), g.d->anchor, c);
    g.boxC1 = colOf(d, rowContaining(d, g.d->caret, c), g.d->caret, c) + g.d->caretPad;
}

void boxMove(int dl, int dc) {
    if (!g.d->doc.linesKnown()) return;
    boxStartAtCaret();
    if (dl < 0 && g.boxL1 == 0) dl = 0;
    g.boxL1 = (u64)((i64)g.boxL1 + dl);
    g.boxC1 = std::max(0, g.boxC1 + dc);
    boxRebuild();
    refresh(true);
}

// Line and visual column under a point (the column may be past the line end).
bool boxPoint(int x, int y, u64& line, int& col) {
    if (g.rows.empty() || !g.d->doc.linesKnown()) return false;
    line = g.d->doc.lineOf(posFromPoint(x, y));
    col = g.d->hcol + std::max(0, (x - g.gutterPx + g.cw / 2) / g.cw);
    return true;
}

// ====================================================================== brackets and indentation

// The bracket at (or just before) the caret and its partner.
bool bracketPair(u64& a, u64& b) {
    static const char open[] = "([{", close[] = ")]}";
    Doc& d = g.d->doc;
    const u64 n = d.size();
    u64 at = NPOS;
    auto isBr = [&](u64 p) { u8 ch = d.at(p); return ch && strchr("()[]{}", ch) != nullptr; };
    if (g.d->caret < n && isBr(g.d->caret)) at = g.d->caret;
    else if (g.d->caret > 0 && isBr(g.d->caret - 1)) at = g.d->caret - 1;
    if (at == NPOS) return false;
    const char ch = (char)d.at(at);
    const char* o = strchr(open, ch);
    const int idx = o ? (int)(o - open) : (int)(strchr(close, ch) - close);
    const u8 me = (u8)open[idx], other = (u8)close[idx];
    const bool fwd = o != nullptr;
    int depth = 0;
    u64 p = at;
    const u64 limit = 1ull << 20;  // enough for any block on screen
    for (u64 steps = 0; steps < limit; steps++) {
        const u8 x = d.at(p);
        if (x == me) depth += fwd ? 1 : -1;
        else if (x == other) depth += fwd ? -1 : 1;
        if (depth == 0) {
            a = at;
            b = p;
            return true;
        }
        if (fwd ? ++p >= n : p-- == 0) return false;
    }
    return false;
}

// The file's indent: a tab, or the number of spaces most lines use.
std::string indentUnit() {
    Doc& d = g.d->doc;
    const u64 lim = std::min<u64>(d.size(), 256u << 10);
    std::string t = d.text(0, lim);
    int tabs = 0, spaces = 0, minSp = 99;
    bool atStart = true;
    for (size_t i = 0; i < t.size(); i++) {
        if (atStart) {
            if (t[i] == '\t') tabs++;
            else if (t[i] == ' ') {
                int k = 0;
                while (i + k < t.size() && t[i + k] == ' ') k++;
                if (k >= 2 && i + k < t.size() && t[i + k] != '\r' && t[i + k] != '\n') {
                    spaces++;
                    minSp = std::min(minSp, k);
                }
            }
        }
        atStart = t[i] == '\n';
    }
    if (spaces > tabs) return std::string((size_t)std::clamp(minSp, 2, 8), ' ');
    return "\t";
}

// What Enter inserts at `at`: a line break and the indent, one level deeper
// after an opening bracket (or a colon in Python); between a pair of brackets
// the closing one moves to a line of its own. `caretOff` is where the caret
// goes in it.
std::string newlineText(u64 at, size_t& caretOff) {
    LayoutCfg c = cfg();
    Doc& d = g.d->doc;
    const u64 ls = lineStartBefore(d, at, c);
    std::string lead;
    for (u64 p = ls; p < at && lead.size() < 256; p++) {
        const u8 ch = d.at(p);
        if (ch != ' ' && ch != '\t') break;
        lead += (char)ch;
    }
    const std::string eol = eolStr();
    u64 p = at;
    while (p > ls && (d.at(p - 1) == ' ' || d.at(p - 1) == '\t')) p--;
    const u8 before = p > ls ? d.at(p - 1) : 0;
    const u8 after = at < d.size() ? d.at(at) : 0;
    const bool python = g.d->lang && std::string(langName(g.d->lang)) == "Python";
    const bool opens = before == '{' || before == '[' || before == '(' || (python && before == ':');
    if (!opens) {
        caretOff = eol.size() + lead.size();
        return eol + lead;
    }
    const std::string inner = eol + lead + indentUnit();
    caretOff = inner.size();
    const bool pair = (before == '{' && after == '}') || (before == '[' && after == ']') || (before == '(' && after == ')');
    return pair ? inner + eol + lead : inner;
}

// Typing a closing bracket on a line that is blank so far lines it up with
// the line of its opening bracket.
bool outdentCloser(u32 cp) {
    if ((cp != '}' && cp != ']' && cp != ')') || hasSel() || multi()) return false;
    LayoutCfg c = cfg();
    Doc& d = g.d->doc;
    const u64 at = g.d->caret, ls = lineStartBefore(d, at, c);
    if (at == ls) return false;
    for (u64 p = ls; p < at; p++)
        if (d.at(p) != ' ' && d.at(p) != '\t') return false;
    const u8 me = (u8)cp, open = cp == '}' ? '{' : cp == ']' ? '[' : '(';
    int depth = 1;
    u64 p = ls;
    for (u64 steps = 0; p > 0 && steps < (1u << 20); steps++) {
        const u8 x = d.at(--p);
        if (x == me) depth++;
        else if (x == open && --depth == 0) break;
    }
    if (depth != 0) return false;
    const u64 ols = lineStartBefore(d, p, c);
    std::string lead;
    for (u64 q = ols; q < p; q++) {
        const u8 ch = d.at(q);
        if (ch != ' ' && ch != '\t') break;
        lead += (char)ch;
    }
    lead += (char)cp;
    replaceRange(ls, at, lead, ls + lead.size(), ls + lead.size());
    return true;
}

// ====================================================================== whole-file edits

// Replaces the whole text as one undo step; the caret keeps its line.
void replaceWhole(const std::string& t) {
    Doc& d = g.d->doc;
    const u64 line = d.linesKnown() ? d.lineOf(g.d->caret) : 0;
    clearExtras();
    d.breakCoalesce();
    d.beginGroup();
    if (d.size()) d.erase(0, d.size());
    if (!t.empty()) d.insert(0, (const u8*)t.data(), t.size());
    d.endGroup();
    d.breakCoalesce();
    u64 p = d.linesKnown() ? d.lineStart(std::min(line, d.newlineCount())) : 0;
    afterEdit(p);
}

// Marks the file as changed without changing the text (the encoding or line
// endings changed, which only shows when it is saved).
void markChanged() {
    Doc& d = g.d->doc;
    d.breakCoalesce();
    d.beginGroup();
    const u8 x = ' ';
    d.insert(0, &x, 1);
    d.erase(0, 1);
    d.endGroup();
    d.breakCoalesce();
}

const u64 kWholeFileMax = 256ull << 20;

bool wholeFileOk() {
    if (g.d->doc.size() <= kWholeFileMax) return true;
    MessageBoxW(g.frame, L"This file is too large to convert (limit 256 MB).", kAppName, MB_OK | MB_ICONINFORMATION);
    return false;
}

// Encoding the file is saved with. Between ANSI and the Unicode ones the text
// is converted; the Unicode ones share the same text and differ only on disk.
void setEncoding(FileEnc e) {
    const FileEnc old = g.d->enc;
    if (e == old) return;
    const bool wasAnsi = old == FE_ANSI, toAnsi = e == FE_ANSI;
    if (wasAnsi != toAnsi) {
        if (!wholeFileOk()) return;
        ensureAnsiTable();
        const std::string t = g.d->doc.text(0, g.d->doc.size());
        std::wstring w;
        const UINT from = wasAnsi ? g.ansiCp : CP_UTF8, to = toAnsi ? g.ansiCp : CP_UTF8;
        int k = t.empty() ? 0 : MultiByteToWideChar(from, 0, t.data(), (int)t.size(), nullptr, 0);
        w.resize((size_t)k);
        if (k) MultiByteToWideChar(from, 0, t.data(), (int)t.size(), &w[0], k);
        BOOL lossy = FALSE;
        int m = w.empty() ? 0 : WideCharToMultiByte(to, 0, w.data(), (int)w.size(), nullptr, 0, nullptr, toAnsi ? &lossy : nullptr);
        std::string out((size_t)m, '\0');
        if (m) WideCharToMultiByte(to, 0, w.data(), (int)w.size(), &out[0], m, nullptr, toAnsi ? &lossy : nullptr);
        if (lossy && MessageBoxW(g.frame,
                                 L"Some characters in this file have no ANSI equivalent and will become \"?\".\n\nConvert anyway?",
                                 kAppName, MB_OKCANCEL | MB_ICONWARNING) != IDOK)
            return;
        g.d->enc = e;  // first, so the new text is laid out with the right decoding
        if (out != t) replaceWhole(out);
        else markChanged();
    } else {
        g.d->enc = e;
        markChanged();
    }
    detectLang();
    refresh(true);
}

// Reads the same bytes as UTF-8 or as ANSI (for a file detected wrongly).
void reinterpretAs(FileEnc e) {
    if (g.d->enc == FE_UTF16LE || g.d->enc == FE_UTF16BE || g.d->enc == e) return;
    if (e == FE_ANSI) ensureAnsiTable();
    g.d->enc = e;
    clearExtras();
    refresh(true);
    InvalidateRect(g.view, nullptr, FALSE);
}

void setLineEndings(bool crlf) {
    if (!wholeFileOk()) return;
    const std::string t = g.d->doc.text(0, g.d->doc.size());
    std::string s = t;
    normalizeEol(s, crlf);
    const bool was = g.d->crlf;
    g.d->crlf = crlf;
    if (s != t) replaceWhole(s);
    else if (was != crlf) markChanged();
    refresh(true);
}

// Lines the line commands work on: the selected ones, or every line.
bool targetLines(u64& l0, u64& l1, bool& whole) {
    if (!g.d->doc.linesKnown()) return false;
    whole = !(hasSel() && selLines(l0, l1) && l1 > l0);
    if (whole) {
        l0 = 0;
        l1 = lastLine();
        // A final line break leaves an empty last line: not one to sort.
        if (l1 > 0 && g.d->doc.lineStart(l1) == g.d->doc.size()) l1--;
    }
    return true;
}

std::vector<std::string> splitLines(const std::string& t) {
    std::vector<std::string> v;
    size_t at = 0;
    while (at < t.size()) {
        size_t k = t.find('\n', at);
        if (k == std::string::npos) k = t.size();
        size_t e = k;
        if (e > at && t[e - 1] == '\r') e--;
        v.push_back(t.substr(at, e - at));
        at = k + 1;
    }
    return v;
}

// mode 0: sort A-Z, 1: sort Z-A, 2: remove duplicates (keeps the first).
void sortLines(int mode) {
    u64 l0, l1;
    bool whole;
    if (!targetLines(l0, l1, whole)) return;
    Doc& d = g.d->doc;
    const u64 a = d.lineStart(l0), b = d.lineStart(l1 + 1);
    if (b - a > kWholeFileMax) {
        wholeFileOk();
        return;
    }
    const std::string orig = d.text(a, b - a);
    const bool hadEol = !orig.empty() && orig.back() == '\n';
    std::vector<std::string> v = splitLines(orig);
    if (mode == 2) {
        std::vector<std::string> out;
        std::vector<size_t> idx(v.size());
        for (size_t i = 0; i < idx.size(); i++) idx[i] = i;
        std::sort(idx.begin(), idx.end(), [&](size_t x, size_t y) { return v[x] != v[y] ? v[x] < v[y] : x < y; });
        std::vector<bool> keep(v.size(), true);
        for (size_t i = 1; i < idx.size(); i++)
            if (v[idx[i]] == v[idx[i - 1]]) keep[idx[i]] = false;
        for (size_t i = 0; i < v.size(); i++)
            if (keep[i]) out.push_back(std::move(v[i]));
        v.swap(out);
    } else {
        auto lower = [](const std::string& s) {
            std::string r = s;
            for (char& ch : r) ch = (char)tolower((unsigned char)ch);
            return r;
        };
        std::vector<std::pair<std::string, size_t>> keys;
        for (size_t i = 0; i < v.size(); i++) keys.push_back({lower(v[i]), i});
        std::stable_sort(keys.begin(), keys.end(), [&](const auto& x, const auto& y) {
            if (x.first != y.first) return mode == 0 ? x.first < y.first : x.first > y.first;
            return mode == 0 ? v[x.second] < v[y.second] : v[x.second] > v[y.second];
        });
        std::vector<std::string> out;
        for (auto& k : keys) out.push_back(v[k.second]);
        v.swap(out);
    }
    std::string t;
    for (size_t i = 0; i < v.size(); i++) {
        t += v[i];
        if (i + 1 < v.size() || hadEol) t += eolStr();
    }
    if (t == orig) return;
    clearExtras();
    if (whole) replaceRange(a, b, t, std::min(g.d->caret, a + t.size()), std::min(g.d->caret, a + t.size()));
    else replaceRange(a, b, t, a, a + t.size());
}

// Spaces and tabs at the ends of lines (selected lines, or all of them).
void trimTrailing() {
    u64 l0, l1;
    bool whole;
    if (!targetLines(l0, l1, whole)) return;
    Doc& d = g.d->doc;
    if (whole) l1 = lastLine();
    clearExtras();
    // The main caret rides along with the edit hook like an extra one.
    g.d->extra.push_back({g.d->caret, g.d->anchor});
    d.breakCoalesce();
    d.beginGroup();
    bool any = false;
    for (u64 l = l1 + 1; l-- > l0;) {
        const u64 ls = d.lineStart(l);
        u64 e = l < lastLine() ? d.lineStart(l + 1) : d.size();
        if (e > ls && d.at(e - 1) == '\n') e--;
        if (e > ls && d.at(e - 1) == '\r') e--;
        u64 p = e;
        while (p > ls && (d.at(p - 1) == ' ' || d.at(p - 1) == '\t')) p--;
        if (p < e) {
            d.erase(p, e - p);
            any = true;
        }
    }
    d.endGroup();
    d.breakCoalesce();
    g.d->caret = g.d->extra.back().caret;
    g.d->anchor = g.d->extra.back().anchor;
    g.d->extra.clear();
    if (!any) MessageBeep(MB_OK);
    refresh(true);
}

// ====================================================================== outside changes to open files

bool statPath(const std::wstring& path, FILETIME& t, u64& size) {
    WIN32_FILE_ATTRIBUTE_DATA a;
    if (path.empty() || !GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &a)) return false;
    t = a.ftLastWriteTime;
    size = ((u64)a.nFileSizeHigh << 32) | a.nFileSizeLow;
    return true;
}

void noteDiskState() {
    g.d->diskKnown = statPath(g.d->path, g.d->diskTime, g.d->diskSize);
}

// The open file changed on disk (sync, checkout, another editor): reload it,
// asking first when it has unsaved edits here. Only the tab in front is
// checked; the others are when they come to the front.
void checkDisk(bool mayAsk) {
    static bool busy = false;
    if (busy || !g.d || g.d->path.empty() || !g.d->diskKnown) return;
    FILETIME t;
    u64 size;
    if (!statPath(g.d->path, t, size)) return;  // gone or locked for now
    if (CompareFileTime(&t, &g.d->diskTime) == 0 && size == g.d->diskSize) return;
    if (g.d->doc.modified()) {
        if (!mayAsk) return;
        busy = true;
        std::wstring msg = fileNameOf(g.d->path) +
                           L"\n\nThis file was changed outside JustNotepad. Reload it and lose the changes made here?";
        DocState* asked = g.d;
        int r = MessageBoxW(g.frame, msg.c_str(), kAppName, MB_YESNO | MB_ICONWARNING);
        busy = false;
        if (g.d != asked) return;
        if (r != IDYES) {
            g.d->diskTime = t;  // keep ours; ask again only if it changes again
            g.d->diskSize = size;
            return;
        }
    }
    const u64 caret = g.d->caret, top = g.d->top;
    const int hcol = g.d->hcol;
    if (openPath(g.d->path, true)) {
        g.d->caret = g.d->anchor = std::min(caret, g.d->doc.size());
        g.d->top = std::min(top, g.d->doc.size());
        g.d->hcol = hcol;
        refresh(false);
        updateSidePane();
    } else {
        noteDiskState();  // could not read it: do not keep trying
    }
}

// ====================================================================== recent files

bool g_restoring = false;  // reopening the last session: not "recent" opens

void saveRecent() {
    std::wstring p = settingsPath(true);
    if (p.empty()) return;
    const wchar_t* f = p.c_str();
    WritePrivateProfileStringW(L"recent", nullptr, nullptr, f);
    for (size_t i = 0; i < g.recent.size(); i++) {
        wchar_t key[16];
        swprintf(key, 16, L"f%d", (int)i);
        WritePrivateProfileStringW(L"recent", key, g.recent[i].c_str(), f);
    }
}

void addRecent(const std::wstring& path) {
    if (path.empty() || g_restoring) return;
    auto& r = g.recent;
    r.erase(std::remove_if(r.begin(), r.end(), [&](const std::wstring& x) { return _wcsicmp(x.c_str(), path.c_str()) == 0; }), r.end());
    r.insert(r.begin(), path);
    if (r.size() > 12) r.resize(12);
    saveRecent();
}

void rebuildRecentMenu() {
    HMENU m = g.recentMenu;
    if (!m) return;
    themeReleasePopup(m);
    while (GetMenuItemCount(m) > 0) DeleteMenu(m, 0, MF_BYPOSITION);
    if (g.recent.empty()) {
        AppendMenuW(m, MF_STRING | MF_GRAYED, IDM_RECENT0, L"No recent files");
    } else {
        for (size_t i = 0; i < g.recent.size(); i++) {
            std::wstring label = (i < 9 ? L"&" + std::to_wstring(i + 1) : std::wstring(L"1&") + std::to_wstring(i - 9)) + L"  ";
            std::wstring p = g.recent[i];
            if (p.size() > 90) p = p.substr(0, 30) + L"\x2026" + p.substr(p.size() - 58);
            for (wchar_t ch : p) {
                if (ch == L'&') label += L'&';  // not a mnemonic
                label += ch;
            }
            AppendMenuW(m, MF_STRING, IDM_RECENT0 + (UINT)i, label.c_str());
        }
        AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(m, MF_STRING, IDM_RECENTCLEAR, L"&Clear Recent Files");
    }
    themeOwnerDrawPopup(m);
    themePopupColors(m, g.pal);
}

void openRecent(size_t i) {
    if (i >= g.recent.size()) return;
    const std::wstring p = g.recent[i];
    if (GetFileAttributesW(p.c_str()) == INVALID_FILE_ATTRIBUTES) {
        std::wstring msg = p + L"\n\nThis file no longer exists. Remove it from the list?";
        if (MessageBoxW(g.frame, msg.c_str(), kAppName, MB_YESNO | MB_ICONQUESTION) == IDYES) {
            g.recent.erase(g.recent.begin() + (std::ptrdiff_t)i);
            saveRecent();
        }
        return;
    }
    openInTab(p);
}

// ====================================================================== printing

void cmdPrint() {
    PRINTDLGW pd = {};
    pd.lStructSize = sizeof(pd);
    pd.hwndOwner = g.frame;
    pd.Flags = PD_RETURNDC | PD_NOPAGENUMS | PD_USEDEVMODECOPIESANDCOLLATE | (hasSel() ? 0 : PD_NOSELECTION);
    if (!PrintDlgW(&pd)) return;
    HDC dc = pd.hDC;
    const bool selOnly = (pd.Flags & PD_SELECTION) != 0;
    if (pd.hDevMode) GlobalFree(pd.hDevMode);
    if (pd.hDevNames) GlobalFree(pd.hDevNames);
    if (!dc) return;
    HCURSOR oldCur = SetCursor(LoadCursorW(nullptr, IDC_WAIT));

    Doc& d = g.d->doc;
    const u64 a = selOnly ? selMin() : 0, b = selOnly ? selMax() : d.size();
    const std::wstring text = fromDocBytes(d.text(a, std::min<u64>(b - a, kWholeFileMax)));

    const int dpiY = GetDeviceCaps(dc, LOGPIXELSY), dpiX = GetDeviceCaps(dc, LOGPIXELSX);
    const int pageW = GetDeviceCaps(dc, HORZRES), pageH = GetDeviceCaps(dc, VERTRES);
    const int mx = dpiX * 6 / 10, my = dpiY * 6 / 10;  // 15 mm margins
    LOGFONTW lf = {};
    lf.lfHeight = -MulDiv(10, dpiY, 72);
    lf.lfCharSet = DEFAULT_CHARSET;
    lf.lfPitchAndFamily = FIXED_PITCH | FF_MODERN;
    lstrcpynW(lf.lfFaceName, g.fontFace.c_str(), LF_FACESIZE);
    HFONT font = CreateFontIndirectW(&lf);
    HGDIOBJ oldFont = SelectObject(dc, font);
    TEXTMETRICW tm = {};
    GetTextMetricsW(dc, &tm);
    const int lh = tm.tmHeight + tm.tmExternalLeading, cw = std::max(1, (int)tm.tmAveCharWidth);
    const int cols = std::max(10, (pageW - 2 * mx) / cw);
    const int rows = std::max(5, (pageH - 2 * my) / lh - 3);  // header and footer lines

    // Lines with tabs expanded, cut to the page width.
    std::vector<std::wstring> out;
    std::wstring cur;
    auto flush = [&]() {
        out.push_back(cur);
        cur.clear();
    };
    for (size_t i = 0; i <= text.size(); i++) {
        if (i == text.size() || text[i] == L'\n') {
            if (!cur.empty() && cur.back() == L'\r') cur.pop_back();
            if (i < text.size() || !cur.empty()) flush();
            continue;
        }
        if (text[i] == L'\t') {
            do cur += L' ';
            while (cur.size() % 4);
        } else {
            cur += text[i];
        }
        if ((int)cur.size() >= cols) flush();
    }
    if (out.empty()) out.push_back(L"");

    const std::wstring name = g.d->path.empty() ? L"Untitled" : g.d->path;
    DOCINFOW di = {};
    di.cbSize = sizeof(di);
    const std::wstring title = fileNameOf(g.d->path);
    di.lpszDocName = title.c_str();
    bool ok = StartDocW(dc, &di) > 0;
    const int pages = (int)((out.size() + rows - 1) / rows);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, RGB(0, 0, 0));
    for (int pg = 0; ok && pg < pages; pg++) {
        if (StartPage(dc) <= 0) {
            ok = false;
            break;
        }
        SelectObject(dc, font);
        TextOutW(dc, mx, my, name.c_str(), (int)std::min<size_t>(name.size(), (size_t)cols));
        RECT rule = {mx, my + lh + lh / 3, pageW - mx, my + lh + lh / 3 + std::max(1, dpiY / 150)};
        FillRect(dc, &rule, (HBRUSH)GetStockObject(BLACK_BRUSH));
        for (int r = 0; r < rows; r++) {
            size_t i = (size_t)pg * rows + r;
            if (i >= out.size()) break;
            TextOutW(dc, mx, my + (r + 2) * lh, out[i].c_str(), (int)out[i].size());
        }
        wchar_t foot[64];
        int fn = swprintf(foot, 64, L"%d / %d", pg + 1, pages);
        SIZE fs;
        GetTextExtentPoint32W(dc, foot, fn, &fs);
        TextOutW(dc, (pageW - fs.cx) / 2, pageH - my - lh, foot, fn);
        if (EndPage(dc) <= 0) ok = false;
    }
    if (ok) EndDoc(dc);
    else AbortDoc(dc);
    SelectObject(dc, oldFont);
    DeleteObject(font);
    DeleteDC(dc);
    SetCursor(oldCur);
    if (!ok) MessageBoxW(g.frame, L"Printing failed.", kAppName, MB_OK | MB_ICONWARNING);
}

// ====================================================================== encoding menu (File > Encoding, status bar)

void addEncodingItems(HMENU enc, HMENU eol) {
    AppendMenuW(enc, MF_STRING, IDM_ENC_UTF8, L"Save as &UTF-8");
    AppendMenuW(enc, MF_STRING, IDM_ENC_UTF8BOM, L"Save as UTF-8 with &BOM");
    AppendMenuW(enc, MF_STRING, IDM_ENC_UTF16LE, L"Save as UTF-16 &LE");
    AppendMenuW(enc, MF_STRING, IDM_ENC_UTF16BE, L"Save as UTF-16 B&E");
    AppendMenuW(enc, MF_STRING, IDM_ENC_ANSI, L"Save as &ANSI");
    AppendMenuW(enc, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(enc, MF_STRING, IDM_REINTERP_UTF8, L"Reinterpret as U&TF-8");
    AppendMenuW(enc, MF_STRING, IDM_REINTERP_ANSI, L"Reinterpret as A&NSI");
    AppendMenuW(eol, MF_STRING, IDM_EOL_CRLF, L"&Windows (CRLF)");
    AppendMenuW(eol, MF_STRING, IDM_EOL_LF, L"&Unix (LF)");
}

void checkEncodingItems(HMENU m) {
    for (int k = 0; k < 5; k++)
        CheckMenuItem(m, IDM_ENC_UTF8 + k, MF_BYCOMMAND | (g.d->enc == (FileEnc)k ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(m, IDM_EOL_CRLF, MF_BYCOMMAND | (g.d->crlf ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(m, IDM_EOL_LF, MF_BYCOMMAND | (!g.d->crlf ? MF_CHECKED : MF_UNCHECKED));
    const bool raw = g.d->enc != FE_UTF16LE && g.d->enc != FE_UTF16BE;
    EnableMenuItem(m, IDM_REINTERP_UTF8, MF_BYCOMMAND | (raw && g.d->enc == FE_ANSI ? MF_ENABLED : MF_GRAYED));
    EnableMenuItem(m, IDM_REINTERP_ANSI, MF_BYCOMMAND | (raw && g.d->enc != FE_ANSI ? MF_ENABLED : MF_GRAYED));
}

void markEncodingRadios(HMENU m) {
    for (int k = 0; k < 5; k++) themeSetMenuRadio(m, IDM_ENC_UTF8 + k);
    themeSetMenuRadio(m, IDM_EOL_CRLF);
    themeSetMenuRadio(m, IDM_EOL_LF);
}

// A click on the status bar: encoding and line endings.
void showEncodingMenu(POINT screen) {
    HMENU m = CreatePopupMenu(), enc = CreatePopupMenu(), eol = CreatePopupMenu();
    addEncodingItems(enc, eol);
    AppendMenuW(m, MF_POPUP, (UINT_PTR)enc, L"&Encoding");
    AppendMenuW(m, MF_POPUP, (UINT_PTR)eol, L"&Line Endings");
    checkEncodingItems(m);
    themeOwnerDrawPopup(m);
    markEncodingRadios(m);
    themePopupColors(m, g.pal);
    UINT cmd = (UINT)TrackPopupMenuEx(m, TPM_RETURNCMD | TPM_NONOTIFY | TPM_BOTTOMALIGN, screen.x, screen.y, g.frame, nullptr);
    themeReleasePopup(m);
    DestroyMenu(m);
    if (cmd) onCommand((int)cmd);
}

// ====================================================================== side pane: data view and Markdown preview
//
// One pane to the right of the editor (View > Data View / Preview,
// Ctrl+Shift+V). JSON and XML show as a tree of their structure: clicking a
// node selects it in the editor, F2 or a double click edits a value in place.
// Markdown shows rendered, following the editor as you scroll and type.

enum SideKind { SIDE_NONE, SIDE_JSON, SIDE_XML, SIDE_MD, SIDE_HEX, SIDE_CSV };

const u64 kSideMax = 64ull << 20;  // larger files: no tree / preview

int sideKind() {
    if (!g.d) return SIDE_NONE;
    if (g.d->forceHex || g.d->binary) return SIDE_HEX;
    const std::wstring& p = g.d->path;
    auto ext = [&](const wchar_t* e) {
        size_t n = wcslen(e);
        return p.size() > n && _wcsicmp(p.c_str() + p.size() - n, e) == 0;
    };
    if (ext(L".csv") || ext(L".tsv")) return SIDE_CSV;
    if (!g.d->lang) return SIDE_NONE;
    const char* n = langName(g.d->lang);
    if (strcmp(n, "JSON") == 0) return SIDE_JSON;
    if (strcmp(n, "XML") == 0) return SIDE_XML;
    if (strcmp(n, "Markdown") == 0) return SIDE_MD;
    return SIDE_NONE;
}

bool sideVisible() { return g.sideOn && !g.diffOn && sideKind() != SIDE_NONE; }

std::wstring docWide(const std::string& bytes) { return fromDocBytes(bytes); }

// ---- data view (tree)

std::wstring nodePath(int id) {
    std::wstring p;
    while (id >= 0 && id < (int)g.data.nodes.size()) {
        p = fromDocBytes(g.data.nodes[(size_t)id].label) + L"/" + p;
        id = g.data.nodes[(size_t)id].parent;
    }
    return p;
}

std::wstring nodeLabel(const DataNode& n) {
    auto cut = [](std::wstring s) {
        for (wchar_t& c : s)
            if (c == L'\n' || c == L'\r' || c == L'\t') c = L' ';
        if (s.size() > 200) s = s.substr(0, 200) + L"\x2026";
        return s;
    };
    const std::wstring label = fromDocBytes(n.label), value = cut(fromDocBytes(n.value));
    switch (n.kind) {
    case DataNode::Object:
    case DataNode::Array: return label + L"  " + value;
    case DataNode::String: return label + L": \"" + value + L"\"";
    case DataNode::Element: return L"<" + label + L">" + (value.empty() ? L"" : L"  " + value);
    case DataNode::Attribute: return label + L" = \"" + value + L"\"";
    case DataNode::Text: return value;
    case DataNode::Comment: return L"<!--" + value + L"-->";
    default: return label + L": " + value;
    }
}

bool nodeEditable(const DataNode& n) {
    switch (n.kind) {
    case DataNode::String:
    case DataNode::Number:
    case DataNode::Bool:
    case DataNode::Null:
    case DataNode::Attribute:
    case DataNode::Text:
    case DataNode::Comment: return n.vend > n.vstart || n.kind == DataNode::Attribute || n.kind == DataNode::String;
    case DataNode::Element: return n.vend > n.vstart;  // an element holding only text
    default: return false;
    }
}

HTREEITEM insertNode(HTREEITEM parent, int id) {
    const DataNode& n = g.data.nodes[(size_t)id];
    std::wstring text = nodeLabel(n);
    TVINSERTSTRUCTW is = {};
    is.hParent = parent;
    is.hInsertAfter = TVI_LAST;
    is.item.mask = TVIF_TEXT | TVIF_PARAM | TVIF_CHILDREN;
    is.item.pszText = (LPWSTR)text.c_str();
    is.item.lParam = id;
    // Children are added only when the node is first expanded (huge files stay quick).
    is.item.cChildren = n.kids.empty() ? 0 : 1;
    return (HTREEITEM)SendMessageW(g.sideTree, TVM_INSERTITEMW, 0, (LPARAM)&is);
}

void fillKids(HTREEITEM item, int id) {
    if (id < 0 || (size_t)id >= g.dataFilled.size() || g.dataFilled[(size_t)id]) return;
    g.dataFilled[(size_t)id] = true;
    for (int k : g.data.nodes[(size_t)id].kids) insertNode(item, k);
}

// Re-reads the document into the tree, keeping what was expanded and selected.
void runDataSearch(bool keepPlace);

void rebuildDataView() {
    if (!g.sideTree) return;
    // Remember the open branches and the selection by path.
    std::vector<std::wstring> open;
    std::wstring selPath;
    std::function<void(HTREEITEM)> walk = [&](HTREEITEM it) {
        for (; it; it = TreeView_GetNextSibling(g.sideTree, it)) {
            TVITEMW tv = {};
            tv.mask = TVIF_PARAM | TVIF_STATE;
            tv.hItem = it;
            tv.stateMask = TVIS_EXPANDED;
            TreeView_GetItem(g.sideTree, &tv);
            if ((int)tv.lParam >= 0 && (tv.state & TVIS_EXPANDED)) {
                open.push_back(nodePath((int)tv.lParam));
                walk(TreeView_GetChild(g.sideTree, it));
            }
        }
    };
    if (g.dataDoc == g.d) {
        walk(TreeView_GetRoot(g.sideTree));
        if (HTREEITEM s = TreeView_GetSelection(g.sideTree)) {
            TVITEMW tv = {};
            tv.mask = TVIF_PARAM;
            tv.hItem = s;
            TreeView_GetItem(g.sideTree, &tv);
            if ((int)tv.lParam >= 0) selPath = nodePath((int)tv.lParam);
        }
    }
    g.treeUpdating = true;
    SendMessageW(g.sideTree, WM_SETREDRAW, FALSE, 0);
    TreeView_DeleteAllItems(g.sideTree);
    g.data = DataTree();
    g.dataFilled.clear();
    g.dataDoc = g.d;
    g.dataVersion = g.d->doc.version();
    const int kind = sideKind();
    if (g.d->doc.size() > kSideMax) {
        g.data.ok = false;
        g.data.error = "The file is too large for the data view (limit 64 MB)";
    } else {
        std::string text = g.d->doc.text(0, g.d->doc.size());
        g.data = kind == SIDE_XML ? parseXml(text.data(), text.size(), 20000000) : parseJson(text.data(), text.size(), 20000000);
    }
    g.dataFilled.assign(g.data.nodes.size(), false);
    if (!g.data.ok) {
        // The first error, on top; clicking it goes there.
        std::wstring m = L"\x26A0 " + widenUtf8(g.data.error);
        if (g.d->doc.linesKnown() && g.d->doc.size() <= kSideMax) {
            wchar_t b[64];
            swprintf(b, 64, L" (line %llu)", (unsigned long long)g.d->doc.lineOf(std::min(g.data.errorPos, g.d->doc.size())) + 1);
            m += b;
        }
        TVINSERTSTRUCTW is = {};
        is.hParent = TVI_ROOT;
        is.hInsertAfter = TVI_LAST;
        is.item.mask = TVIF_TEXT | TVIF_PARAM;
        is.item.pszText = (LPWSTR)m.c_str();
        is.item.lParam = -1;
        SendMessageW(g.sideTree, TVM_INSERTITEMW, 0, (LPARAM)&is);
    }
    if (!g.data.nodes.empty()) {
        // XML: the document node itself is not shown, its children are the top level.
        std::vector<int> top;
        if (kind == SIDE_XML) top = g.data.nodes[0].kids, g.dataFilled[0] = true;
        else top.push_back(0);
        std::function<void(HTREEITEM, int)> restore = [&](HTREEITEM it, int id) {
            const std::wstring path = nodePath(id);
            const bool wasOpen = std::find(open.begin(), open.end(), path) != open.end();
            const bool firstTime = open.empty() && g.data.nodes[(size_t)id].parent <= 0;  // a new file: open the top
            if ((wasOpen || firstTime) && !g.data.nodes[(size_t)id].kids.empty()) {
                fillKids(it, id);
                TreeView_Expand(g.sideTree, it, TVE_EXPAND);
                for (HTREEITEM c = TreeView_GetChild(g.sideTree, it); c; c = TreeView_GetNextSibling(g.sideTree, c)) {
                    TVITEMW tv = {};
                    tv.mask = TVIF_PARAM;
                    tv.hItem = c;
                    TreeView_GetItem(g.sideTree, &tv);
                    restore(c, (int)tv.lParam);
                }
            }
            if (!selPath.empty() && path == selPath) TreeView_SelectItem(g.sideTree, it);
        };
        for (int id : top) {
            HTREEITEM it = insertNode(TVI_ROOT, id);
            restore(it, id);
        }
    }
    if (g.data.truncated) {
        TVINSERTSTRUCTW is = {};
        is.hParent = TVI_ROOT;
        is.hInsertAfter = TVI_LAST;
        is.item.mask = TVIF_TEXT | TVIF_PARAM;
        is.item.pszText = (LPWSTR)L"\x2026 (too many nodes to show them all)";
        is.item.lParam = -1;
        SendMessageW(g.sideTree, TVM_INSERTITEMW, 0, (LPARAM)&is);
    }
    SendMessageW(g.sideTree, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(g.sideTree, nullptr, TRUE);
    g.treeUpdating = false;
    if (g_dataBar.edit && GetWindowTextLengthW(g_dataBar.edit)) runDataSearch(true);
}

int treeNodeOf(HTREEITEM it) {
    if (!it) return -2;
    TVITEMW tv = {};
    tv.mask = TVIF_PARAM;
    tv.hItem = it;
    TreeView_GetItem(g.sideTree, &tv);
    return (int)tv.lParam;
}

// Selects a node's text in the editor (without taking the focus).
void showNodeInEditor(int id) {
    u64 a, b;
    if (id == -1) {  // the error line
        a = b = std::min(g.data.errorPos, g.d->doc.size());
    } else if (id >= 0 && (size_t)id < g.data.nodes.size()) {
        a = g.data.nodes[(size_t)id].start;
        b = g.data.nodes[(size_t)id].end;
    } else {
        return;
    }
    if (g.dataVersion != g.d->doc.version()) return;  // the tree is behind; it catches up shortly
    a = std::min(a, g.d->doc.size());
    b = std::min(b, g.d->doc.size());
    g.d->caret = a;  // caret on the start so it scrolls there; the selection covers the node
    g.d->anchor = b;
    g.d->wantCol = -1;
    refresh(true);
}

// Writes an edited value back into the document (one undo step).
bool applyNodeEdit(int id, const std::wstring& text) {
    if (id < 0 || (size_t)id >= g.data.nodes.size() || g.dataVersion != g.d->doc.version()) return false;
    const DataNode n = g.data.nodes[(size_t)id];
    std::string v = toDocBytes(text.c_str(), (int)text.size()), raw;
    switch (n.kind) {
    case DataNode::String: raw = jsonQuote(v); break;
    case DataNode::Number:
    case DataNode::Bool:
    case DataNode::Null: raw = jsonIsLiteral(v) ? v : jsonQuote(v); break;  // anything else becomes a string
    case DataNode::Attribute: raw = xmlEscape(v, true); break;
    case DataNode::Comment:
        if (v.find("--") != std::string::npos) {
            MessageBoxW(g.frame, L"An XML comment cannot contain \"--\".", kAppName, MB_OK | MB_ICONINFORMATION);
            return false;
        }
        raw = v;
        break;
    default: raw = xmlEscape(v, false); break;
    }
    replaceRange(n.vstart, n.vend, raw, n.vstart, n.vstart + raw.size());
    rebuildDataView();
    return true;
}

// ---- data view search

std::vector<int> g_dataHits;  // nodes matching the search, in document order
int g_dataHit = -1;

std::string lowerAscii(std::string s) {
    for (char& c : s) c = (char)((c >= 'A' && c <= 'Z') ? c + 32 : c);
    return s;
}

// Finds the tree item for node `id`, opening its parents on the way (they
// are filled in lazily).
HTREEITEM revealNode(int id) {
    std::vector<int> chain;
    for (int k = id; k >= 0; k = g.data.nodes[(size_t)k].parent) chain.push_back(k);
    std::reverse(chain.begin(), chain.end());
    if (sideKind() == SIDE_XML && !chain.empty() && chain[0] == 0) chain.erase(chain.begin());  // the hidden document node
    HTREEITEM level = TreeView_GetRoot(g.sideTree), found = nullptr;
    for (size_t i = 0; i < chain.size(); i++) {
        found = nullptr;
        for (HTREEITEM it = level; it; it = TreeView_GetNextSibling(g.sideTree, it))
            if (treeNodeOf(it) == chain[i]) {
                found = it;
                break;
            }
        if (!found) return nullptr;
        if (i + 1 < chain.size()) {
            fillKids(found, chain[i]);
            TreeView_Expand(g.sideTree, found, TVE_EXPAND);
            level = TreeView_GetChild(g.sideTree, found);
        }
    }
    return found;
}

void showDataHit() {
    wchar_t b[64];
    if (g_dataHits.empty()) {
        g_dataBar.status = editText(g_dataBar.edit).empty() ? L"" : L"No results";
        g_dataBar.noMatch = !g_dataBar.status.empty();
    } else {
        swprintf(b, 64, L"%d of %d", g_dataHit + 1, (int)g_dataHits.size());
        g_dataBar.status = b;
        g_dataBar.noMatch = false;
        if (HTREEITEM it = revealNode(g_dataHits[(size_t)g_dataHit])) {
            g.treeUpdating = true;
            TreeView_SelectItem(g.sideTree, it);
            TreeView_EnsureVisible(g.sideTree, it);
            g.treeUpdating = false;
            showNodeInEditor(g_dataHits[(size_t)g_dataHit]);
        }
    }
    InvalidateRect(g_dataBar.wnd, nullptr, FALSE);
    InvalidateRect(g_dataBar.edit, nullptr, TRUE);
}

// Every key, name and value of the whole document, opened or not (ignores case).
void runDataSearch(bool keepPlace) {
    std::wstring w = editText(g_dataBar.edit);
    const std::string q = lowerAscii(toDocBytes(w.c_str(), (int)w.size()));
    const int before = g_dataHit;
    g_dataHits.clear();
    g_dataHit = -1;
    if (!q.empty()) {
        for (size_t i = 0; i < g.data.nodes.size(); i++) {
            const DataNode& n = g.data.nodes[i];
            if (sideKind() == SIDE_XML && i == 0) continue;
            const bool container = n.kind == DataNode::Object || n.kind == DataNode::Array;
            if (lowerAscii(n.label).find(q) != std::string::npos || (!container && lowerAscii(n.value).find(q) != std::string::npos))
                g_dataHits.push_back((int)i);
        }
        if (!g_dataHits.empty()) g_dataHit = keepPlace ? std::clamp(before, 0, (int)g_dataHits.size() - 1) : 0;
    }
    showDataHit();
}

void onDataBarAction(int id) {
    if (id == BAR_ESC) {
        SetWindowTextW(g_dataBar.edit, L"");
        SetFocus(g.sideTree);
        return;
    }
    if (g_dataHits.empty()) return;
    const int n = (int)g_dataHits.size();
    if (id == BAR_ENTER || id == BAR_NEXT) g_dataHit = (g_dataHit + 1) % n;
    if (id == BAR_SHIFTENTER || id == BAR_PREV) g_dataHit = (g_dataHit + n - 1) % n;
    showDataHit();
}

void setupDataBar() {
    createBar(g_dataBar, g.frame, false);
    g_dataBar.label1 = L"Search";
    g_dataBar.buttons = {
        {BAR_PREV, nullptr, GLYPH_UP, false, false, 0, {}, L"Previous (Shift+Enter)"},
        {BAR_NEXT, nullptr, GLYPH_DOWN, false, false, 0, {}, L"Next (Enter)"},
    };
    g_dataBar.action = onDataBarAction;
    g_dataBar.changed = []() { runDataSearch(false); };
}

// ---- Markdown preview

enum MdFont { MF_BODY = 0, MF_BOLD = 1, MF_ITALIC = 2, MF_BOLDITALIC = 3, MF_CODE = 4, MF_H1 = 5 };  // H1..H6 = 5..10

struct MdOp {
    enum Kind : u8 { Text, Fill, Box };
    Kind kind = Text;
    RECT rc = {};
    std::wstring text;
    int font = MF_BODY;
    COLORREF color = 0;
    bool strike = false, underline = false;
    std::string url;
};

struct MdLayout {
    std::vector<MdOp> ops;
    std::vector<int> blockTop;  // y of each block
    int height = 0, width = -1;
    u64 version = NPOS;
    const DocState* doc = nullptr;
};
MdLayout g_md;
std::vector<MdBlock> g_mdBlocks;
HFONT g_mdFonts[11] = {};
int g_mdFontDpi = 0, g_mdFontPt = 0;

void makeMdFonts() {
    if (g_mdFontDpi == g.dpi && g_mdFontPt == g.fontPt && g_mdFonts[0]) return;
    for (HFONT& f : g_mdFonts)
        if (f) DeleteObject(f), f = nullptr;
    g_mdFontDpi = g.dpi;
    g_mdFontPt = g.fontPt;
    const int body = -MulDiv(g.fontPt, g.dpi, 72);
    auto mk = [&](int h, int weight, bool italic, const wchar_t* face) {
        return CreateFontW(h, 0, 0, 0, weight, italic, 0, 0, DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, 0, face);
    };
    g_mdFonts[MF_BODY] = mk(body, FW_NORMAL, false, L"Segoe UI");
    g_mdFonts[MF_BOLD] = mk(body, FW_SEMIBOLD, false, L"Segoe UI");
    g_mdFonts[MF_ITALIC] = mk(body, FW_NORMAL, true, L"Segoe UI");
    g_mdFonts[MF_BOLDITALIC] = mk(body, FW_SEMIBOLD, true, L"Segoe UI");
    g_mdFonts[MF_CODE] = mk(body * 92 / 100, FW_NORMAL, false, g.fontFace.c_str());
    static const int scale[6] = {200, 160, 135, 115, 100, 90};
    for (int k = 0; k < 6; k++) g_mdFonts[MF_H1 + k] = mk(body * scale[k] / 100, FW_BOLD, false, L"Segoe UI");
}

int fontHeight(HDC dc, int f) {
    TEXTMETRICW tm;
    SelectObject(dc, g_mdFonts[f]);
    GetTextMetricsW(dc, &tm);
    return tm.tmHeight;
}

void layoutMarkdown(int width) {
    makeMdFonts();
    HDC dc = GetDC(g.sidePreview);
    HGDIOBJ oldFont = SelectObject(dc, g_mdFonts[MF_BODY]);
    MdLayout L;
    L.width = width;
    L.version = g.d->doc.version();
    L.doc = g.d;
    const int px = MulDiv(1, g.dpi, 96);
    const int pad = 18 * px, maxX = std::max(pad + 40 * px, width - pad);
    const COLORREF fg = g.pal.fg, bg = g.pal.bg, dim = blend(fg, bg, 55), codeBg = blend(fg, bg, g.pal.dark ? 9 : 6);
    const COLORREF heading = g.pal.dark ? RGB(240, 240, 240) : RGB(10, 10, 10);
    const int bodyH = fontHeight(dc, MF_BODY);
    int y = pad;

    auto measure = [&](const std::wstring& s, int f) {
        SIZE sz = {};
        SelectObject(dc, g_mdFonts[f]);
        GetTextExtentPoint32W(dc, s.c_str(), (int)s.size(), &sz);
        return (int)sz.cx;
    };
    auto fontFor = [&](const MdRun& r, int base) {
        if (r.style & MdRun::Code) return (int)MF_CODE;
        if (base >= MF_H1) return base;
        int f = MF_BODY;
        if (r.style & MdRun::Bold) f |= MF_BOLD;
        if (r.style & MdRun::Italic) f |= MF_ITALIC;
        return f;
    };
    // Lays out runs from x0 to maxX starting at y; returns the y below the last line.
    auto flow = [&](const std::vector<MdRun>& runs, int x0, int top, int baseFont, COLORREF color) {
        const int lh = fontHeight(dc, baseFont) * 135 / 100;
        int x = x0, yy = top;
        for (const MdRun& r : runs) {
            const int f = fontFor(r, baseFont);
            const std::wstring w = fromDocBytes(r.text);
            size_t i = 0;
            while (i < w.size()) {
                if (w[i] == L'\n') {  // hard line break
                    x = x0;
                    yy += lh;
                    i++;
                    continue;
                }
                // A token: a word with the spaces after it.
                size_t e = i;
                while (e < w.size() && w[e] != L' ' && w[e] != L'\n') e++;
                while (e < w.size() && w[e] == L' ') e++;
                std::wstring tok = w.substr(i, e - i);
                int tw = measure(tok, f);
                if (x + tw > maxX && x > x0) {
                    x = x0;
                    yy += lh;
                    while (!tok.empty() && tok[0] == L' ') tok.erase(0, 1);
                    tw = measure(tok, f);
                }
                if (r.style & MdRun::Code) {
                    std::wstring core = tok;
                    while (!core.empty() && core.back() == L' ') core.pop_back();
                    MdOp b;
                    b.kind = MdOp::Fill;
                    b.rc = {x - 2 * px, yy, x + measure(core, f) + 2 * px, yy + lh - 2 * px};
                    b.color = codeBg;
                    L.ops.push_back(b);
                }
                MdOp t;
                t.kind = MdOp::Text;
                t.rc = {x, yy + (lh - fontHeight(dc, f)) / 2, x + tw, yy + lh};
                t.text = tok;
                t.font = f;
                t.color = (r.style & MdRun::Link) ? g.pal.accent : color;
                t.strike = (r.style & MdRun::Strike) != 0;
                t.underline = (r.style & MdRun::Link) != 0;
                t.url = r.url;
                L.ops.push_back(std::move(t));
                x += tw;
                i = e;
            }
        }
        return yy + lh;
    };

    L.blockTop.resize(g_mdBlocks.size());
    for (size_t bi = 0; bi < g_mdBlocks.size(); bi++) {
        const MdBlock& b = g_mdBlocks[bi];
        const int qx = pad + b.quote * 16 * px;
        const int blockY = y;
        L.blockTop[bi] = y;
        switch (b.kind) {
        case MdBlock::Heading: {
            const int f = MF_H1 + std::clamp(b.level, 1, 6) - 1;
            y += (b.level <= 2 ? 14 : 8) * px;
            L.blockTop[bi] = y;
            y = flow(b.runs, qx, y, f, heading);
            if (b.level <= 2) {  // a rule under the big headings
                MdOp r;
                r.kind = MdOp::Fill;
                r.rc = {qx, y + 2 * px, maxX, y + 3 * px};
                r.color = g.pal.border;
                L.ops.push_back(r);
                y += 6 * px;
            }
            y += 6 * px;
            break;
        }
        case MdBlock::Paragraph: y = flow(b.runs, qx, y, MF_BODY, fg) + bodyH * 6 / 10; break;
        case MdBlock::ListItem: {
            const int ind = qx + b.level * 22 * px, textX = ind + 20 * px;
            const int lh = bodyH * 135 / 100;
            if (b.task >= 0) {  // a square check box
                const int s = bodyH * 7 / 10;
                MdOp box;
                box.kind = MdOp::Box;
                box.rc = {ind + 2 * px, y + (lh - s) / 2, ind + 2 * px + s, y + (lh - s) / 2 + s};
                box.color = b.task ? g.pal.accent : dim;
                box.strike = b.task == 1;  // "checked"
                L.ops.push_back(box);
            } else {
                MdOp m;
                m.kind = MdOp::Text;
                m.text = b.ordered ? std::to_wstring(b.number) + L"." : (b.level % 2 ? L"\x25E6" : L"\x2022");
                m.font = MF_BODY;
                m.color = dim;
                m.rc = {ind, y + (lh - bodyH) / 2, textX, y + lh};
                L.ops.push_back(m);
            }
            y = flow(b.runs, textX, y, MF_BODY, b.task == 1 ? dim : fg);
            // A small gap after the last item of a list.
            bool lastInList = bi + 1 == g_mdBlocks.size() || g_mdBlocks[bi + 1].kind != MdBlock::ListItem;
            y += lastInList ? bodyH * 6 / 10 : 2 * px;
            break;
        }
        case MdBlock::CodeBlock: {
            const int lh = fontHeight(dc, MF_CODE) * 125 / 100, ip = 10 * px;
            MdOp box;
            box.kind = MdOp::Fill;
            box.rc = {qx, y, maxX, y + (int)b.code.size() * lh + 2 * ip};
            box.color = codeBg;
            L.ops.push_back(box);
            int yy = y + ip;
            for (const std::string& line : b.code) {
                std::wstring w = fromDocBytes(line);
                for (size_t k = 0; k < w.size(); k++)  // tabs as 4 spaces
                    if (w[k] == L'\t') w.replace(k, 1, L"    ");
                MdOp t;
                t.kind = MdOp::Text;
                t.text = w;
                t.font = MF_CODE;
                t.color = fg;
                t.rc = {qx + ip, yy, maxX - ip, yy + lh};
                L.ops.push_back(std::move(t));
                yy += lh;
            }
            y = yy + ip + bodyH * 6 / 10;
            break;
        }
        case MdBlock::Rule: {
            MdOp r;
            r.kind = MdOp::Fill;
            r.rc = {qx, y + bodyH / 2, maxX, y + bodyH / 2 + 2 * px};
            r.color = g.pal.border;
            L.ops.push_back(r);
            y += bodyH + bodyH * 6 / 10;
            break;
        }
        case MdBlock::Table: {
            const size_t cols = b.align.size();
            const int cp = 8 * px, lh = bodyH * 150 / 100;
            std::vector<int> w(cols, 0);
            auto cellText = [&](const std::vector<MdRun>& runs) {
                std::wstring s;
                for (const MdRun& r : runs) s += fromDocBytes(r.text);
                return s;
            };
            for (size_t r = 0; r < b.rows.size(); r++)
                for (size_t c = 0; c < cols && c < b.rows[r].size(); c++)
                    w[c] = std::max(w[c], measure(cellText(b.rows[r][c]), r == 0 ? MF_BOLD : MF_BODY) + 2 * cp);
            int total = 0;
            for (int v : w) total += v;
            const int avail = maxX - qx;
            if (total > avail && total > 0)  // too wide: shrink the columns in proportion
                for (int& v : w) v = std::max(24 * px, v * avail / total);
            for (size_t r = 0; r < b.rows.size(); r++) {
                int x = qx;
                if (r == 0) {
                    MdOp h;
                    h.kind = MdOp::Fill;
                    h.rc = {qx, y, qx + std::min(total, avail), y + lh};
                    h.color = codeBg;
                    L.ops.push_back(h);
                }
                for (size_t c = 0; c < cols; c++) {
                    std::wstring s = c < b.rows[r].size() ? cellText(b.rows[r][c]) : L"";
                    const int f = r == 0 ? MF_BOLD : MF_BODY, tw = measure(s, f);
                    int tx = x + cp;
                    if (b.align[c] == 0) tx = x + (w[c] - tw) / 2;
                    else if (b.align[c] == 1) tx = x + w[c] - cp - tw;
                    MdOp t;
                    t.kind = MdOp::Text;
                    t.text = s;
                    t.font = f;
                    t.color = fg;
                    t.rc = {std::max(tx, x + cp), y + (lh - bodyH) / 2, x + w[c] - cp, y + lh};
                    L.ops.push_back(std::move(t));
                    // Cell border: square, one pixel.
                    MdOp e;
                    e.kind = MdOp::Box;
                    e.rc = {x, y, x + w[c] + 1, y + lh + 1};
                    e.color = g.pal.border;
                    L.ops.push_back(e);
                    x += w[c];
                }
                y += lh;
            }
            y += bodyH * 6 / 10;
            break;
        }
        }
        if (b.quote > 0) {  // the quote bar, the height of the block
            for (int q = 0; q < b.quote; q++) {
                MdOp bar;
                bar.kind = MdOp::Fill;
                bar.rc = {pad + q * 16 * px, blockY, pad + q * 16 * px + 3 * px, y - bodyH * 4 / 10};
                bar.color = blend(fg, bg, 30);
                L.ops.push_back(bar);
            }
        }
    }
    L.height = y + pad;
    SelectObject(dc, oldFont);
    ReleaseDC(g.sidePreview, dc);
    g_md = std::move(L);
}

void reparseMarkdown() {
    if (g.d->doc.size() > kSideMax) g_mdBlocks.clear();
    else g_mdBlocks = parseMarkdown(g.d->doc.text(0, g.d->doc.size()));
    g_md.width = -1;  // lay out again on the next paint
}

void updatePreviewScrollbar() {
    RECT rc;
    GetClientRect(g.sidePreview, &rc);
    SCROLLINFO si = {sizeof(si), SIF_RANGE | SIF_PAGE | SIF_POS};
    si.nMax = std::max(0, g_md.height);
    si.nPage = (UINT)rc.bottom;
    g.mdScroll = std::clamp(g.mdScroll, 0, std::max(0, g_md.height - (int)rc.bottom));
    si.nPos = g.mdScroll;
    SetScrollInfo(g.sidePreview, SB_VERT, &si, TRUE);
}

// Scrolls the preview to where the editor is: the block holding the editor's
// top line, plus how far into that block it is.
// The preview and the editor are tied by blocks: block i covers source lines
// [line_i, line_(i+1)) and preview rows [top_i, top_(i+1)); inside a block the
// two are interpolated. Used both ways, so they scroll together like the diff.
double previewYForLine(double line) {
    const size_t n = g_mdBlocks.size();
    if (!n) return 0;
    for (size_t i = 0; i < n; i++) {
        const double l0 = g_mdBlocks[i].line, l1 = i + 1 < n ? g_mdBlocks[i + 1].line : std::max<double>(l0 + 1, g_mdBlocks[i].lineEnd);
        const double y0 = g_md.blockTop[i], y1 = i + 1 < n ? g_md.blockTop[i + 1] : g_md.height;
        if (line < l0) return i == 0 ? y0 * (line / std::max(1.0, l0)) : y0;  // before the first block
        if (line < l1) return y0 + (y1 - y0) * (line - l0) / std::max(1.0, l1 - l0);
    }
    return g_md.height;
}

double previewLineForY(double y) {
    const size_t n = g_mdBlocks.size();
    if (!n) return 0;
    for (size_t i = 0; i < n; i++) {
        const double y0 = g_md.blockTop[i], y1 = i + 1 < n ? g_md.blockTop[i + 1] : g_md.height;
        const double l0 = g_mdBlocks[i].line, l1 = i + 1 < n ? g_mdBlocks[i + 1].line : std::max<double>(l0 + 1, g_mdBlocks[i].lineEnd);
        if (y < y0) return i == 0 ? l0 * (y / std::max(1.0, y0)) : l0;
        if (y < y1) return l0 + (l1 - l0) * (y - y0) / std::max(1.0, y1 - y0);
    }
    return g_mdBlocks.back().lineEnd;
}

// Editor moved: the preview follows.
void syncPreviewToEditor() {
    if (!g.sidePreview || !IsWindowVisible(g.sidePreview) || !g.d->doc.linesKnown() || g_md.blockTop.empty()) return;
    const u64 top = g.d->doc.lineOf(g.d->top);
    if (top == g.mdSyncedTop) return;  // already level (or the preview just moved the editor here)
    g.mdSyncedTop = top;
    g.mdScroll = std::max(0, (int)std::lround(previewYForLine((double)top)));
    updatePreviewScrollbar();
    InvalidateRect(g.sidePreview, nullptr, FALSE);
}

// Preview moved (wheel, scroll bar): the editor follows.
void syncEditorToPreview() {
    if (!g.d->doc.linesKnown() || g_md.blockTop.empty()) return;
    const u64 last = g.d->doc.newlineCount();
    const u64 line = std::min<u64>((u64)std::max(0.0, previewLineForY(g.mdScroll)), last);
    g.mdSyncedTop = line;  // so the editor does not push the preview back
    g.d->top = g.d->doc.lineStart(line);
    g.d->topFill = 0;
    refresh(false);
}

void paintPreview(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(hwnd, &ps);
    RECT rc;
    GetClientRect(hwnd, &rc);
    const int W = rc.right, H = rc.bottom;
    if (W <= 0 || H <= 0) {
        EndPaint(hwnd, &ps);
        return;
    }
    if (g_md.width != W || g_md.version != g.d->doc.version() || g_md.doc != g.d) {
        if (g_md.version != g.d->doc.version() || g_md.doc != g.d) reparseMarkdown();
        layoutMarkdown(W);
        updatePreviewScrollbar();
    }
    HDC dc = CreateCompatibleDC(wdc);
    HBITMAP bmp = CreateCompatibleBitmap(wdc, W, H);
    HGDIOBJ oldBmp = SelectObject(dc, bmp);
    fillRect(dc, 0, 0, W, H, g.pal.bg);
    fillRect(dc, 0, 0, edgeW(), H, g.pal.border);  // a line between editor and preview
    SetBkMode(dc, TRANSPARENT);
    HGDIOBJ oldFont = SelectObject(dc, g_mdFonts[MF_BODY]);
    const int s = g.mdScroll;
    for (const MdOp& op : g_md.ops) {
        if (op.rc.bottom - s < 0 || op.rc.top - s > H) continue;
        RECT r = op.rc;
        OffsetRect(&r, 0, -s);
        if (op.kind == MdOp::Fill) {
            fillRect(dc, r.left, r.top, r.right, r.bottom, op.color);
        } else if (op.kind == MdOp::Box) {
            HBRUSH b = CreateSolidBrush(op.color);
            FrameRect(dc, &r, b);
            if (op.strike) {  // a ticked task box: filled inside
                RECT in = r;
                InflateRect(&in, -3, -3);
                FillRect(dc, &in, b);
            }
            DeleteObject(b);
        } else {
            SelectObject(dc, g_mdFonts[op.font]);
            SetTextColor(dc, op.color);
            ExtTextOutW(dc, r.left, r.top, ETO_CLIPPED, &r, op.text.c_str(), (UINT)op.text.size(), nullptr);
            if (op.strike || op.underline) {
                SIZE sz = {};
                std::wstring t = op.text;
                while (!t.empty() && t.back() == L' ') t.pop_back();
                GetTextExtentPoint32W(dc, t.c_str(), (int)t.size(), &sz);
                TEXTMETRICW tm;
                GetTextMetricsW(dc, &tm);
                int ly = op.strike ? r.top + tm.tmAscent * 6 / 10 : r.top + tm.tmAscent + 1;
                fillRect(dc, r.left, ly, r.left + sz.cx, ly + std::max(1, MulDiv(1, g.dpi, 96)), op.color);
            }
        }
    }
    if (g_mdBlocks.empty()) {
        SelectObject(dc, g.uiFont);
        SetTextColor(dc, blend(g.pal.fg, g.pal.bg, 45));
        RECT t = rc;
        DrawTextW(dc, g.d->doc.size() > kSideMax ? L"The file is too large to preview (limit 64 MB)." : L"Nothing to preview yet.",
                  -1, &t, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }
    SelectObject(dc, oldFont);
    BitBlt(wdc, 0, 0, W, H, dc, 0, 0, SRCCOPY);
    SelectObject(dc, oldBmp);
    DeleteObject(bmp);
    DeleteDC(dc);
    EndPaint(hwnd, &ps);
}

const MdOp* linkAt(int x, int y) {
    for (const MdOp& op : g_md.ops) {
        if (op.kind != MdOp::Text || op.url.empty()) continue;
        RECT r = op.rc;
        OffsetRect(&r, 0, -g.mdScroll);
        POINT pt = {x, y};
        if (PtInRect(&r, pt)) return &op;
    }
    return nullptr;
}

void openLink(const std::string& url) {
    std::wstring u = widenUtf8(url);
    if (u.rfind(L"http://", 0) == 0 || u.rfind(L"https://", 0) == 0 || u.rfind(L"mailto:", 0) == 0) {
        ShellExecuteW(nullptr, L"open", u.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
        return;
    }
    if (u.empty() || u[0] == L'#') return;
    // A relative link to another file: open it in a tab.
    for (wchar_t& c : u)
        if (c == L'/') c = L'\\';
    std::wstring path = dirOf(g.d->path) + L"\\" + u;
    if (GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES) openInTab(path);
}

// The divider between the editor and the side pane: drag it, double-click to reset.
LRESULT CALLBACK splitProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC dc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);
        fillRect(dc, 0, 0, rc.right, rc.bottom, g.pal.face);
        fillRect(dc, 0, 0, 1, rc.bottom, g.pal.dark ? blend(RGB(255, 255, 255), g.pal.face, 18) : RGB(255, 255, 255));
        fillRect(dc, rc.right - 1, 0, rc.right, rc.bottom, g.pal.border);
        // Grip: three small squares in the middle.
        const int cx = rc.right / 2, cy = rc.bottom / 2, k = std::max(2, MulDiv(2, g.dpi, 96));
        for (int n = -1; n <= 1; n++) fillRect(dc, cx - k / 2, cy + n * 3 * k, cx - k / 2 + k, cy + n * 3 * k + k, blend(g.pal.faceText, g.pal.face, 45));
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_SETCURSOR: SetCursor(LoadCursorW(nullptr, IDC_SIZEWE)); return TRUE;
    case WM_LBUTTONDBLCLK:
        g.sideSplit = 0.42;
        layoutFrame();
        saveSettings();
        return 0;
    case WM_LBUTTONDOWN:
        g.dragSplit = 2;
        g.dragOffset = GET_X_LPARAM(lp);
        SetCapture(hwnd);
        return 0;
    case WM_MOUSEMOVE:
        if (g.dragSplit == 2 && GetCapture() == hwnd) {
            POINT f = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
            MapWindowPoints(hwnd, g.frame, &f, 1);
            RECT rc;
            GetClientRect(g.frame, &rc);
            // Where the divider's left edge goes; the pane gets everything right of it.
            const int x = f.x - g.dragOffset;
            if (rc.right > 0) g.sideSplit = std::clamp((double)(rc.right - x - splitW()) / rc.right, 0.1, 0.9);
            layoutFrame();
            UpdateWindow(g.frame);
        }
        return 0;
    case WM_LBUTTONUP:
        if (GetCapture() == hwnd) {
            ReleaseCapture();
            saveSettings();
            SetFocus(g.view);
        }
        g.dragSplit = 0;
        return 0;
    case WM_CAPTURECHANGED: g.dragSplit = 0; return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT CALLBACK previewProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paintPreview(hwnd); return 0;
    case WM_SIZE:
        g_md.width = -1;
        InvalidateRect(hwnd, nullptr, FALSE);
        return 0;
    case WM_MOUSEWHEEL: {
        int lines = 3;
        SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
        const int step = MulDiv(g.ch, lines, 1);
        g.mdScroll -= GET_WHEEL_DELTA_WPARAM(wp) * step / WHEEL_DELTA;
        updatePreviewScrollbar();
        InvalidateRect(hwnd, nullptr, FALSE);
        syncEditorToPreview();
        return 0;
    }
    case WM_VSCROLL: {
        RECT rc;
        GetClientRect(hwnd, &rc);
        SCROLLINFO si = {sizeof(si), SIF_ALL};
        GetScrollInfo(hwnd, SB_VERT, &si);
        switch (LOWORD(wp)) {
        case SB_LINEUP: g.mdScroll -= g.ch * 2; break;
        case SB_LINEDOWN: g.mdScroll += g.ch * 2; break;
        case SB_PAGEUP: g.mdScroll -= rc.bottom * 9 / 10; break;
        case SB_PAGEDOWN: g.mdScroll += rc.bottom * 9 / 10; break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: g.mdScroll = si.nTrackPos; break;
        case SB_TOP: g.mdScroll = 0; break;
        case SB_BOTTOM: g.mdScroll = g_md.height; break;
        default: return 0;
        }
        updatePreviewScrollbar();
        InvalidateRect(hwnd, nullptr, FALSE);
        syncEditorToPreview();
        UpdateWindow(g.view);  // both move in the same frame while the thumb is dragged
        return 0;
    }
    case WM_SETCURSOR: {
        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(hwnd, &pt);
        if (LOWORD(lp) == HTCLIENT && linkAt(pt.x, pt.y)) {
            SetCursor(LoadCursorW(nullptr, IDC_HAND));
            return TRUE;
        }
        break;
    }
    case WM_LBUTTONUP:
        if (const MdOp* op = linkAt(GET_X_LPARAM(lp), GET_Y_LPARAM(lp))) openLink(op->url);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ---- hex view (binary files, or View > Show as Hex)
//
// Offset | 16 bytes in hex | the same bytes as text. Reads only the rows on
// screen, so it works on files of any size. Typing hex digits (or text in the
// right column) overwrites bytes; Delete removes one; Save writes the bytes.

struct HexState {
    u64 top = 0;          // first row shown
    u64 caret = 0;        // byte under the caret
    int nibble = 0;       // 0: high digit next, 1: low digit
    bool ascii = false;   // caret in the text column
    u64 syncedTop = NPOS; // editor top line the view last followed
};
HexState g_hex;

u64 hexRows() { return g.d->doc.size() / 16 + 1; }  // +1: room to type after the end
int hexOffDigits() { return g.d->doc.size() > 0xFFFFFFFFull ? 16 : 8; }
int hexX(int i, bool ascii) {
    const int pad = g.cw, hexStart = pad + (hexOffDigits() + 2) * g.cw;
    if (!ascii) return hexStart + (i * 3 + (i >= 8 ? 1 : 0)) * g.cw;
    return hexStart + (16 * 3 + 3) * g.cw + i * g.cw;
}
int hexVisibleRows() {
    if (g.ch <= 0) return 1;
    RECT rc;
    GetClientRect(g.sideHex, &rc);
    return std::max(1, (int)rc.bottom / g.ch);
}

void hexScrollbar() {
    if (!g.d || !g.sideHex || g.ch <= 0) return;  // not ready during window creation
    SCROLLINFO si = {sizeof(si), SIF_RANGE | SIF_PAGE | SIF_POS};
    const u64 rows = hexRows();
    int sh = 0;
    while ((rows >> sh) > 0x3FFFFFFFull) sh++;
    g.hexShift = sh;
    si.nMax = (int)(rows >> sh);
    si.nPage = (UINT)std::max(1, hexVisibleRows() >> sh);
    si.nPos = (int)(g_hex.top >> sh);
    SetScrollInfo(g.sideHex, SB_VERT, &si, TRUE);
}

void hexEnsureCaretVisible() {
    const u64 row = g_hex.caret / 16, vis = (u64)hexVisibleRows();
    if (row < g_hex.top) g_hex.top = row;
    else if (row >= g_hex.top + vis) g_hex.top = row - vis + 1;
}

// The hex view moved: the editor goes to the same bytes.
void hexSyncEditor(bool caretToo) {
    if (caretToo) {
        g.d->caret = g.d->anchor = std::min(g_hex.caret, g.d->doc.size());
        g.d->wantCol = -1;
    }
    g.d->top = std::min(g_hex.top * 16, g.d->doc.size());
    g.d->topFill = 0;
    g_hex.syncedTop = g.d->doc.linesKnown() ? g.d->doc.lineOf(g.d->top) : NPOS;
    refresh(false);
}

// The editor moved: the hex view follows.
void syncHexToEditor() {
    if (!g.sideHex || !IsWindowVisible(g.sideHex) || !g.d->doc.linesKnown()) return;
    const u64 line = g.d->doc.lineOf(g.d->top);
    if (line == g_hex.syncedTop) return;
    g_hex.syncedTop = line;
    g_hex.top = std::min(g.d->top / 16, hexRows() - 1);
    hexScrollbar();
    InvalidateRect(g.sideHex, nullptr, FALSE);
}

void paintHex(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(hwnd, &ps);
    RECT rc;
    GetClientRect(hwnd, &rc);
    const int W = rc.right, H = rc.bottom;
    HDC dc = CreateCompatibleDC(wdc);
    HBITMAP bmp = CreateCompatibleBitmap(wdc, std::max(1, W), std::max(1, H));
    HGDIOBJ ob = SelectObject(dc, bmp);
    const COLORREF bg = g.pal.bg, fg = g.pal.fg, dim = blend(fg, bg, 40), sel = g.pal.selection;
    fillRect(dc, 0, 0, W, H, bg);
    HGDIOBJ of = SelectObject(dc, g.font);
    SetBkMode(dc, TRANSPARENT);
    const u64 size = g.d->doc.size(), s0 = selMin(), s1 = selMax();
    const bool focused = GetFocus() == hwnd;
    u8 buf[16];
    for (int r = 0; r * g.ch < H; r++) {
        const u64 row = g_hex.top + (u64)r;
        const u64 off = row * 16;
        if (off > size) break;
        const int y = r * g.ch;
        const u64 n = std::min<u64>(16, size - off);
        g.d->doc.copy(off, buf, n);
        wchar_t t[24];
        swprintf(t, 24, hexOffDigits() == 16 ? L"%016llX" : L"%08llX", (unsigned long long)off);
        SetTextColor(dc, dim);
        TextOutW(dc, g.cw, y, t, hexOffDigits());
        for (int i = 0; i < 16; i++) {
            const u64 at = off + (u64)i;
            const bool isCaret = at == g_hex.caret;
            const bool inSel = at >= s0 && at < s1;
            if ((u64)i >= n && !isCaret) continue;
            const int xh = hexX(i, false), xa = hexX(i, true);
            if (inSel) {
                fillRect(dc, xh, y, xh + 2 * g.cw, y + g.ch, sel);
                fillRect(dc, xa, y, xa + g.cw, y + g.ch, sel);
            }
            if (isCaret) {
                // The active column gets a filled caret cell; the other one an outline.
                const COLORREF cc = focused ? g.pal.accent : blend(g.pal.accent, bg, 50);
                const int ax = g_hex.ascii ? xa : xh + g_hex.nibble * g.cw, aw = g.cw;
                fillRect(dc, ax, y, ax + aw, y + g.ch, blend(cc, bg, 45));
                fillRect(dc, ax, y + g.ch - 2, ax + aw, y + g.ch, cc);
                const int ox = g_hex.ascii ? xh : xa, ow = g_hex.ascii ? 2 * g.cw : g.cw;
                fillRect(dc, ox, y, ox + ow, y + 1, cc);
                fillRect(dc, ox, y + g.ch - 1, ox + ow, y + g.ch, cc);
            }
            if ((u64)i >= n) continue;
            const u8 b = buf[i];
            swprintf(t, 4, L"%02X", b);
            SetTextColor(dc, b == 0 ? dim : fg);
            TextOutW(dc, xh, y, t, 2);
            wchar_t ch = (b >= 0x20 && b < 0x7F) ? (wchar_t)b : L'\x00B7';
            SetTextColor(dc, (b >= 0x20 && b < 0x7F) ? fg : dim);
            TextOutW(dc, xa, y, &ch, 1);
        }
    }
    // Column separators.
    const int sx = hexX(0, false) - g.cw, ax = hexX(0, true) - g.cw;
    fillRect(dc, sx, 0, sx + 1, H, g.pal.border);
    fillRect(dc, ax, 0, ax + 1, H, g.pal.border);
    fillRect(dc, 0, 0, edgeW(), H, g.pal.border);
    SelectObject(dc, of);
    BitBlt(wdc, 0, 0, W, H, dc, 0, 0, SRCCOPY);
    SelectObject(dc, ob);
    DeleteObject(bmp);
    DeleteDC(dc);
    EndPaint(hwnd, &ps);
}

// Writes one byte at `at` (appends at the end), as its own undo step.
void hexPutByte(u64 at, u8 b) {
    const u64 size = g.d->doc.size();
    g.d->doc.breakCoalesce();
    g.d->doc.beginGroup();
    if (at < size) g.d->doc.erase(at, 1);
    g.d->doc.insert(std::min(at, size), &b, 1);
    g.d->doc.endGroup();
}

void hexMove(i64 delta, bool toAscii = false) {
    (void)toAscii;
    const i64 c = (i64)g_hex.caret + delta;
    g_hex.caret = (u64)std::clamp<i64>(c, 0, (i64)g.d->doc.size());
    g_hex.nibble = 0;
    hexEnsureCaretVisible();
    hexScrollbar();
    hexSyncEditor(true);
    InvalidateRect(g.sideHex, nullptr, FALSE);
}

int hexDigit(wchar_t c) {
    if (c >= L'0' && c <= L'9') return c - L'0';
    if (c >= L'a' && c <= L'f') return c - L'a' + 10;
    if (c >= L'A' && c <= L'F') return c - L'A' + 10;
    return -1;
}

LRESULT CALLBACK hexProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paintHex(hwnd); return 0;
    case WM_SIZE: hexScrollbar(); return 0;
    case WM_SETFOCUS:
    case WM_KILLFOCUS: InvalidateRect(hwnd, nullptr, FALSE); return 0;
    case WM_GETDLGCODE: return DLGC_WANTALLKEYS | DLGC_WANTCHARS;
    case WM_MOUSEWHEEL: {
        int lines = 3;
        SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
        i64 t = (i64)g_hex.top - (i64)GET_WHEEL_DELTA_WPARAM(wp) * lines / WHEEL_DELTA;
        g_hex.top = (u64)std::clamp<i64>(t, 0, (i64)hexRows() - 1);
        hexScrollbar();
        InvalidateRect(hwnd, nullptr, FALSE);
        hexSyncEditor(false);
        return 0;
    }
    case WM_VSCROLL: {
        SCROLLINFO si = {sizeof(si), SIF_ALL};
        GetScrollInfo(hwnd, SB_VERT, &si);
        i64 t = (i64)g_hex.top, page = std::max(1, hexVisibleRows() - 1);
        switch (LOWORD(wp)) {
        case SB_LINEUP: t--; break;
        case SB_LINEDOWN: t++; break;
        case SB_PAGEUP: t -= page; break;
        case SB_PAGEDOWN: t += page; break;
        case SB_TOP: t = 0; break;
        case SB_BOTTOM: t = (i64)hexRows(); break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: t = (i64)((u64)(unsigned)si.nTrackPos << g.hexShift); break;
        default: return 0;
        }
        g_hex.top = (u64)std::clamp<i64>(t, 0, (i64)hexRows() - 1);
        hexScrollbar();
        InvalidateRect(hwnd, nullptr, FALSE);
        hexSyncEditor(false);
        UpdateWindow(g.view);
        return 0;
    }
    case WM_LBUTTONDOWN: {
        SetFocus(hwnd);
        const int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp);
        const u64 row = g_hex.top + (u64)std::max(0, y / g.ch);
        int best = -1;
        bool ascii = x >= hexX(0, true) - g.cw / 2;
        for (int i = 0; i < 16; i++) {
            const int x0 = hexX(i, ascii), w = ascii ? g.cw : 3 * g.cw;
            if (x >= x0 - (ascii ? 0 : g.cw / 2) && x < x0 + w) best = i;
        }
        if (best < 0) best = x < hexX(0, ascii) ? 0 : 15;
        g_hex.ascii = ascii;
        g_hex.caret = std::min(row * 16 + (u64)best, g.d->doc.size());
        g_hex.nibble = (!ascii && x >= hexX(best, false) + g.cw) ? 1 : 0;
        hexSyncEditor(true);
        InvalidateRect(hwnd, nullptr, FALSE);
        return 0;
    }
    case WM_KEYDOWN: {
        const bool ctrl = GetKeyState(VK_CONTROL) < 0;
        const i64 page = (i64)std::max(1, hexVisibleRows() - 1) * 16;
        switch (wp) {
        case VK_LEFT: hexMove(-1); return 0;
        case VK_RIGHT: hexMove(1); return 0;
        case VK_UP: hexMove(-16); return 0;
        case VK_DOWN: hexMove(16); return 0;
        case VK_PRIOR: hexMove(-page); return 0;
        case VK_NEXT: hexMove(page); return 0;
        case VK_HOME: hexMove(ctrl ? -(i64)g_hex.caret : -(i64)(g_hex.caret % 16)); return 0;
        case VK_END: hexMove(ctrl ? (i64)(g.d->doc.size() - g_hex.caret) : (i64)(15 - g_hex.caret % 16)); return 0;
        case VK_TAB:
            g_hex.ascii = !g_hex.ascii;
            g_hex.nibble = 0;
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        case VK_BACK: hexMove(-1); return 0;
        case VK_DELETE:
            if (g_hex.caret < g.d->doc.size()) {
                g.d->doc.breakCoalesce();
                g.d->doc.erase(g_hex.caret, 1);
                hexSyncEditor(true);
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            return 0;
        }
        return 0;
    }
    case WM_CHAR: {
        const wchar_t c = (wchar_t)wp;
        if (c < 0x20) return 0;
        if (g_hex.ascii) {
            if (c > 0xFF) {
                MessageBeep(MB_OK);
                return 0;
            }
            hexPutByte(g_hex.caret, (u8)c);
            hexMove(1);
            return 0;
        }
        const int d = hexDigit(c);
        if (d < 0) {
            MessageBeep(MB_OK);
            return 0;
        }
        u8 b = 0;
        if (g_hex.caret < g.d->doc.size()) g.d->doc.copy(g_hex.caret, &b, 1);
        b = g_hex.nibble == 0 ? (u8)((d << 4) | (b & 0x0F)) : (u8)((b & 0xF0) | d);
        hexPutByte(g_hex.caret, b);
        if (g_hex.nibble == 0) {
            g_hex.nibble = 1;
            hexSyncEditor(true);
            InvalidateRect(hwnd, nullptr, FALSE);
        } else {
            hexMove(1);
        }
        return 0;
    }
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ---- CSV / TSV table

struct CsvState {
    std::vector<u64> rows;       // row starts + end
    std::vector<int> widths;     // column widths in pixels
    char delim = ',';
    bool truncated = false;
    u64 version = NPOS;
    const DocState* doc = nullptr;
    u64 top = 1;                 // first body row shown (row 0 is the header)
    int scrollX = 0;
    u64 row = 1;                 // selected cell
    int col = 0;
    HWND edit = nullptr;         // in-place cell editor
    u64 syncedTop = NPOS;
    int dragCol = -1, dragX = 0, dragW = 0;  // resizing a column
};
CsvState g_csv;

bool isTsvPath() {
    const std::wstring& p = g.d->path;
    return p.size() > 4 && _wcsicmp(p.c_str() + p.size() - 4, L".tsv") == 0;
}

// Row r's text, without its line break.
std::string csvRowText(u64 r) {
    if (r + 1 >= g_csv.rows.size()) return std::string();
    u64 a = g_csv.rows[(size_t)r], b = std::min(g_csv.rows[(size_t)r + 1], g.d->doc.size());
    if (b <= a) return std::string();
    std::string t = g.d->doc.text(a, std::min<u64>(b - a, 1u << 20));
    while (!t.empty() && (t.back() == '\n' || t.back() == '\r')) t.pop_back();
    return t;
}

std::vector<CsvField> csvRowFields(u64 r, std::string& text) {
    text = csvRowText(r);
    std::vector<CsvField> f;
    csvFields(text.data(), text.size(), 0, g_csv.delim, f);
    return f;
}

std::wstring csvCell(const std::string& text, const CsvField& f) {
    std::wstring w = fromDocBytes(csvValue(text.data() + f.start, (size_t)(f.end - f.start)));
    for (wchar_t& c : w)
        if (c == L'\n') c = L'\x21B5';  // line breaks inside a cell: shown as ↵
    w.erase(std::remove(w.begin(), w.end(), L'\r'), w.end());
    return w;
}

int csvRowH() { return g.ch + MulDiv(6, g.dpi, 96); }
int csvGutterW() {
    wchar_t b[32];
    int n = swprintf(b, 32, L"%llu", (unsigned long long)std::max<size_t>(1, g_csv.rows.size()));
    return (n + 2) * g.cw;
}

void rebuildCsv() {
    g_csv.doc = g.d;
    g_csv.version = g.d->doc.version();
    g_csv.rows.clear();
    if (g.d->doc.size() > kSideMax) return;
    std::string text = g.d->doc.text(0, g.d->doc.size());
    g_csv.delim = csvSniff(text.data(), text.size(), isTsvPath());
    g_csv.rows = csvRows(text.data(), text.size(), 5000000, g_csv.truncated);
    // Column widths from the header and the first rows, kept if already set.
    HDC dc = GetDC(g.sideCsv);
    HGDIOBJ of = SelectObject(dc, g.font);
    std::vector<int> w;
    const u64 sample = std::min<u64>(g_csv.rows.size() ? g_csv.rows.size() - 1 : 0, 300);
    for (u64 r = 0; r < sample; r++) {
        std::string t;
        std::vector<CsvField> f = csvRowFields(r, t);
        if (w.size() < f.size()) w.resize(f.size(), 0);
        for (size_t c = 0; c < f.size(); c++) {
            std::wstring v = csvCell(t, f[c]);
            SIZE sz = {};
            GetTextExtentPoint32W(dc, v.c_str(), (int)std::min<size_t>(v.size(), 60), &sz);
            w[c] = std::max(w[c], (int)sz.cx);
        }
    }
    SelectObject(dc, of);
    ReleaseDC(g.sideCsv, dc);
    const int pad = 2 * g.cw;
    for (size_t c = 0; c < w.size(); c++) {
        int v = std::clamp(w[c] + pad, 5 * g.cw, 40 * g.cw);
        if (c < g_csv.widths.size() && g_csv.widths[c] > 0) v = g_csv.widths[c];  // the user resized it
        w[c] = v;
    }
    g_csv.widths = w;
    if (g_csv.widths.empty()) g_csv.widths.push_back(10 * g.cw);
    const u64 body = g_csv.rows.size() > 2 ? g_csv.rows.size() - 2 : 0;  // rows minus header, minus the end entry
    g_csv.top = std::clamp<u64>(g_csv.top, 1, std::max<u64>(1, body));
    g_csv.row = std::clamp<u64>(g_csv.row, 1, std::max<u64>(1, body));
    g_csv.col = std::clamp(g_csv.col, 0, (int)g_csv.widths.size() - 1);
}

u64 csvBodyRows() { return g_csv.rows.size() > 2 ? g_csv.rows.size() - 2 : 0; }
int csvVisibleRows() {
    RECT rc;
    GetClientRect(g.sideCsv, &rc);
    if (csvRowH() <= 0) return 1;
    return std::max(1, ((int)rc.bottom - csvRowH()) / csvRowH());
}
int csvColX(int c) {
    int x = csvGutterW() - g_csv.scrollX;
    for (int k = 0; k < c && k < (int)g_csv.widths.size(); k++) x += g_csv.widths[(size_t)k];
    return x;
}
int csvTotalW() {
    int w = csvGutterW();
    for (int v : g_csv.widths) w += v;
    return w;
}

void csvScrollbars() {
    if (!g.d || !g.sideCsv || g.ch <= 0) return;  // not ready during window creation
    RECT rc;
    GetClientRect(g.sideCsv, &rc);
    SCROLLINFO si = {sizeof(si), SIF_RANGE | SIF_PAGE | SIF_POS};
    si.nMin = 1;
    si.nMax = (int)std::min<u64>(std::max<u64>(1, csvBodyRows()), 0x7FFFFFFF);
    si.nPage = (UINT)csvVisibleRows();
    si.nPos = (int)g_csv.top;
    SetScrollInfo(g.sideCsv, SB_VERT, &si, TRUE);
    si.nMin = 0;
    si.nMax = csvTotalW();
    si.nPage = (UINT)rc.right;
    g_csv.scrollX = std::clamp(g_csv.scrollX, 0, std::max(0, csvTotalW() - (int)rc.right));
    si.nPos = g_csv.scrollX;
    SetScrollInfo(g.sideCsv, SB_HORZ, &si, TRUE);
}

void paintCsv(HWND hwnd) {
    if (g_csv.doc != g.d || g_csv.version != g.d->doc.version()) rebuildCsv();
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(hwnd, &ps);
    RECT rc;
    GetClientRect(hwnd, &rc);
    const int W = rc.right, H = rc.bottom, rh = csvRowH(), gut = csvGutterW(), px = MulDiv(1, g.dpi, 96);
    HDC dc = CreateCompatibleDC(wdc);
    HBITMAP bmp = CreateCompatibleBitmap(wdc, std::max(1, W), std::max(1, H));
    HGDIOBJ ob = SelectObject(dc, bmp);
    const COLORREF bg = g.pal.bg, fg = g.pal.fg, grid = g.pal.border, head = g.pal.face, dim = blend(fg, bg, 45);
    fillRect(dc, 0, 0, W, H, bg);
    HGDIOBJ of = SelectObject(dc, g.font);
    SetBkMode(dc, TRANSPARENT);
    const bool focused = GetFocus() == hwnd || (g_csv.edit && GetFocus() == g_csv.edit);
    auto drawRow = [&](u64 r, int y, bool header) {
        std::string t;
        std::vector<CsvField> f = csvRowFields(r, t);
        if (header) fillRect(dc, 0, y, W, y + rh, head);
        else if (r % 2 == 0) fillRect(dc, gut, y, W, y + rh, blend(fg, bg, g.pal.dark ? 4 : 3));  // zebra rows
        // Row number.
        wchar_t b[32];
        int k = header ? 0 : swprintf(b, 32, L"%llu", (unsigned long long)r);
        RECT nr = {0, y, gut - g.cw / 2, y + rh};
        fillRect(dc, 0, y, gut, y + rh, head);
        SetTextColor(dc, dim);
        if (k) DrawTextW(dc, b, k, &nr, DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
        for (size_t c = 0; c < g_csv.widths.size(); c++) {
            const int x = csvColX((int)c), w = g_csv.widths[c];
            if (x > W) break;
            if (x + w < gut) continue;
            RECT cr = {std::max(x, gut), y, x + w, y + rh};
            const bool isSel = !header && r == g_csv.row && (int)c == g_csv.col;
            if (isSel) fillRect(dc, cr.left, cr.top, cr.right, cr.bottom, focused ? blend(g.pal.accent, bg, 30) : g.pal.selection);
            if (c < f.size()) {
                std::wstring v = csvCell(t, f[c]);
                RECT tr = {x + g.cw, y, x + w - g.cw / 2, y + rh};
                IntersectClipRect(dc, cr.left, cr.top, cr.right, cr.bottom);
                SetTextColor(dc, fg);
                HGDIOBJ hf = nullptr;
                DrawTextW(dc, v.c_str(), (int)v.size(), &tr, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | DT_END_ELLIPSIS);
                SelectClipRgn(dc, nullptr);
                (void)hf;
            }
            fillRect(dc, x + w - 1, y, x + w, y + rh, grid);  // column line
            if (isSel) {  // a square frame around the selected cell
                HBRUSH fb = CreateSolidBrush(g.pal.accent);
                RECT fr = {x, y, x + w, y + rh};
                FrameRect(dc, &fr, fb);
                InflateRect(&fr, -px, -px);
                FrameRect(dc, &fr, fb);
                DeleteObject(fb);
            }
        }
        fillRect(dc, 0, y + rh - 1, W, y + rh, grid);  // row line
        fillRect(dc, gut - 1, y, gut, y + rh, grid);
    };
    if (g_csv.rows.size() >= 2) {
        drawRow(0, 0, true);  // the header stays on top
        int y = rh;
        for (u64 r = g_csv.top; r + 1 < g_csv.rows.size() && y < H; r++, y += rh) drawRow(r, y, false);
    } else {
        RECT t = rc;
        SetTextColor(dc, dim);
        DrawTextW(dc, g.d->doc.size() > kSideMax ? L"The file is too large for the table (limit 64 MB)." : L"Empty",
                  -1, &t, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }
    fillRect(dc, 0, 0, edgeW(), H, grid);
    SelectObject(dc, of);
    BitBlt(wdc, 0, 0, W, H, dc, 0, 0, SRCCOPY);
    SelectObject(dc, ob);
    DeleteObject(bmp);
    DeleteDC(dc);
    EndPaint(hwnd, &ps);
}

// Selects the current cell's text in the editor (keeps the focus in the table).
void csvShowInEditor() {
    if (g_csv.row + 1 >= g_csv.rows.size()) return;
    std::string t;
    std::vector<CsvField> f = csvRowFields(g_csv.row, t);
    const u64 base = g_csv.rows[(size_t)g_csv.row];
    u64 a = base, b = base;
    if (g_csv.col < (int)f.size()) {
        a = base + f[(size_t)g_csv.col].start;
        b = base + f[(size_t)g_csv.col].end;
    }
    g.d->caret = std::min(a, g.d->doc.size());
    g.d->anchor = std::min(b, g.d->doc.size());
    g.d->wantCol = -1;
    g_csv.syncedTop = NPOS;
    refresh(true);
}

void csvEnsureVisible() {
    const u64 vis = (u64)csvVisibleRows();
    if (g_csv.row < g_csv.top) g_csv.top = g_csv.row;
    else if (g_csv.row >= g_csv.top + vis) g_csv.top = g_csv.row - vis + 1;
    RECT rc;
    GetClientRect(g.sideCsv, &rc);
    const int x = csvColX(g_csv.col) + g_csv.scrollX, w = g_csv.widths[(size_t)g_csv.col], gut = csvGutterW();
    if (x - g_csv.scrollX < gut) g_csv.scrollX = std::max(0, x - gut);
    else if (x + w - g_csv.scrollX > rc.right) g_csv.scrollX = x + w - (int)rc.right;
}

void csvMove(i64 dr, int dc) {
    const u64 body = csvBodyRows();
    if (!body) return;
    g_csv.row = (u64)std::clamp<i64>((i64)g_csv.row + dr, 1, (i64)body);
    g_csv.col = std::clamp(g_csv.col + dc, 0, (int)g_csv.widths.size() - 1);
    csvEnsureVisible();
    csvScrollbars();
    InvalidateRect(g.sideCsv, nullptr, FALSE);
    csvShowInEditor();
}

// The editor moved: the table follows (the row holding the editor's top line).
void syncCsvToEditor() {
    if (!g.sideCsv || !IsWindowVisible(g.sideCsv) || !g.d->doc.linesKnown() || g_csv.rows.size() < 3) return;
    const u64 line = g.d->doc.lineOf(g.d->top);
    if (line == g_csv.syncedTop) return;
    g_csv.syncedTop = line;
    const u64 off = g.d->top;
    auto it = std::upper_bound(g_csv.rows.begin(), g_csv.rows.end() - 1, off);
    u64 r = (u64)(it - g_csv.rows.begin());
    r = r ? r - 1 : 0;
    g_csv.top = std::clamp<u64>(r, 1, std::max<u64>(1, csvBodyRows()));
    csvScrollbars();
    InvalidateRect(g.sideCsv, nullptr, FALSE);
}

// The table moved: the editor follows.
void csvSyncEditor() {
    if (g_csv.top + 1 >= g_csv.rows.size()) return;
    g.d->top = std::min(g_csv.rows[(size_t)g_csv.top], g.d->doc.size());
    g.d->topFill = 0;
    g_csv.syncedTop = g.d->doc.linesKnown() ? g.d->doc.lineOf(g.d->top) : NPOS;
    refresh(false);
}

void csvCommitEdit(bool keep) {
    if (!g_csv.edit) return;
    HWND e = g_csv.edit;
    g_csv.edit = nullptr;
    if (keep) {
        if (g_csv.version != g.d->doc.version()) rebuildCsv();
        std::string t;
        std::vector<CsvField> f = csvRowFields(g_csv.row, t);
        const u64 base = g_csv.rows[(size_t)g_csv.row];
        std::wstring w = editText(e);
        std::replace(w.begin(), w.end(), L'\x21B5', L'\n');  // ↵ back to a line break
        std::string raw = csvQuote(toDocBytes(w.c_str(), (int)w.size()), g_csv.delim);
        if (g_csv.col < (int)f.size()) {
            const u64 a = base + f[(size_t)g_csv.col].start, b = base + f[(size_t)g_csv.col].end;
            if (g.d->doc.text(a, b - a) != raw) replaceRange(a, b, raw, a, a + raw.size());
        } else {
            // A cell past the row's last field: add the missing delimiters.
            std::string add(1, g_csv.delim);
            for (int k = (int)f.size(); k < g_csv.col; k++) add += g_csv.delim;
            const u64 at = base + t.size();
            replaceRange(at, at, add + raw, at, at + add.size() + raw.size());
        }
        rebuildCsv();
    }
    DestroyWindow(e);
    SetFocus(g.sideCsv);
    InvalidateRect(g.sideCsv, nullptr, FALSE);
}

LRESULT CALLBACK csvEditProc(HWND e, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR, DWORD_PTR) {
    if (msg == WM_KEYDOWN) {
        if (wp == VK_RETURN) {
            csvCommitEdit(true);
            csvMove(1, 0);
            return 0;
        }
        if (wp == VK_ESCAPE) {
            csvCommitEdit(false);
            return 0;
        }
        if (wp == VK_TAB) {
            csvCommitEdit(true);
            csvMove(0, GetKeyState(VK_SHIFT) < 0 ? -1 : 1);
            return 0;
        }
    }
    if (msg == WM_CHAR && (wp == L'\r' || wp == 27 || wp == L'\t')) return 0;
    if (msg == WM_KILLFOCUS && g_csv.edit == e) {
        csvCommitEdit(true);
        return 0;
    }
    return DefSubclassProc(e, msg, wp, lp);
}

void csvBeginEdit() {
    if (g_csv.edit || g_csv.row + 1 >= g_csv.rows.size()) return;
    csvEnsureVisible();
    csvScrollbars();
    std::string t;
    std::vector<CsvField> f = csvRowFields(g_csv.row, t);
    std::wstring v = g_csv.col < (int)f.size() ? csvCell(t, f[(size_t)g_csv.col]) : L"";
    const int rh = csvRowH(), x = csvColX(g_csv.col), y = rh + (int)(g_csv.row - g_csv.top) * rh;
    g_csv.edit = CreateWindowExW(0, L"EDIT", v.c_str(), WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, x + 1, y + 1,
                                 g_csv.widths[(size_t)g_csv.col] - 2, rh - 2, g.sideCsv, nullptr, g.inst, nullptr);
    SendMessageW(g_csv.edit, WM_SETFONT, (WPARAM)g.font, TRUE);
    SetWindowTheme(g_csv.edit, L"", L"");
    SetWindowSubclass(g_csv.edit, csvEditProc, 1, 0);
    SetFocus(g_csv.edit);
    SendMessageW(g_csv.edit, EM_SETSEL, 0, -1);
}

bool csvHit(int x, int y, u64& row, int& col, bool& onBorder) {
    const int rh = csvRowH();
    onBorder = false;
    col = -1;
    for (int c = 0; c < (int)g_csv.widths.size(); c++) {
        const int x0 = csvColX(c), x1 = x0 + g_csv.widths[(size_t)c];
        if (y < rh && std::abs(x - x1) <= 3) {  // the header's column edge: resize
            onBorder = true;
            col = c;
            return true;
        }
        if (x >= x0 && x < x1) col = c;
    }
    if (col < 0 || x < csvGutterW()) return false;
    if (y < rh) return false;  // header
    row = g_csv.top + (u64)((y - rh) / rh);
    return row >= 1 && row <= csvBodyRows();
}

LRESULT CALLBACK csvProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg >= WM_CTLCOLORMSGBOX && msg <= WM_CTLCOLORSTATIC) return themeCtlColor(msg, wp, g.pal);
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paintCsv(hwnd); return 0;
    case WM_SIZE: csvScrollbars(); return 0;
    case WM_SETFOCUS:
    case WM_KILLFOCUS: InvalidateRect(hwnd, nullptr, FALSE); return 0;
    case WM_GETDLGCODE: return DLGC_WANTALLKEYS;
    case WM_MOUSEWHEEL: {
        int lines = 3;
        SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
        csvCommitEdit(true);
        i64 t = (i64)g_csv.top - (i64)GET_WHEEL_DELTA_WPARAM(wp) * lines / WHEEL_DELTA;
        g_csv.top = (u64)std::clamp<i64>(t, 1, std::max<i64>(1, (i64)csvBodyRows()));
        csvScrollbars();
        InvalidateRect(hwnd, nullptr, FALSE);
        csvSyncEditor();
        return 0;
    }
    case WM_VSCROLL:
    case WM_HSCROLL: {
        csvCommitEdit(true);
        const bool v = msg == WM_VSCROLL;
        SCROLLINFO si = {sizeof(si), SIF_ALL};
        GetScrollInfo(hwnd, v ? SB_VERT : SB_HORZ, &si);
        i64 pos = v ? (i64)g_csv.top : g_csv.scrollX, page = v ? csvVisibleRows() - 1 : (i64)si.nPage * 9 / 10,
            line = v ? 1 : g.cw * 4;
        switch (LOWORD(wp)) {
        case SB_LINEUP: pos -= line; break;
        case SB_LINEDOWN: pos += line; break;
        case SB_PAGEUP: pos -= std::max<i64>(1, page); break;
        case SB_PAGEDOWN: pos += std::max<i64>(1, page); break;
        case SB_THUMBTRACK:
        case SB_THUMBPOSITION: pos = si.nTrackPos; break;
        default: return 0;
        }
        if (v) g_csv.top = (u64)std::clamp<i64>(pos, 1, std::max<i64>(1, (i64)csvBodyRows()));
        else g_csv.scrollX = (int)std::max<i64>(0, pos);
        csvScrollbars();
        InvalidateRect(hwnd, nullptr, FALSE);
        if (v) {
            csvSyncEditor();
            UpdateWindow(g.view);
        }
        return 0;
    }
    case WM_SETCURSOR: {
        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(hwnd, &pt);
        u64 r;
        int c;
        bool border;
        csvHit(pt.x, pt.y, r, c, border);
        if (border || g_csv.dragCol >= 0) {
            SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
            return TRUE;
        }
        break;
    }
    case WM_LBUTTONDOWN:
    case WM_LBUTTONDBLCLK: {
        csvCommitEdit(true);
        SetFocus(hwnd);
        u64 r = 0;
        int c;
        bool border;
        const int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp);
        if (csvHit(x, y, r, c, border)) {
            if (border) {  // drag a column edge
                g_csv.dragCol = c;
                g_csv.dragX = x;
                g_csv.dragW = g_csv.widths[(size_t)c];
                SetCapture(hwnd);
                return 0;
            }
            g_csv.row = r;
            g_csv.col = c;
            InvalidateRect(hwnd, nullptr, FALSE);
            csvShowInEditor();
            if (msg == WM_LBUTTONDBLCLK) csvBeginEdit();
        }
        return 0;
    }
    case WM_MOUSEMOVE:
        if (g_csv.dragCol >= 0 && GetCapture() == hwnd) {
            g_csv.widths[(size_t)g_csv.dragCol] = std::max(3 * g.cw, g_csv.dragW + GET_X_LPARAM(lp) - g_csv.dragX);
            csvScrollbars();
            InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
    case WM_LBUTTONUP:
        if (g_csv.dragCol >= 0) {
            g_csv.dragCol = -1;
            ReleaseCapture();
        }
        return 0;
    case WM_KEYDOWN: {
        const bool ctrl = GetKeyState(VK_CONTROL) < 0, shift = GetKeyState(VK_SHIFT) < 0;
        switch (wp) {
        case VK_UP: csvMove(-1, 0); return 0;
        case VK_DOWN: csvMove(1, 0); return 0;
        case VK_LEFT: csvMove(0, -1); return 0;
        case VK_RIGHT: csvMove(0, 1); return 0;
        case VK_TAB: csvMove(0, shift ? -1 : 1); return 0;
        case VK_PRIOR: csvMove(-(csvVisibleRows() - 1), 0); return 0;
        case VK_NEXT: csvMove(csvVisibleRows() - 1, 0); return 0;
        case VK_HOME: csvMove(ctrl ? -(i64)g_csv.row : 0, -g_csv.col); return 0;
        case VK_END: csvMove(ctrl ? (i64)csvBodyRows() : 0, (int)g_csv.widths.size()); return 0;
        case VK_RETURN:
        case VK_F2: csvBeginEdit(); return 0;
        case VK_DELETE:  // empties the cell
            csvBeginEdit();
            if (g_csv.edit) {
                SetWindowTextW(g_csv.edit, L"");
                csvCommitEdit(true);
            }
            return 0;
        }
        return 0;
    }
    case WM_CHAR:
        // Typing starts editing the cell with that character.
        if (wp >= 0x20 && !g_csv.edit) {
            csvBeginEdit();
            if (g_csv.edit) {
                wchar_t s[2] = {(wchar_t)wp, 0};
                SetWindowTextW(g_csv.edit, s);
                SendMessageW(g_csv.edit, EM_SETSEL, 1, 1);
            }
        }
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ---- the pane as a whole

void themeSidePane() {
    if (!g.sideTree) return;
    SetWindowTheme(g.sideTree, g.pal.dark ? L"DarkMode_Explorer" : L"Explorer", nullptr);
    TreeView_SetBkColor(g.sideTree, g.pal.bg);
    TreeView_SetTextColor(g.sideTree, g.pal.fg);
    TreeView_SetLineColor(g.sideTree, blend(g.pal.fg, g.pal.bg, 35));
    SendMessageW(g.sideTree, WM_SETFONT, (WPARAM)g.font, TRUE);
    SetWindowTheme(g.sidePreview, g.pal.dark ? L"DarkMode_Explorer" : nullptr, nullptr);
    for (HWND w : {g.sideHex, g.sideCsv}) {
        SetWindowTheme(w, g.pal.dark ? L"DarkMode_Explorer" : nullptr, nullptr);
        InvalidateRect(w, nullptr, FALSE);
    }
    g_md.width = -1;
}

// Brings the pane up to date with the open tab (after switching tab, opening,
// toggling, or an edit settled).
void updateSidePane() {
    if (!g.sideTree) return;
    const bool vis = sideVisible();
    const int kind = sideKind();
    ShowWindow(g.sideTree, vis && (kind == SIDE_JSON || kind == SIDE_XML) ? SW_SHOW : SW_HIDE);
    if (g_dataBar.wnd) ShowWindow(g_dataBar.wnd, vis && (kind == SIDE_JSON || kind == SIDE_XML) ? SW_SHOW : SW_HIDE);
    ShowWindow(g.sidePreview, vis && kind == SIDE_MD ? SW_SHOW : SW_HIDE);
    ShowWindow(g.sideHex, vis && kind == SIDE_HEX ? SW_SHOW : SW_HIDE);
    ShowWindow(g.sideCsv, vis && kind == SIDE_CSV ? SW_SHOW : SW_HIDE);
    if (vis && kind == SIDE_HEX) {
        g_hex.caret = std::min(g.d->caret, g.d->doc.size());
        g_hex.syncedTop = NPOS;
        syncHexToEditor();
        hexScrollbar();
        InvalidateRect(g.sideHex, nullptr, FALSE);
        return;
    }
    if (vis && kind == SIDE_CSV) {
        if (g_csv.doc != g.d || g_csv.version != g.d->doc.version()) {
            if (g_csv.doc != g.d) g_csv = CsvState();  // another file: fresh widths and selection
            rebuildCsv();
        }
        g_csv.syncedTop = NPOS;
        syncCsvToEditor();
        csvScrollbars();
        InvalidateRect(g.sideCsv, nullptr, FALSE);
        return;
    }
    if (!vis) return;
    if (kind == SIDE_MD) {
        g.mdSyncedTop = NPOS;
        InvalidateRect(g.sidePreview, nullptr, FALSE);
        syncPreviewToEditor();
    } else if (g.dataDoc != g.d || g.dataVersion != g.d->doc.version()) {
        rebuildDataView();
    }
}

void toggleSidePane() {
    if (sideKind() == SIDE_NONE) {
        MessageBeep(MB_OK);
        return;
    }
    g.sideOn = !(g.sideOn && !g.diffOn);
    if (g.sideOn && g.diffOn) toggleDiff();  // one or the other beside the editor
    layoutFrame();
    updateSidePane();
    saveSettings();
}

// WM_NOTIFY from the tree.
LRESULT onTreeNotify(NMHDR* nm) {
    switch (nm->code) {
    case TVN_ITEMEXPANDINGW: {
        NMTREEVIEWW* tv = (NMTREEVIEWW*)nm;
        if (tv->action == TVE_EXPAND) fillKids(tv->itemNew.hItem, (int)tv->itemNew.lParam);
        return FALSE;
    }
    case TVN_SELCHANGEDW: {
        NMTREEVIEWW* tv = (NMTREEVIEWW*)nm;
        if (!g.treeUpdating && tv->action != TVC_UNKNOWN) showNodeInEditor((int)tv->itemNew.lParam);
        return 0;
    }
    case TVN_KEYDOWN: {
        NMTVKEYDOWN* k = (NMTVKEYDOWN*)nm;
        if (k->wVKey == VK_F2) {
            if (HTREEITEM s = TreeView_GetSelection(g.sideTree)) TreeView_EditLabel(g.sideTree, s);
            return TRUE;
        }
        if (k->wVKey == VK_RETURN) {
            showNodeInEditor(treeNodeOf(TreeView_GetSelection(g.sideTree)));
            SetFocus(g.view);
            return TRUE;
        }
        return 0;
    }
    case NM_DBLCLK: {
        // Leaves: double click edits the value. Branches open and close as usual.
        HTREEITEM s = TreeView_GetSelection(g.sideTree);
        int id = treeNodeOf(s);
        if (id >= 0 && (size_t)id < g.data.nodes.size() && g.data.nodes[(size_t)id].kids.empty() &&
            nodeEditable(g.data.nodes[(size_t)id])) {
            TreeView_EditLabel(g.sideTree, s);
            return TRUE;
        }
        return 0;
    }
    case TVN_BEGINLABELEDITW: {
        NMTVDISPINFOW* di = (NMTVDISPINFOW*)nm;
        int id = (int)di->item.lParam;
        if (g.dataVersion != g.d->doc.version()) {  // the tree is behind the text: catch up first
            rebuildDataView();
            return TRUE;
        }
        if (id < 0 || (size_t)id >= g.data.nodes.size() || !nodeEditable(g.data.nodes[(size_t)id])) return TRUE;
        if (HWND e = TreeView_GetEditControl(g.sideTree))
            SetWindowTextW(e, fromDocBytes(g.data.nodes[(size_t)id].value).c_str());
        return FALSE;
    }
    case TVN_ENDLABELEDITW: {
        NMTVDISPINFOW* di = (NMTVDISPINFOW*)nm;
        if (di->item.pszText) applyNodeEdit((int)di->item.lParam, di->item.pszText);
        return FALSE;  // the tree is rebuilt from the new text, labels included
    }
    }
    return 0;
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
    themeSidePane();
    themeBars();
    if (g.sidePreview) InvalidateRect(g.sidePreview, nullptr, FALSE);
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
    g.recentMenu = CreatePopupMenu();
    AppendMenuW(g.recentMenu, MF_STRING | MF_GRAYED, IDM_RECENT0, L"No recent files");
    AppendMenuW(file, MF_POPUP, (UINT_PTR)g.recentMenu, L"Open &Recent");
    AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
    g.encMenu = CreatePopupMenu();
    g.eolMenu = CreatePopupMenu();
    addEncodingItems(g.encMenu, g.eolMenu);
    AppendMenuW(file, MF_POPUP, (UINT_PTR)g.encMenu, L"&Encoding");
    AppendMenuW(file, MF_POPUP, (UINT_PTR)g.eolMenu, L"&Line Endings");
    AppendMenuW(file, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(file, MF_STRING, IDM_PRINT, L"&Print...\tCtrl+P");
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
    AppendMenuW(line, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(line, MF_STRING, IDM_SORT_ASC, L"&Sort Lines");
    AppendMenuW(line, MF_STRING, IDM_SORT_DESC, L"Sort Lines Des&cending");
    AppendMenuW(line, MF_STRING, IDM_DEDUPE, L"Remove Duplicate &Lines");
    AppendMenuW(line, MF_STRING, IDM_TRIM, L"&Trim Trailing Whitespace");
    AppendMenuW(edit, MF_POPUP, (UINT_PTR)line, L"&Line");
    HMENU carets = CreatePopupMenu();
    AppendMenuW(carets, MF_STRING, IDM_ADDNEXT, L"Add Caret at &Next Match\tShift+Alt+.");
    AppendMenuW(carets, MF_STRING, IDM_ADDALL, L"Add Carets at &All Matches\tShift+Alt+;");
    AppendMenuW(carets, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(carets, MF_STRING, IDM_BOX_UP, L"Column Select &Up\tShift+Alt+Up");
    AppendMenuW(carets, MF_STRING, IDM_BOX_DOWN, L"Column Select &Down\tShift+Alt+Down");
    AppendMenuW(carets, MF_STRING, IDM_BOX_LEFT, L"Column Select &Left\tShift+Alt+Left");
    AppendMenuW(carets, MF_STRING, IDM_BOX_RIGHT, L"Column Select &Right\tShift+Alt+Right");
    AppendMenuW(edit, MF_POPUP, (UINT_PTR)carets, L"&Multiple Carets");
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
    AppendMenuW(view, MF_STRING, IDM_SHOWWS, L"Show White&space");
    AppendMenuW(view, MF_STRING, IDM_SIDEVIEW, L"Data View / &Preview\tCtrl+Shift+V");
    AppendMenuW(view, MF_STRING, IDM_HEXVIEW, L"Show as He&x");
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
    AppendMenuW(sc, MF_STRING, IDM_VCS_HISTORY, L"File &History...\tCtrl+Shift+H");
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
        {FVIRTKEY | FCONTROL | FSHIFT, 'V', IDM_SIDEVIEW},
        {FVIRTKEY | FCONTROL | FSHIFT, 'H', IDM_VCS_HISTORY},
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
        {FVIRTKEY | FCONTROL, 'P', IDM_PRINT},
        {FVIRTKEY | FALT | FSHIFT, VK_OEM_PERIOD, IDM_ADDNEXT},
        {FVIRTKEY | FALT | FSHIFT, VK_OEM_1, IDM_ADDALL},
        {FVIRTKEY | FALT | FSHIFT, VK_UP, IDM_BOX_UP},
        {FVIRTKEY | FALT | FSHIFT, VK_DOWN, IDM_BOX_DOWN},
        {FVIRTKEY | FALT | FSHIFT, VK_LEFT, IDM_BOX_LEFT},
        {FVIRTKEY | FALT | FSHIFT, VK_RIGHT, IDM_BOX_RIGHT},
    };
    return CreateAcceleratorTableW(a, (int)(sizeof(a) / sizeof(a[0])));
}

void layoutFrame() {
    RECT rc;
    GetClientRect(g.frame, &rc);
    int h = std::max<int>(0, rc.bottom - g.statusH);
    const int headerTop = tabBarH() + headerH();
    // The find bar sits right above the text, across the window.
    const int fb = barHeight(g_findBar);
    if (g_findBar.wnd && fb) MoveWindow(g_findBar.wnd, 0, headerTop, rc.right, fb, TRUE);
    const int top = headerTop + fb;
    int left = 0;
    int right = rc.right;
    SetRectEmpty(&g.sideSplitRect);
    if (g.sideSplitter && !sideVisible()) ShowWindow(g.sideSplitter, SW_HIDE);
    if (sideVisible()) {
        // The data view / preview on the right, after a divider you can drag.
        const int min = MulDiv(160, g.dpi, 96), gap = splitW();
        int sw = (int)(rc.right * g.sideSplit);
        sw = std::clamp(sw, std::min(min, (int)rc.right / 2), std::max(0, (int)rc.right - min - gap));
        right = rc.right - sw - gap;
        g.sideSplitRect = {right, top, right + gap, h};
        if (g.sideSplitter) {
            MoveWindow(g.sideSplitter, right, top, gap, h - top, TRUE);
            ShowWindow(g.sideSplitter, SW_SHOW);
        }
        // The data view's search box on top of its tree.
        const int sb = barRowH();
        if (g_dataBar.wnd) MoveWindow(g_dataBar.wnd, right + gap, top, sw, sb, TRUE);
        MoveWindow(g.sideTree, right + gap, top + sb, sw, std::max(0, h - top - sb), TRUE);
        MoveWindow(g.sidePreview, right + gap, top, sw, h - top, TRUE);
        MoveWindow(g.sideHex, right + gap, top, sw, h - top, TRUE);
        MoveWindow(g.sideCsv, right + gap, top, sw, h - top, TRUE);
    }
    if (g.diffOn) {
        // The strip in the middle can be dragged: the left side gets diffSplit of the rest.
        int sw = stripW();
        const int avail = std::max(0, right - sw), min = MulDiv(120, g.dpi, 96);
        left = std::clamp((int)(avail * g.diffSplit), std::min(min, avail / 2), std::max(0, avail - min));
        MoveWindow(g.diffPane, 0, top, left, h - top, TRUE);
        MoveWindow(g.strip, left, top, sw, h - top, TRUE);
        left += sw;
    }
    g.sideShown = sideVisible();
    MoveWindow(g.view, left, top, right - left, h - top, TRUE);
    RECT s = {0, h, rc.right, rc.bottom};
    InvalidateRect(g.frame, &s, FALSE);
    RECT hd = {0, 0, rc.right, headerTop};
    InvalidateRect(g.frame, &hd, FALSE);
    if (!IsRectEmpty(&g.sideSplitRect)) InvalidateRect(g.frame, &g.sideSplitRect, FALSE);
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
    layoutFrame();
    updateSidePane();
    SetFocus(g.view);
    PostMessageW(g.frame, WM_APP_DISKCHECK, 0, 0);  // changed on disk while in the background?
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
    addRecent(path);
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
    themeSidePane();
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
    themeSidePane();
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

// Where the caret is drawn, in view coordinates (for the keyboard context menu).
POINT caretPoint() {
    POINT pt = {g.gutterPx, 0};
    for (size_t i = 0; i < g.rows.size(); i++) {
        if (!g.rowFill[i] && rowHas(g.d->doc, g.rows[i], g.d->caret)) {
            int col = colOf(g.d->doc, g.rows[i], g.d->caret, cfg());
            pt.x = std::max(g.gutterPx, g.gutterPx + (col - g.d->hcol) * g.cw);
            pt.y = (int)(i + 1) * g.ch;
            break;
        }
    }
    return pt;
}

// The right-click menu: everything that applies to the selection or the caret.
void showEditorMenu(POINT screen) {
    const bool sel = hasSel(), ro = false;
    (void)ro;
    HMENU m = CreatePopupMenu();
    auto item = [&](HMENU into, UINT id, const wchar_t* text, bool enabled = true) {
        AppendMenuW(into, MF_STRING | (enabled ? 0 : MF_GRAYED), id, text);
    };
    auto sep = [&](HMENU into) { AppendMenuW(into, MF_SEPARATOR, 0, nullptr); };
    item(m, IDM_UNDO, L"&Undo\tCtrl+Z", g.d->doc.canUndo());
    item(m, IDM_REDO, L"&Redo\tCtrl+Y", g.d->doc.canRedo());
    sep(m);
    item(m, IDM_CUT, sel ? L"Cu&t\tCtrl+X" : L"Cu&t Line\tCtrl+X");
    item(m, IDM_COPY, sel ? L"&Copy\tCtrl+C" : L"&Copy Line\tCtrl+C");
    item(m, IDM_PASTE, L"&Paste\tCtrl+V", IsClipboardFormatAvailable(CF_UNICODETEXT) != 0);
    item(m, IDM_DELETE, L"&Delete\tDel", sel);
    sep(m);
    item(m, IDM_SELECTALL, L"Select &All\tCtrl+A");
    item(m, IDM_SELECTWORD, L"Select &Word");
    sep(m);
    // Selection: case, find more of it.
    item(m, IDM_UPPER, L"Make U&ppercase\tCtrl+Shift+U", sel);
    item(m, IDM_LOWER, L"Make &Lowercase\tCtrl+U", sel);
    item(m, IDM_FINDWORD, sel ? L"Find &Next of Selection\tCtrl+F3" : L"Find &Next of Word\tCtrl+F3");
    item(m, IDM_FINDWORD_PREV, sel ? L"Find Pre&vious of Selection\tCtrl+Shift+F3" : L"Find Pre&vious of Word\tCtrl+Shift+F3");
    sep(m);
    HMENU line = CreatePopupMenu();
    item(line, IDM_LINE_DUP, sel ? L"D&uplicate Selection\tCtrl+D" : L"D&uplicate Line\tCtrl+D");
    item(line, IDM_LINE_DELETE, L"&Delete Line\tCtrl+Shift+L");
    sep(line);
    item(line, IDM_LINE_UP, L"Move &Up\tAlt+Up");
    item(line, IDM_LINE_DOWN, L"Move Dow&n\tAlt+Down");
    sep(line);
    item(line, IDM_OUTDENT, L"&Outdent\tShift+Tab");
    item(line, IDM_LINE_ABOVE, L"Insert Line &Above\tCtrl+Enter");
    item(line, IDM_LINE_BELOW, L"Insert Line &Below\tCtrl+Shift+Enter");
    AppendMenuW(m, MF_POPUP, (UINT_PTR)line, sel ? L"&Lines" : L"&Line");
    item(m, IDM_BRACKET, L"Go to Matching &Bracket\tCtrl+]");
    // Version control and views, when they apply to this file.
    const bool vcs = g.d->vcs.kind != VcsKind::None && !g.d->vcs.toolMissing;
    if (vcs || sideKind() != SIDE_NONE) sep(m);
    if (vcs) {
        item(m, IDM_DIFF, g.diffOn ? L"Hide &Diff\tF8" : L"Show &Diff\tF8");
        item(m, IDM_VCS_HISTORY, L"File &History...\tCtrl+Shift+H", g.d->vcs.tracked);
        if (g.diffOn) item(m, IDM_REVERTCHANGE, L"Use Depot &Version of This Change\tCtrl+R", !g.d->hunks.empty());
    }
    if (sideKind() == SIDE_MD) item(m, IDM_SIDEVIEW, sideVisible() ? L"Hide Markdown &Preview\tCtrl+Shift+V" : L"Markdown &Preview\tCtrl+Shift+V");
    else if (sideKind() != SIDE_NONE) item(m, IDM_SIDEVIEW, sideVisible() ? L"Hide Data &View\tCtrl+Shift+V" : L"Data &View\tCtrl+Shift+V");

    themeOwnerDrawPopup(m);
    themePopupColors(m, g.pal);
    // Commands come back here; the frame draws the items (owner-drawn).
    UINT cmd = (UINT)TrackPopupMenuEx(m, TPM_RETURNCMD | TPM_RIGHTBUTTON | TPM_NONOTIFY, screen.x, screen.y, g.frame, nullptr);
    themeReleasePopup(m);
    DestroyMenu(m);
    if (cmd) onCommand((int)cmd);
}

void onCommand(int id) {
    // Commands that work with several carets keep them; the rest go back to one.
    switch (id) {
    case IDM_COPY:
    case IDM_CUT:
    case IDM_PASTE:
    case IDM_DELETE:
    case IDM_ADDNEXT:
    case IDM_ADDALL:
    case IDM_BOX_UP:
    case IDM_BOX_DOWN:
    case IDM_BOX_LEFT:
    case IDM_BOX_RIGHT:
    case IDM_SHOWWS:
    case IDM_ZOOMIN:
    case IDM_ZOOMOUT:
    case IDM_ZOOMRESET: break;
    default:
        if (multi()) clearExtras();
        g.boxActive = false;
    }
    if (id >= IDM_RECENT0 && id < IDM_RECENT0 + 12) {
        openRecent((size_t)(id - IDM_RECENT0));
        return;
    }
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
        if (multi()) {
            if (multiCopy()) multiInsert("");
            break;
        }
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
    case IDM_SIDEVIEW: toggleSidePane(); break;
    case IDM_HEXVIEW:
        g.d->forceHex = !g.d->forceHex;
        if (g.d->forceHex) {
            g.sideOn = true;
            if (g.diffOn) toggleDiff();
        }
        layoutFrame();
        updateSidePane();
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
    case IDM_VCS_HISTORY: cmdHistory(); break;
    case IDM_ABOUT: DialogBoxParamW(g.inst, MAKEINTRESOURCEW(IDD_ABOUT), g.frame, aboutProc, 0); break;
    case IDM_PRINT: cmdPrint(); break;
    case IDM_RECENTCLEAR:
        g.recent.clear();
        saveRecent();
        break;
    case IDM_ENC_UTF8:
    case IDM_ENC_UTF8BOM:
    case IDM_ENC_UTF16LE:
    case IDM_ENC_UTF16BE:
    case IDM_ENC_ANSI: setEncoding((FileEnc)(id - IDM_ENC_UTF8)); break;
    case IDM_REINTERP_UTF8: reinterpretAs(FE_UTF8); break;
    case IDM_REINTERP_ANSI: reinterpretAs(FE_ANSI); break;
    case IDM_EOL_CRLF: setLineEndings(true); break;
    case IDM_EOL_LF: setLineEndings(false); break;
    case IDM_SORT_ASC: sortLines(0); break;
    case IDM_SORT_DESC: sortLines(1); break;
    case IDM_DEDUPE: sortLines(2); break;
    case IDM_TRIM: trimTrailing(); break;
    case IDM_ADDNEXT: addNextMatch(false); break;
    case IDM_ADDALL: addNextMatch(true); break;
    case IDM_BOX_UP: boxMove(-1, 0); break;
    case IDM_BOX_DOWN: boxMove(1, 0); break;
    case IDM_BOX_LEFT: boxMove(0, -1); break;
    case IDM_BOX_RIGHT: boxMove(0, 1); break;
    case IDM_SHOWWS:
        g.showWs = !g.showWs;
        saveSettings();
        InvalidateRect(g.view, nullptr, FALSE);
        break;
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
    {
        const int k = sideKind();
        themeSetMenuText(g.menu, IDM_SIDEVIEW, k == SIDE_MD                     ? L"Markdown &Preview\tCtrl+Shift+V"
                                               : k == SIDE_JSON || k == SIDE_XML ? L"&Data View\tCtrl+Shift+V"
                                               : k == SIDE_HEX                   ? L"&Hex View\tCtrl+Shift+V"
                                               : k == SIDE_CSV                   ? L"&Table View\tCtrl+Shift+V"
                                                                                 : L"Data View / &Preview\tCtrl+Shift+V");
        CheckMenuItem(g.menu, IDM_HEXVIEW, MF_BYCOMMAND | (g.d->forceHex || g.d->binary ? MF_CHECKED : MF_UNCHECKED));
        en(IDM_HEXVIEW, !g.d->binary);
        en(IDM_SIDEVIEW, k != SIDE_NONE);
        CheckMenuItem(g.menu, IDM_SIDEVIEW, MF_BYCOMMAND | (sideVisible() ? MF_CHECKED : MF_UNCHECKED));
    }
    en(IDM_WRAP, !g.diffOn);
    CheckMenuItem(g.menu, IDM_DIFF, MF_BYCOMMAND | (g.diffOn ? MF_CHECKED : MF_UNCHECKED));
    const bool p4 = g.d->vcs.kind == VcsKind::Perforce;
    const bool vcs = g.d->vcs.kind != VcsKind::None && !g.d->vcs.toolMissing;
    themeSetMenuText(g.menu, IDM_VCS_COMMIT, p4 || !vcs ? L"&Submit Current Changelist..." : L"&Commit All Changes...");
    en(IDM_VCS_COMMIT, vcs && (!p4 || g.d->vcs.opened));
    en(IDM_VCS_REVERTFILE, vcs && g.d->vcs.tracked);
    en(IDM_VCS_REFRESH, !g.d->path.empty());
    en(IDM_VCS_FILES, vcs);
    en(IDM_VCS_HISTORY, vcs && g.d->vcs.tracked);
    for (int k = 0; k < 3; k++)
        CheckMenuItem(g.menu, IDM_THEME_SYSTEM + k, MF_BYCOMMAND | (g.themeMode == k ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(g.menu, IDM_SHOWWS, MF_BYCOMMAND | (g.showWs ? MF_CHECKED : MF_UNCHECKED));
    checkEncodingItems(g.menu);
}

// ====================================================================== view window

void onKeyDown(WPARAM vk) {
    const bool ctrl = GetKeyState(VK_CONTROL) < 0;
    const bool shift = GetKeyState(VK_SHIFT) < 0;
    LayoutCfg c = cfg();
    if (multi()) {
        if (vk == VK_ESCAPE) {
            clearExtras();
            refresh(false);
            return;
        }
        if (multiKey(vk, ctrl, shift)) return;
        // Typed characters arrive as WM_CHAR after this; only moves that
        // several carets do not do go back to one.
        if (vk == VK_PRIOR || vk == VK_NEXT || vk == VK_HOME || vk == VK_END || vk == VK_UP || vk == VK_DOWN) clearExtras();
    }
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
    case VK_ESCAPE:
        if (findBarOpen()) closeFindBar();
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
    if (ch == L'\t' && multi()) {
        typeChar('\t');
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
        const int mx = GET_X_LPARAM(lp), my = GET_Y_LPARAM(lp);
        const bool alt = GetKeyState(VK_MENU) < 0;
        u64 p = posFromPoint(mx, my);
        if (alt && (wp & MK_CONTROL)) {  // Ctrl+Alt+click: one more caret
            g.altClicked = true;
            addCaretAt(p);
            return 0;
        }
        if (alt) {  // Alt+drag: column selection
            u64 line;
            int col;
            g.altClicked = true;
            if (boxPoint(mx, my, line, col)) {
                g.boxL0 = g.boxL1 = line;
                g.boxC0 = g.boxC1 = col;
                g.boxDrag = true;
                SetCapture(hwnd);
                boxRebuild();
                refresh(false);
            }
            return 0;
        }
        clearExtras();
        g.selecting = true;
        SetCapture(hwnd);
        SetTimer(hwnd, kTimerDrag, 50, nullptr);
        moveCaret(p, (wp & MK_SHIFT) != 0);
        return 0;
    }
    case WM_RBUTTONDOWN: {
        // A right click outside the selection moves the caret there first.
        SetFocus(hwnd);
        u64 p = posFromPoint(GET_X_LPARAM(lp), GET_Y_LPARAM(lp));
        if (!(hasSel() && p >= selMin() && p <= selMax())) moveCaret(p, false);
        return 0;
    }
    case WM_CONTEXTMENU: {
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        if (pt.x == -1 && pt.y == -1) {  // Shift+F10 / menu key: at the caret
            pt = caretPoint();
            ClientToScreen(hwnd, &pt);
        }
        showEditorMenu(pt);
        return 0;
    }
    case WM_LBUTTONDBLCLK: {
        clearExtras();
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
        if (g.boxDrag && (wp & MK_LBUTTON)) {
            u64 line;
            int col;
            const int my = GET_Y_LPARAM(lp);
            if (my < 0 || my >= g.visRows * g.ch) {  // past the top or bottom: scroll along
                scrollRows(my < 0 ? -1 : 1);
                relayout();
            }
            if (boxPoint(GET_X_LPARAM(lp), GET_Y_LPARAM(lp), line, col) && (line != g.boxL1 || col != g.boxC1)) {
                g.boxL1 = line;
                g.boxC1 = col;
                boxRebuild();
                refresh(false);
            }
            return 0;
        }
        if (g.selecting && (wp & MK_LBUTTON)) dragTo(GET_X_LPARAM(lp), GET_Y_LPARAM(lp));
        return 0;
    case WM_SYSKEYUP:
        // Releasing Alt after Alt+click must not open the menu bar.
        if (wp == VK_MENU && g.altClicked) {
            g.altClicked = false;
            return 0;
        }
        return DefWindowProcW(hwnd, msg, wp, lp);
    case WM_SYSKEYDOWN:
        if (wp != VK_MENU) g.altClicked = false;
        return DefWindowProcW(hwnd, msg, wp, lp);
    case WM_LBUTTONUP:
    case WM_CAPTURECHANGED:
        if (g.boxDrag) {
            g.boxDrag = false;
            if (msg == WM_LBUTTONUP) ReleaseCapture();
        }
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
        g.sideTree = CreateWindowExW(0, WC_TREEVIEWW, L"",
                                     WS_CHILD | WS_TABSTOP | TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT |
                                         TVS_SHOWSELALWAYS | TVS_EDITLABELS | TVS_INFOTIP,
                                     0, 0, 0, 0, hwnd, (HMENU)(INT_PTR)1, g.inst, nullptr);
        TreeView_SetExtendedStyle(g.sideTree, TVS_EX_DOUBLEBUFFER, TVS_EX_DOUBLEBUFFER);
        setupFindBar();
        setupDataBar();
        g.sideHex = CreateWindowExW(0, kHexClass, L"", WS_CHILD | WS_VSCROLL | WS_TABSTOP, 0, 0, 0, 0, hwnd, nullptr, g.inst, nullptr);
        g.sideCsv = CreateWindowExW(0, kCsvClass, L"", WS_CHILD | WS_VSCROLL | WS_HSCROLL | WS_TABSTOP | WS_CLIPCHILDREN, 0, 0, 0, 0,
                                    hwnd, nullptr, g.inst, nullptr);
        g.sideSplitter = CreateWindowExW(0, kSplitClass, L"", WS_CHILD, 0, 0, 0, 0, hwnd, nullptr, g.inst, nullptr);
        g.sidePreview = CreateWindowExW(0, kPreviewClass, L"", WS_CHILD | WS_VSCROLL, 0, 0, 0, 0, hwnd, nullptr, g.inst, nullptr);
        SetTimer(hwnd, kTimerDisk, 2000, nullptr);
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
    case WM_INITMENUPOPUP:
        if ((HMENU)wp == g.recentMenu) rebuildRecentMenu();
        updateMenu();
        return 0;
    case WM_SETCURSOR:
        if ((HWND)wp == hwnd && LOWORD(lp) == HTCLIENT) {
            POINT pt;
            GetCursorPos(&pt);
            ScreenToClient(hwnd, &pt);
            if (PtInRect(&g.sideSplitRect, pt) || g.dragSplit == 2) {
                SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
                return TRUE;
            }
        }
        return DefWindowProcW(hwnd, msg, wp, lp);
    case WM_MBUTTONUP: {
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        int i = tabAt(pt, nullptr);
        if (i >= 0) closeTab(i);
        return 0;
    }
    case WM_MOUSEMOVE: {
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        if (g.dragSplit == 2) {  // moving the side divider
            RECT rc;
            GetClientRect(hwnd, &rc);
            if (rc.right > 0) g.sideSplit = std::clamp((double)(rc.right - pt.x - splitW() / 2) / rc.right, 0.1, 0.9);
            layoutFrame();
            UpdateWindow(hwnd);
            return 0;
        }
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
        if (PtInRect(&g.sideSplitRect, pt)) {
            if (msg == WM_LBUTTONDBLCLK) {  // double click: back to the usual width
                g.sideSplit = 0.42;
                layoutFrame();
                saveSettings();
            } else {
                g.dragSplit = 2;
                SetCapture(hwnd);
            }
            return 0;
        }
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
            RECT cr;
            GetClientRect(hwnd, &cr);
            if (msg == WM_LBUTTONDOWN && pt.y >= cr.bottom - g.statusH) {  // status bar: encoding, line endings
                POINT sp = pt;
                ClientToScreen(hwnd, &sp);
                showEncodingMenu(sp);
                return 0;
            }
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
        if (g.dragSplit == 2) {
            g.dragSplit = 0;
            ReleaseCapture();
            saveSettings();
            SetFocus(g.view);
            return 0;
        }
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
        if (g.dragSplit == 2) g.dragSplit = 0;
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
    case WM_APP_BLAME: onBlame((unsigned)wp, (Blame*)lp); return 0;
    case WM_COPYDATA:
        onOpenFromOtherCopy((const COPYDATASTRUCT*)lp);
        return TRUE;
    case WM_NOTIFY:
        if (((NMHDR*)lp)->hwndFrom == g.sideTree) return onTreeNotify((NMHDR*)lp);
        return DefWindowProcW(hwnd, msg, wp, lp);
    case WM_APP_P4ASK: onP4Ask(); return 0;
    case WM_APP_DISKCHECK: checkDisk(IsWindowEnabled(hwnd) != 0); return 0;
    case WM_ACTIVATEAPP:
        if (wp) PostMessageW(hwnd, WM_APP_DISKCHECK, 0, 0);
        return 0;
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
        if (wp == kTimerSide) {
            KillTimer(hwnd, kTimerSide);
            updateSidePane();
            return 0;
        }
        if (wp == kTimerSession) {
            autosaveSession();
            return 0;
        }
        if (wp == kTimerDisk) {
            // Unsaved edits are only asked about while this window is in front.
            checkDisk(GetForegroundWindow() == hwnd && IsWindowEnabled(hwnd));
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

// One window: a second copy started with files (double-clicking a file, "Open
// with", the command line) hands them to the running window as tabs and quits.
const ULONG_PTR kCopyDataOpen = 0x4A4E4F50;  // "JNOP"

bool handToRunningCopy() {
    // The mutex closes the gap between "is one running?" and its window appearing.
    HANDLE mutex = CreateMutexW(nullptr, FALSE, L"Local\\JustNotepad.SingleWindow");
    const bool another = mutex && GetLastError() == ERROR_ALREADY_EXISTS;
    if (!another) return false;  // we are the first: keep the mutex for our lifetime
    HWND w = nullptr;
    for (int i = 0; i < 40 && !(w = FindWindowW(kFrameClass, nullptr)); i++) Sleep(50);
    if (!w) return false;
    // Full paths, so the other window does not depend on our working folder.
    std::wstring list;
    for (int i = 1; i < __argc; i++) {
        if (!__wargv[i][0]) continue;
        wchar_t full[32768];
        DWORD n = GetFullPathNameW(__wargv[i], 32768, full, nullptr);
        list += (n && n < 32768 ? full : __wargv[i]);
        list += L'\n';
    }
    DWORD pid = 0;
    GetWindowThreadProcessId(w, &pid);
    AllowSetForegroundWindow(pid);  // let it come to the front
    COPYDATASTRUCT cd = {kCopyDataOpen, (DWORD)((list.size() + 1) * sizeof(wchar_t)), (void*)list.c_str()};
    DWORD_PTR res = 0;
    return SendMessageTimeoutW(w, WM_COPYDATA, 0, (LPARAM)&cd, SMTO_ABORTIFHUNG, 5000, &res) != 0;
}

// The running window receives files from another copy.
void onOpenFromOtherCopy(const COPYDATASTRUCT* cd) {
    if (!cd || cd->dwData != kCopyDataOpen || !cd->lpData) return;
    std::wstring list((const wchar_t*)cd->lpData, cd->cbData / sizeof(wchar_t));
    size_t a = 0;
    while (a < list.size()) {
        size_t b = list.find(L'\n', a);
        if (b == std::wstring::npos) b = list.size();
        std::wstring f = list.substr(a, b - a);
        while (!f.empty() && f.back() == L'\0') f.pop_back();
        if (!f.empty()) openInTab(f);
        a = b + 1;
    }
    if (IsIconic(g.frame)) ShowWindow(g.frame, SW_RESTORE);
    SetForegroundWindow(g.frame);
    SetFocus(g.view);
}

int WINAPI wWinMain(HINSTANCE inst, HINSTANCE, PWSTR, int show) {
    g.inst = inst;
    if (handToRunningCopy()) return 0;
    loadSettings();
    INITCOMMONCONTROLSEX icc = {sizeof(icc), ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES | ICC_LINK_CLASS | ICC_TREEVIEW_CLASSES};
    InitCommonControlsEx(&icc);
    themeInit();
    gdiplusStart();
    g.pal = makePalette(wantDark());
    themeSetAppDark(g.pal.dark);

    WNDCLASSEXW fc = {};
    fc.cbSize = sizeof(fc);
    fc.lpfnWndProc = frameProc;
    fc.style = CS_DBLCLKS;
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
    dcl.style = CS_DBLCLKS;
    RegisterClassExW(&dcl);
    dcl.style = 0;
    dcl.lpfnWndProc = previewProc;
    dcl.lpszClassName = kPreviewClass;
    RegisterClassExW(&dcl);
    dcl.lpfnWndProc = splitProc;
    dcl.lpszClassName = kSplitClass;
    dcl.style = CS_DBLCLKS;
    RegisterClassExW(&dcl);
    dcl.lpfnWndProc = hexProc;
    dcl.lpszClassName = kHexClass;
    dcl.hCursor = LoadCursorW(nullptr, IDC_IBEAM);
    RegisterClassExW(&dcl);
    dcl.lpfnWndProc = csvProc;
    dcl.lpszClassName = kCsvClass;
    dcl.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    RegisterClassExW(&dcl);
    dcl.style = 0;
    dcl.lpfnWndProc = barProc;
    dcl.lpszClassName = kBarClass;
    dcl.style = CS_DBLCLKS;
    RegisterClassExW(&dcl);

    g.menu = buildMenu();
    themeOwnerDrawMenu(g.menu);
    markEncodingRadios(g.menu);
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
    g_restoring = true;
    restoreSession();
    g_restoring = false;
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
        // Text boxes (find bar, data search, a tree value being edited) get
        // their own keys: Ctrl+V, Ctrl+Z, Ctrl+A ... must not reach the editor.
        // Only find-related shortcuts still work from them.
        if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN || msg.message == WM_CHAR) {
            wchar_t cls[8] = L"";
            HWND f = GetFocus();
            if (f && GetClassNameW(f, cls, 8) && lstrcmpiW(cls, L"Edit") == 0) {
                const bool ctrl = GetKeyState(VK_CONTROL) < 0;
                const bool findKey = msg.message == WM_KEYDOWN &&
                                     (msg.wParam == VK_F3 || (ctrl && (msg.wParam == 'F' || msg.wParam == 'H')));
                if (!(findKey && TranslateAcceleratorW(g.frame, g.accel, &msg))) {
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
                continue;
            }
        }
        if (TranslateAcceleratorW(g.frame, g.accel, &msg)) continue;
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    gdiplusStop();
    return (int)msg.wParam;
}
