// Terminal panel: see terminal.h.
//
// The shell runs on a pseudo console (CreatePseudoConsole, Windows 10 1809+).
// Its output is a stream of UTF-8 text with VT escape sequences, which a small
// emulator here turns into a grid of cells; keys go back as VT sequences.
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include "terminal.h"

#include <windowsx.h>
#include <shellapi.h>

#include <algorithm>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <cwchar>
#include <atomic>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

namespace {

// ---- pseudo console API, looked up at run time (older Windows lacks it)

typedef void* HPCON_;
typedef HRESULT(WINAPI* CreatePC)(COORD, HANDLE, HANDLE, DWORD, HPCON_*);
typedef HRESULT(WINAPI* ResizePC)(HPCON_, COORD);
typedef void(WINAPI* ClosePC)(HPCON_);
const DWORD_PTR kAttrPseudoConsole = 0x00020016;  // PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE

CreatePC pCreate = nullptr;
ResizePC pResize = nullptr;
ClosePC pClose = nullptr;

bool loadApi() {
    if (pCreate) return true;
    HMODULE k = GetModuleHandleW(L"kernel32.dll");
    pCreate = (CreatePC)GetProcAddress(k, "CreatePseudoConsole");
    pResize = (ResizePC)GetProcAddress(k, "ResizePseudoConsole");
    pClose = (ClosePC)GetProcAddress(k, "ClosePseudoConsole");
    return pCreate && pResize && pClose;
}

const UINT WM_TERM_OUTPUT = WM_APP + 41, WM_TERM_EXITED = WM_APP + 42;
const wchar_t* const kTermClass = L"JustNotepadTerminal";

// ---- the screen

// A colour: default, one of 256 indexed, or 24-bit.
const uint32_t kDefault = 0xFF000000u;
inline uint32_t indexed(int i) { return 0x01000000u | (uint32_t)i; }
inline uint32_t rgbColor(int r, int g, int b) { return 0x02000000u | RGB(r, g, b); }

enum { F_BOLD = 1, F_UNDER = 2, F_INVERSE = 4, F_DIM = 8 };

struct Cell {
    wchar_t ch = L' ';
    uint32_t fg = kDefault, bg = kDefault;
    uint8_t fl = 0;
};

struct Line {
    std::vector<Cell> c;
    bool wrapped = false;  // continues on the next line (no line break when copied)
};

struct Screen {
    std::vector<Line> lines;
    int cx = 0, cy = 0;
    bool wrapPending = false;
    int top = 0, bot = 0;  // scrolling region, inclusive
    int savedX = 0, savedY = 0;
};

struct Term {
    HWND wnd = nullptr, parent = nullptr;
    HFONT font = nullptr, bold = nullptr, uiFont = nullptr;
    int cw = 8, ch = 16, headerH = 22;
    Palette pal{};
    int cols = 80, rows = 24;

    Screen mainScr, altScr;
    Screen* scr = &mainScr;
    bool alt = false;
    std::vector<Line> scrollback;
    int viewOff = 0;  // lines scrolled back from the bottom
    Cell pen;         // current colours and attributes (ch unused)

    // modes
    bool appCursor = false, cursorVisible = true, autowrap = true, bracketedPaste = false, focusEvents = false;

    // parser
    int state = 0;  // 0 ground, 1 ESC, 2 CSI, 3 OSC, 4 OSC got ESC, 5 skip one char (charset)
    std::string params;
    char csiPrivate = 0;
    std::string osc;
    uint32_t u8cp = 0;
    int u8need = 0;

    // selection, in absolute line numbers (scrollback first, then the screen)
    bool selecting = false, hasSel = false;
    long long selA = 0, selB = 0;  // line
    int selAc = 0, selBc = 0;       // column

    // header: drag to resize, × to close
    bool dragging = false;
    int dragStartY = 0, dragStartH = 0;
    RECT closeRc{};
    bool closeHot = false;

    // the process
    HPCON_ hpc = nullptr;
    HANDLE inWrite = nullptr, outRead = nullptr, process = nullptr;
    bool running = false;
    std::wstring title, shellName;
    std::mutex mu;
    std::string pending;  // output read but not yet parsed
    unsigned gen = 0;     // which shell a reply belongs to
    std::wstring folder;

    bool focused = false, cursorOn = true;
    bool eatEsc = false;  // Esc closed the panel: its WM_CHAR is not for the shell
};

Term T;

void resetScreen(Screen& s) {
    s.lines.assign((size_t)T.rows, Line());
    for (auto& l : s.lines) l.c.assign((size_t)T.cols, Cell());
    s.cx = s.cy = 0;
    s.wrapPending = false;
    s.top = 0;
    s.bot = T.rows - 1;
}

Cell blank() {
    Cell c;
    c.bg = T.pen.bg;  // erased cells keep the background colour, as consoles do
    return c;
}

void clearSel() { T.hasSel = false; }

// ---- sending to the shell

void send(const std::string& s) {
    if (!T.running || !T.inWrite || s.empty()) return;
    DWORD put = 0;
    WriteFile(T.inWrite, s.data(), (DWORD)s.size(), &put, nullptr);
}

std::string utf8(const std::wstring& w) {
    if (w.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    std::string s((size_t)n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), &s[0], n, nullptr, nullptr);
    return s;
}

// ---- screen operations

void scrollUp(int n) {
    Screen& s = *T.scr;
    for (int k = 0; k < n; k++) {
        if (s.top == 0 && !T.alt) {
            T.scrollback.push_back(std::move(s.lines[0]));
            if (T.scrollback.size() > 5000) {
                T.scrollback.erase(T.scrollback.begin(), T.scrollback.begin() + 500);
                clearSel();
            }
            if (T.viewOff > 0) T.viewOff = std::min<int>(T.viewOff + 1, (int)T.scrollback.size());
        }
        s.lines.erase(s.lines.begin() + s.top);
        Line l;
        l.c.assign((size_t)T.cols, blank());
        s.lines.insert(s.lines.begin() + s.bot, std::move(l));
    }
}

void scrollDown(int n) {
    Screen& s = *T.scr;
    for (int k = 0; k < n; k++) {
        s.lines.erase(s.lines.begin() + s.bot);
        Line l;
        l.c.assign((size_t)T.cols, blank());
        s.lines.insert(s.lines.begin() + s.top, std::move(l));
    }
}

void lineFeed() {
    Screen& s = *T.scr;
    s.wrapPending = false;
    if (s.cy == s.bot) scrollUp(1);
    else if (s.cy < T.rows - 1) s.cy++;
}

void putChar(wchar_t ch) {
    Screen& s = *T.scr;
    if (s.wrapPending) {
        if (T.autowrap) {
            s.lines[(size_t)s.cy].wrapped = true;
            s.cx = 0;
            lineFeed();
        }
        s.wrapPending = false;
    }
    Cell c = T.pen;
    c.ch = ch;
    s.lines[(size_t)s.cy].c[(size_t)s.cx] = c;
    if (s.cx == T.cols - 1) s.wrapPending = true;
    else s.cx++;
}

void eraseCells(Line& l, int from, int to) {
    from = std::clamp(from, 0, T.cols);
    to = std::clamp(to, 0, T.cols);
    for (int i = from; i < to; i++) l.c[(size_t)i] = blank();
}

int param(const std::vector<int>& p, size_t i, int def) { return i < p.size() && p[i] > 0 ? p[i] : def; }

void sgr(const std::vector<int>& p) {
    if (p.empty()) {
        T.pen = Cell();
        return;
    }
    for (size_t i = 0; i < p.size(); i++) {
        const int v = p[i] < 0 ? 0 : p[i];
        if (v == 0) T.pen = Cell();
        else if (v == 1) T.pen.fl |= F_BOLD;
        else if (v == 2) T.pen.fl |= F_DIM;
        else if (v == 4) T.pen.fl |= F_UNDER;
        else if (v == 7) T.pen.fl |= F_INVERSE;
        else if (v == 22) T.pen.fl &= (uint8_t)~(F_BOLD | F_DIM);
        else if (v == 24) T.pen.fl &= (uint8_t)~F_UNDER;
        else if (v == 27) T.pen.fl &= (uint8_t)~F_INVERSE;
        else if (v >= 30 && v <= 37) T.pen.fg = indexed(v - 30);
        else if (v == 39) T.pen.fg = kDefault;
        else if (v >= 40 && v <= 47) T.pen.bg = indexed(v - 40);
        else if (v == 49) T.pen.bg = kDefault;
        else if (v >= 90 && v <= 97) T.pen.fg = indexed(v - 90 + 8);
        else if (v >= 100 && v <= 107) T.pen.bg = indexed(v - 100 + 8);
        else if (v == 38 || v == 48) {
            uint32_t col = kDefault;
            if (i + 2 < p.size() && p[i + 1] == 5) {
                col = indexed(std::clamp(p[i + 2], 0, 255));
                i += 2;
            } else if (i + 4 < p.size() && p[i + 1] == 2) {
                col = rgbColor(std::clamp(p[i + 2], 0, 255), std::clamp(p[i + 3], 0, 255), std::clamp(p[i + 4], 0, 255));
                i += 4;
            } else {
                break;
            }
            (v == 38 ? T.pen.fg : T.pen.bg) = col;
        }
    }
}

void setAlt(bool on) {
    if (on == T.alt) return;
    T.alt = on;
    T.scr = on ? &T.altScr : &T.mainScr;
    if (on) {
        resetScreen(T.altScr);
        T.viewOff = 0;
    }
    clearSel();
}

void csi(char fin) {
    Screen& s = *T.scr;
    std::vector<int> p;
    {
        int cur = -1;
        for (char ch : T.params) {
            if (ch >= '0' && ch <= '9') cur = (cur < 0 ? 0 : cur) * 10 + (ch - '0');
            else if (ch == ';' || ch == ':') {
                p.push_back(cur);
                cur = -1;
            }
        }
        if (cur >= 0 || !T.params.empty()) p.push_back(cur);
    }
    if (T.csiPrivate == '?') {
        if (fin == 'h' || fin == 'l') {
            const bool on = fin == 'h';
            for (int m : p) {
                if (m == 1) T.appCursor = on;
                else if (m == 7) T.autowrap = on;
                else if (m == 25) T.cursorVisible = on;
                else if (m == 47 || m == 1047 || m == 1049) setAlt(on);
                else if (m == 2004) T.bracketedPaste = on;
                else if (m == 1004) T.focusEvents = on;
            }
        }
        return;
    }
    if (T.csiPrivate) return;  // '>' and '=' requests: nothing to answer
    const int n1 = param(p, 0, 1);
    auto clampCursor = [&]() {
        s.cx = std::clamp(s.cx, 0, T.cols - 1);
        s.cy = std::clamp(s.cy, 0, T.rows - 1);
        s.wrapPending = false;
    };
    switch (fin) {
    case 'A': s.cy = std::max(s.cy - n1, s.cy >= s.top ? s.top : 0); clampCursor(); break;
    case 'B': s.cy = std::min(s.cy + n1, s.cy <= s.bot ? s.bot : T.rows - 1); clampCursor(); break;
    case 'C': s.cx += n1; clampCursor(); break;
    case 'D': s.cx -= n1; clampCursor(); break;
    case 'E': s.cy += n1; s.cx = 0; clampCursor(); break;
    case 'F': s.cy -= n1; s.cx = 0; clampCursor(); break;
    case 'G': case '`': s.cx = n1 - 1; clampCursor(); break;
    case 'd': s.cy = n1 - 1; clampCursor(); break;
    case 'H': case 'f': s.cy = param(p, 0, 1) - 1; s.cx = param(p, 1, 1) - 1; clampCursor(); break;
    case 'J': {
        const int m = p.empty() || p[0] < 0 ? 0 : p[0];
        if (m == 0) {
            eraseCells(s.lines[(size_t)s.cy], s.cx, T.cols);
            for (int y = s.cy + 1; y < T.rows; y++) eraseCells(s.lines[(size_t)y], 0, T.cols);
        } else if (m == 1) {
            eraseCells(s.lines[(size_t)s.cy], 0, s.cx + 1);
            for (int y = 0; y < s.cy; y++) eraseCells(s.lines[(size_t)y], 0, T.cols);
        } else if (m == 2) {
            for (auto& l : s.lines) {
                eraseCells(l, 0, T.cols);
                l.wrapped = false;
            }
        } else if (m == 3) {
            T.scrollback.clear();
            T.viewOff = 0;
            clearSel();
        }
        break;
    }
    case 'K': {
        const int m = p.empty() || p[0] < 0 ? 0 : p[0];
        Line& l = s.lines[(size_t)s.cy];
        if (m == 0) eraseCells(l, s.cx, T.cols);
        else if (m == 1) eraseCells(l, 0, s.cx + 1);
        else eraseCells(l, 0, T.cols);
        if (m != 1) l.wrapped = false;
        break;
    }
    case 'X': eraseCells(s.lines[(size_t)s.cy], s.cx, s.cx + n1); break;
    case 'P': {
        auto& c = s.lines[(size_t)s.cy].c;
        const int k = std::min(n1, T.cols - s.cx);
        c.erase(c.begin() + s.cx, c.begin() + s.cx + k);
        c.insert(c.end(), (size_t)k, blank());
        break;
    }
    case '@': {
        auto& c = s.lines[(size_t)s.cy].c;
        const int k = std::min(n1, T.cols - s.cx);
        c.insert(c.begin() + s.cx, (size_t)k, blank());
        c.resize((size_t)T.cols);
        break;
    }
    case 'L':
    case 'M':
        if (s.cy >= s.top && s.cy <= s.bot) {
            const int saveTop = s.top;
            s.top = s.cy;
            const bool wasAlt = T.alt;
            T.alt = true;  // lines pushed out here never go to the scrollback
            if (fin == 'L') scrollDown(std::min(n1, s.bot - s.cy + 1));
            else scrollUp(std::min(n1, s.bot - s.cy + 1));
            T.alt = wasAlt;
            s.top = saveTop;
        }
        break;
    case 'S': scrollUp(n1); break;
    case 'T': scrollDown(n1); break;
    case 'r': {
        int t = param(p, 0, 1) - 1, b = param(p, 1, T.rows) - 1;
        t = std::clamp(t, 0, T.rows - 1);
        b = std::clamp(b, 0, T.rows - 1);
        if (t < b) {
            s.top = t;
            s.bot = b;
        } else {
            s.top = 0;
            s.bot = T.rows - 1;
        }
        s.cx = s.cy = 0;
        s.wrapPending = false;
        break;
    }
    case 'm': sgr(p); break;
    case 's': s.savedX = s.cx; s.savedY = s.cy; break;
    case 'u': s.cx = s.savedX; s.cy = s.savedY; clampCursor(); break;
    case 'n':
        if (param(p, 0, 0) == 6) {
            char b[32];
            snprintf(b, sizeof(b), "\x1b[%d;%dR", s.cy + 1, s.cx + 1);
            send(b);
        } else if (param(p, 0, 0) == 5) {
            send("\x1b[0n");
        }
        break;
    case 'c': send("\x1b[?1;0c"); break;
    }
}

void escFinal(char ch) {
    Screen& s = *T.scr;
    switch (ch) {
    case '7': s.savedX = s.cx; s.savedY = s.cy; break;
    case '8': s.cx = s.savedX; s.cy = s.savedY; s.wrapPending = false; break;
    case 'M':  // reverse index
        if (s.cy == s.top) scrollDown(1);
        else if (s.cy > 0) s.cy--;
        break;
    case 'D': lineFeed(); break;
    case 'E': s.cx = 0; lineFeed(); break;
    case 'c':
        T.pen = Cell();
        resetScreen(*T.scr);
        break;
    }
}

void oscDone() {
    // "0;title" or "2;title": the window title, shown in the header.
    size_t semi = T.osc.find(';');
    if (semi != std::string::npos) {
        std::string code = T.osc.substr(0, semi);
        if (code == "0" || code == "2") {
            std::string t = T.osc.substr(semi + 1);
            int n = MultiByteToWideChar(CP_UTF8, 0, t.data(), (int)t.size(), nullptr, 0);
            std::wstring w((size_t)n, L'\0');
            if (n) MultiByteToWideChar(CP_UTF8, 0, t.data(), (int)t.size(), &w[0], n);
            T.title = w;
        }
    }
    T.osc.clear();
}

void feedChar(uint32_t cp) {
    switch (T.state) {
    case 0:
        if (cp == 0x1b) {
            T.state = 1;
        } else if (cp == '\r') {
            T.scr->cx = 0;
            T.scr->wrapPending = false;
        } else if (cp == '\n' || cp == 0x0b || cp == 0x0c) {
            lineFeed();
        } else if (cp == '\b') {
            if (T.scr->cx > 0) T.scr->cx--;
            T.scr->wrapPending = false;
        } else if (cp == '\t') {
            T.scr->cx = std::min(T.cols - 1, (T.scr->cx / 8 + 1) * 8);
        } else if (cp >= 0x20 && cp != 0x7f) {
            putChar(cp < 0x10000 ? (wchar_t)cp : L'\xFFFD');
        }
        break;
    case 1:
        if (cp == '[') {
            T.state = 2;
            T.params.clear();
            T.csiPrivate = 0;
        } else if (cp == ']') {
            T.state = 3;
            T.osc.clear();
        } else if (cp == '(' || cp == ')' || cp == '*' || cp == '+') {
            T.state = 5;
        } else {
            escFinal((char)cp);
            T.state = 0;
        }
        break;
    case 2:
        if (cp >= 0x40 && cp <= 0x7e) {
            csi((char)cp);
            T.state = 0;
        } else if ((cp == '?' || cp == '>' || cp == '=' || cp == '<') && T.params.empty()) {
            T.csiPrivate = (char)cp;
        } else if (cp >= 0x20 && cp < 0x40) {
            if (T.params.size() < 256) T.params += (char)cp;
        } else {
            T.state = 0;  // malformed
        }
        break;
    case 3:
        if (cp == 0x07) {
            oscDone();
            T.state = 0;
        } else if (cp == 0x1b) {
            T.state = 4;
        } else if (T.osc.size() < 4096) {
            // keep it as UTF-8
            char b[4];
            int n = 0;
            if (cp < 0x80) b[n++] = (char)cp;
            else if (cp < 0x800) {
                b[n++] = (char)(0xC0 | (cp >> 6));
                b[n++] = (char)(0x80 | (cp & 0x3F));
            } else {
                b[n++] = (char)(0xE0 | (cp >> 12));
                b[n++] = (char)(0x80 | ((cp >> 6) & 0x3F));
                b[n++] = (char)(0x80 | (cp & 0x3F));
            }
            T.osc.append(b, (size_t)n);
        }
        break;
    case 4:  // ESC \ ends the OSC
        oscDone();
        T.state = 0;
        if (cp != '\\') feedChar(cp);
        break;
    case 5: T.state = 0; break;
    }
}

void feed(const std::string& bytes) {
    for (unsigned char b : bytes) {
        if (T.u8need) {
            if ((b & 0xC0) == 0x80) {
                T.u8cp = (T.u8cp << 6) | (b & 0x3F);
                if (--T.u8need == 0) feedChar(T.u8cp);
                continue;
            }
            T.u8need = 0;
            feedChar(0xFFFD);
        }
        if (b < 0x80) feedChar(b);
        else if ((b & 0xE0) == 0xC0) {
            T.u8cp = b & 0x1F;
            T.u8need = 1;
        } else if ((b & 0xF0) == 0xE0) {
            T.u8cp = b & 0x0F;
            T.u8need = 2;
        } else if ((b & 0xF8) == 0xF0) {
            T.u8cp = b & 0x07;
            T.u8need = 3;
        } else {
            feedChar(0xFFFD);
        }
    }
}

// ---- size

void computeSize(int& cols, int& rows) {
    RECT rc;
    GetClientRect(T.wnd, &rc);
    const int pad = T.cw / 2;
    cols = std::max(10, (int)(rc.right - 2 * pad) / std::max(1, T.cw));
    rows = std::max(2, (int)(rc.bottom - T.headerH - pad) / std::max(1, T.ch));
}

void resizeScreen(Screen& s, int cols, int rows) {
    // Keep the bottom of the screen (where the prompt is) when it gets shorter.
    if ((int)s.lines.size() > rows) {
        const int drop = std::min((int)s.lines.size() - rows, std::max(0, s.cy - rows + 1));
        for (int i = 0; i < drop; i++) {
            if (&s == &T.mainScr) T.scrollback.push_back(std::move(s.lines[(size_t)i]));
        }
        s.lines.erase(s.lines.begin(), s.lines.begin() + drop);
        s.cy -= drop;
        s.lines.resize((size_t)rows);
    }
    while ((int)s.lines.size() < rows) s.lines.push_back(Line());
    for (auto& l : s.lines) l.c.resize((size_t)cols, Cell());
    s.cx = std::clamp(s.cx, 0, cols - 1);
    s.cy = std::clamp(s.cy, 0, rows - 1);
    s.top = 0;
    s.bot = rows - 1;
    s.wrapPending = false;
}

void applySize() {
    int cols, rows;
    computeSize(cols, rows);
    if (cols == T.cols && rows == T.rows) return;
    resizeScreen(T.mainScr, cols, rows);
    resizeScreen(T.altScr, cols, rows);
    T.cols = cols;
    T.rows = rows;
    T.viewOff = std::min<int>(T.viewOff, (int)T.scrollback.size());
    clearSel();
    if (T.running && T.hpc) pResize(T.hpc, COORD{(SHORT)cols, (SHORT)rows});
}

// ---- the shell process

std::wstring findShell() {
    wchar_t buf[MAX_PATH];
    if (SearchPathW(nullptr, L"pwsh.exe", nullptr, MAX_PATH, buf, nullptr)) {
        T.shellName = L"PowerShell";
        return std::wstring(L"\"") + buf + L"\" -NoLogo";
    }
    wchar_t sys[MAX_PATH];
    GetSystemDirectoryW(sys, MAX_PATH);
    T.shellName = L"Windows PowerShell";
    return std::wstring(L"\"") + sys + L"\\WindowsPowerShell\\v1.0\\powershell.exe\" -NoLogo";
}

void stopShell() {
    if (!T.running && !T.hpc) return;
    T.running = false;
    T.gen++;  // replies from the old shell are ignored from now on
    if (T.inWrite) CloseHandle(T.inWrite);
    T.inWrite = nullptr;
    // Closing the pseudo console ends the shell and breaks the output pipe,
    // which ends the reader thread. It can wait for that pipe to be drained,
    // so it runs on a thread of its own.
    HPCON_ h = T.hpc;
    T.hpc = nullptr;
    if (h) std::thread([h]() { pClose(h); }).detach();
    if (T.process) {
        TerminateProcess(T.process, 0);
        CloseHandle(T.process);
        T.process = nullptr;
    }
}

bool startShell() {
    if (!loadApi()) {
        std::string msg = "\x1b[31mThe terminal needs Windows 10 version 1809 or later.\x1b[0m\r\n";
        feed(msg);
        return false;
    }
    stopShell();
    applySize();
    HANDLE inRead = nullptr, inWrite = nullptr, outRead = nullptr, outWrite = nullptr;
    if (!CreatePipe(&inRead, &inWrite, nullptr, 0) || !CreatePipe(&outRead, &outWrite, nullptr, 0)) return false;
    HPCON_ hpc = nullptr;
    if (FAILED(pCreate(COORD{(SHORT)T.cols, (SHORT)T.rows}, inRead, outWrite, 0, &hpc))) {
        CloseHandle(inRead);
        CloseHandle(inWrite);
        CloseHandle(outRead);
        CloseHandle(outWrite);
        return false;
    }
    SIZE_T sz = 0;
    InitializeProcThreadAttributeList(nullptr, 1, 0, &sz);
    std::vector<unsigned char> attrBuf(sz);
    auto* attrs = (LPPROC_THREAD_ATTRIBUTE_LIST)attrBuf.data();
    InitializeProcThreadAttributeList(attrs, 1, 0, &sz);
    UpdateProcThreadAttribute(attrs, 0, kAttrPseudoConsole, hpc, sizeof(hpc), nullptr, nullptr);
    STARTUPINFOEXW si = {};
    si.StartupInfo.cb = sizeof(si);
    si.lpAttributeList = attrs;
    PROCESS_INFORMATION pi = {};
    std::wstring cmd = findShell();
    const wchar_t* dir = !T.folder.empty() && GetFileAttributesW(T.folder.c_str()) != INVALID_FILE_ATTRIBUTES ? T.folder.c_str() : nullptr;
    BOOL ok = CreateProcessW(nullptr, &cmd[0], nullptr, nullptr, FALSE, EXTENDED_STARTUPINFO_PRESENT, nullptr, dir,
                             &si.StartupInfo, &pi);
    DeleteProcThreadAttributeList(attrs);
    // The pseudo console holds its own copies of these.
    CloseHandle(inRead);
    CloseHandle(outWrite);
    if (!ok) {
        pClose(hpc);
        CloseHandle(inWrite);
        CloseHandle(outRead);
        feed("\x1b[31mCould not start PowerShell.\x1b[0m\r\n");
        return false;
    }
    CloseHandle(pi.hThread);
    T.hpc = hpc;
    T.inWrite = inWrite;
    T.process = pi.hProcess;
    T.running = true;
    const unsigned gen = ++T.gen;
    HWND wnd = T.wnd;

    // Reader: output goes into `pending`, the window parses it.
    std::thread([outRead, wnd, gen]() {
        char buf[16384];
        for (;;) {
            DWORD got = 0;
            if (!ReadFile(outRead, buf, sizeof(buf), &got, nullptr) || got == 0) break;
            bool first;
            {
                std::lock_guard<std::mutex> lk(T.mu);
                first = T.pending.empty();
                T.pending.append(buf, got);
            }
            if (first) PostMessageW(wnd, WM_TERM_OUTPUT, gen, 0);
        }
        CloseHandle(outRead);
    }).detach();
    // Watcher: tells the window when the shell exits.
    HANDLE proc = nullptr;
    DuplicateHandle(GetCurrentProcess(), pi.hProcess, GetCurrentProcess(), &proc, SYNCHRONIZE, FALSE, 0);
    if (proc)
        std::thread([proc, wnd, gen]() {
            WaitForSingleObject(proc, INFINITE);
            CloseHandle(proc);
            PostMessageW(wnd, WM_TERM_EXITED, gen, 0);
        }).detach();
    return true;
}

// ---- colours

COLORREF ansi(int i) {
    static const COLORREF dark[16] = {RGB(12, 12, 12),    RGB(197, 15, 31),  RGB(19, 161, 14),   RGB(193, 156, 0),
                                      RGB(0, 55, 218),    RGB(136, 23, 152), RGB(58, 150, 221),  RGB(204, 204, 204),
                                      RGB(118, 118, 118), RGB(231, 72, 86),  RGB(22, 198, 12),   RGB(249, 241, 165),
                                      RGB(59, 120, 255),  RGB(180, 0, 158),  RGB(97, 214, 214),  RGB(242, 242, 242)};
    static const COLORREF light[16] = {RGB(56, 58, 66),   RGB(228, 86, 73),  RGB(80, 161, 79),   RGB(193, 131, 1),
                                       RGB(1, 132, 188),  RGB(166, 38, 164), RGB(9, 151, 179),   RGB(160, 161, 167),
                                       RGB(79, 82, 93),   RGB(202, 18, 67),  RGB(56, 130, 55),   RGB(152, 104, 1),
                                       RGB(64, 120, 242), RGB(166, 38, 164), RGB(1, 132, 188),   RGB(9, 10, 11)};
    if (i < 16) return (T.pal.dark ? dark : light)[i];
    if (i < 232) {
        i -= 16;
        static const int lv[6] = {0, 95, 135, 175, 215, 255};
        return RGB(lv[(i / 36) % 6], lv[(i / 6) % 6], lv[i % 6]);
    }
    const int v = 8 + (i - 232) * 10;
    return RGB(v, v, v);
}

COLORREF resolve(uint32_t c, bool fg, bool bold) {
    if (c == kDefault) return fg ? T.pal.fg : T.pal.bg;
    if ((c >> 24) == 1) {
        int i = (int)(c & 0xFF);
        if (fg && bold && i < 8) i += 8;  // bold shows as the bright colour
        return ansi(i);
    }
    return (COLORREF)(c & 0xFFFFFF);
}

COLORREF mix(COLORREF a, COLORREF b, int pctA) {
    return RGB((GetRValue(a) * pctA + GetRValue(b) * (100 - pctA)) / 100,
               (GetGValue(a) * pctA + GetGValue(b) * (100 - pctA)) / 100,
               (GetBValue(a) * pctA + GetBValue(b) * (100 - pctA)) / 100);
}

// ---- drawing

const Line* lineAt(long long abs) {
    const long long sb = (long long)T.scrollback.size();
    if (abs < 0) return nullptr;
    if (abs < sb && !T.alt) return &T.scrollback[(size_t)abs];
    const long long r = T.alt ? abs - sb : abs - sb;
    if (r >= 0 && r < (long long)T.scr->lines.size()) return &T.scr->lines[(size_t)r];
    return nullptr;
}

long long topAbs() { return (long long)T.scrollback.size() - (T.alt ? 0 : T.viewOff); }

bool inSel(long long line, int col) {
    if (!T.hasSel) return false;
    long long a = T.selA, b = T.selB;
    int ac = T.selAc, bc = T.selBc;
    if (a > b || (a == b && ac > bc)) {
        std::swap(a, b);
        std::swap(ac, bc);
    }
    if (line < a || line > b) return false;
    if (line == a && col < ac) return false;
    if (line == b && col >= bc) return false;
    return true;
}

void paint() {
    PAINTSTRUCT ps;
    HDC wdc = BeginPaint(T.wnd, &ps);
    RECT rc;
    GetClientRect(T.wnd, &rc);
    HDC dc = CreateCompatibleDC(wdc);
    HBITMAP bmp = CreateCompatibleBitmap(wdc, std::max(1, (int)rc.right), std::max(1, (int)rc.bottom));
    HGDIOBJ ob = SelectObject(dc, bmp);

    auto fill = [&](int l, int t, int r, int b, COLORREF c) {
        RECT x = {l, t, r, b};
        SetBkColor(dc, c);
        ExtTextOutW(dc, 0, 0, ETO_OPAQUE, &x, L"", 0, nullptr);
    };
    fill(0, 0, rc.right, rc.bottom, T.pal.bg);

    // No header: a thin edge on top, which drags to resize.
    const int hh = T.headerH;
    fill(0, 0, rc.right, 1, T.pal.border);
    HGDIOBJ of = SelectObject(dc, T.font);
    T.closeRc = {};

    // The grid.
    SetBkMode(dc, OPAQUE);
    const int x0 = T.cw / 2, y0 = hh + T.cw / 4;
    std::vector<INT> dx((size_t)T.cols, T.cw);
    const long long top = topAbs();
    const Screen& s = *T.scr;
    const bool cursorHere = T.cursorVisible && (T.alt || T.viewOff == 0);
    for (int r = 0; r < T.rows; r++) {
        const Line* l = lineAt(top + r);
        if (!l) continue;
        const int y = y0 + r * T.ch;
        int c = 0;
        const int n = std::min<int>((int)l->c.size(), T.cols);
        while (c < n) {
            const Cell& first = l->c[(size_t)c];
            const bool sel = inSel(top + r, c);
            int e = c + 1;
            while (e < n) {
                const Cell& x = l->c[(size_t)e];
                if (x.fg != first.fg || x.bg != first.bg || x.fl != first.fl || inSel(top + r, e) != sel) break;
                e++;
            }
            COLORREF fg = resolve(first.fg, true, (first.fl & F_BOLD) != 0), bg = resolve(first.bg, false, false);
            if (first.fl & F_INVERSE) std::swap(fg, bg);
            if (first.fl & F_DIM) fg = mix(fg, bg, 60);
            if (sel) bg = T.pal.selection, fg = T.pal.fg;
            std::wstring text;
            for (int i = c; i < e; i++) text += l->c[(size_t)i].ch ? l->c[(size_t)i].ch : L' ';
            SelectObject(dc, (first.fl & F_BOLD) && T.bold ? T.bold : T.font);
            SetTextColor(dc, fg);
            SetBkColor(dc, bg);
            RECT cr = {x0 + c * T.cw, y, x0 + e * T.cw, y + T.ch};
            ExtTextOutW(dc, cr.left, y, ETO_OPAQUE | ETO_CLIPPED, &cr, text.c_str(), (UINT)text.size(), dx.data());
            if (first.fl & F_UNDER) fill(cr.left, y + T.ch - 1, cr.right, y + T.ch, fg);
            c = e;
        }
        // The cursor: a block when focused, an outline when not.
        if (cursorHere && top + r - (long long)T.scrollback.size() == s.cy && !T.alt ? true : (T.alt && cursorHere && r == s.cy)) {
            const int cx = x0 + std::min(s.cx, T.cols - 1) * T.cw;
            const COLORREF cc = T.pal.fg;
            if (T.focused && T.cursorOn) {
                const Cell& under = l->c[(size_t)std::min(s.cx, n - 1)];
                RECT cr = {cx, y, cx + T.cw, y + T.ch};
                SetBkColor(dc, cc);
                SetTextColor(dc, T.pal.bg);
                SelectObject(dc, T.font);
                wchar_t ch = under.ch ? under.ch : L' ';
                ExtTextOutW(dc, cx, y, ETO_OPAQUE | ETO_CLIPPED, &cr, &ch, 1, nullptr);
            } else if (!T.focused) {
                fill(cx, y, cx + T.cw, y + 1, cc);
                fill(cx, y + T.ch - 1, cx + T.cw, y + T.ch, cc);
                fill(cx, y, cx + 1, y + T.ch, cc);
                fill(cx + T.cw - 1, y, cx + T.cw, y + T.ch, cc);
            }
        }
    }
    SelectObject(dc, of);
    BitBlt(wdc, 0, 0, rc.right, rc.bottom, dc, 0, 0, SRCCOPY);
    SelectObject(dc, ob);
    DeleteObject(bmp);
    DeleteDC(dc);
    EndPaint(T.wnd, &ps);
}

void invalidate() { InvalidateRect(T.wnd, nullptr, FALSE); }

// ---- selection and clipboard

bool cellAt(int x, int y, long long& line, int& col) {
    const int x0 = T.cw / 2, y0 = T.headerH + T.cw / 4;
    int r = (y - y0) / std::max(1, T.ch);
    if (y < y0) r = -1;
    r = std::clamp(r, 0, T.rows - 1);
    col = std::clamp((x - x0 + T.cw / 2) / std::max(1, T.cw), 0, T.cols);
    line = topAbs() + r;
    return true;
}

std::wstring selectedText() {
    long long a = T.selA, b = T.selB;
    int ac = T.selAc, bc = T.selBc;
    if (a > b || (a == b && ac > bc)) {
        std::swap(a, b);
        std::swap(ac, bc);
    }
    std::wstring out;
    for (long long ln = a; ln <= b; ln++) {
        const Line* l = lineAt(ln);
        if (!l) continue;
        const int from = ln == a ? ac : 0, to = ln == b ? bc : (int)l->c.size();
        std::wstring part;
        for (int i = from; i < to && i < (int)l->c.size(); i++) part += l->c[(size_t)i].ch ? l->c[(size_t)i].ch : L' ';
        if (!(l->wrapped && ln != b)) {
            while (!part.empty() && part.back() == L' ') part.pop_back();
        }
        out += part;
        if (ln != b && !l->wrapped) out += L"\r\n";
    }
    return out;
}

void copySelection() {
    std::wstring w = selectedText();
    if (w.empty() || !OpenClipboard(T.wnd)) return;
    EmptyClipboard();
    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE, (w.size() + 1) * sizeof(wchar_t));
    if (h) {
        memcpy(GlobalLock(h), w.c_str(), (w.size() + 1) * sizeof(wchar_t));
        GlobalUnlock(h);
        if (!SetClipboardData(CF_UNICODETEXT, h)) GlobalFree(h);
    }
    CloseClipboard();
}

void paste() {
    if (!OpenClipboard(T.wnd)) return;
    std::wstring w;
    if (HANDLE h = GetClipboardData(CF_UNICODETEXT)) {
        if (const wchar_t* p = (const wchar_t*)GlobalLock(h)) {
            w = p;
            GlobalUnlock(h);
        }
    }
    CloseClipboard();
    if (w.empty()) return;
    // Line breaks as Enter presses.
    std::wstring t;
    for (size_t i = 0; i < w.size(); i++) {
        if (w[i] == L'\r' && i + 1 < w.size() && w[i + 1] == L'\n') continue;
        t += w[i] == L'\n' ? L'\r' : w[i];
    }
    std::string s = utf8(t);
    if (T.bracketedPaste) s = "\x1b[200~" + s + "\x1b[201~";
    T.viewOff = 0;
    send(s);
}

// ---- keys

bool key(WPARAM vk) {
    const bool ctrl = GetKeyState(VK_CONTROL) < 0, shift = GetKeyState(VK_SHIFT) < 0, alt = GetKeyState(VK_MENU) < 0;
    const int mod = 1 + (shift ? 1 : 0) + (alt ? 2 : 0) + (ctrl ? 4 : 0);
    auto csiKey = [&](char fin) {
        char b[16];
        if (mod > 1) snprintf(b, sizeof(b), "\x1b[1;%d%c", mod, fin);
        else snprintf(b, sizeof(b), T.appCursor ? "\x1bO%c" : "\x1b[%c", fin);
        send(b);
    };
    auto tilde = [&](int n) {
        char b[16];
        if (mod > 1) snprintf(b, sizeof(b), "\x1b[%d;%d~", n, mod);
        else snprintf(b, sizeof(b), "\x1b[%d~", n);
        send(b);
    };
    // Esc hides the panel (full-screen programs such as vim still get it).
    if (vk == VK_ESCAPE && !T.alt && !ctrl && !shift && !alt) {
        T.eatEsc = true;
        PostMessageW(T.parent, WM_APP_TERM, TERM_CLOSE, 0);
        return true;
    }
    // Scrolling back through output.
    if (shift && !ctrl && (vk == VK_PRIOR || vk == VK_NEXT) && !T.alt) {
        const int page = std::max(1, T.rows - 1);
        T.viewOff = std::clamp(T.viewOff + (vk == VK_PRIOR ? page : -page), 0, (int)T.scrollback.size());
        invalidate();
        return true;
    }
    if (ctrl && (vk == VK_OEM_3)) {  // Ctrl+`: the editor
        PostMessageW(T.parent, WM_APP_TERM, TERM_TOGGLE, 0);
        return true;
    }
    if (ctrl && shift && vk == 'C') {
        if (T.hasSel) copySelection();
        return true;
    }
    if (ctrl && (vk == 'V' || (shift && vk == 'V'))) {
        paste();
        return true;
    }
    if (ctrl && vk == 'C' && T.hasSel) {  // with a selection Ctrl+C copies; otherwise it interrupts
        copySelection();
        clearSel();
        invalidate();
        return true;
    }
    if (ctrl && vk == VK_INSERT) {
        if (T.hasSel) copySelection();
        return true;
    }
    if (shift && vk == VK_INSERT) {
        paste();
        return true;
    }
    switch (vk) {
    case VK_UP: csiKey('A'); return true;
    case VK_DOWN: csiKey('B'); return true;
    case VK_RIGHT: csiKey('C'); return true;
    case VK_LEFT: csiKey('D'); return true;
    case VK_HOME: csiKey('H'); return true;
    case VK_END: csiKey('F'); return true;
    case VK_INSERT: tilde(2); return true;
    case VK_DELETE: tilde(3); return true;
    case VK_PRIOR: tilde(5); return true;
    case VK_NEXT: tilde(6); return true;
    case VK_F1: case VK_F2: case VK_F3: case VK_F4: {
        char b[16];
        if (mod > 1) snprintf(b, sizeof(b), "\x1b[1;%d%c", mod, (char)('P' + (vk - VK_F1)));
        else snprintf(b, sizeof(b), "\x1bO%c", (char)('P' + (vk - VK_F1)));
        send(b);
        return true;
    }
    case VK_F5: tilde(15); return true;
    case VK_F6: tilde(17); return true;
    case VK_F7: tilde(18); return true;
    case VK_F8: tilde(19); return true;
    case VK_F9: tilde(20); return true;
    case VK_F10: tilde(21); return true;
    case VK_F11: tilde(23); return true;
    case VK_F12: tilde(24); return true;
    }
    return false;
}

void onChar(wchar_t ch, bool alt) {
    static wchar_t high = 0;
    if (ch == 0x1b && T.eatEsc) {
        T.eatEsc = false;
        return;
    }
    if (!T.running) {
        if (ch == L'\r') startShell();
        return;
    }
    if (ch >= 0xD800 && ch <= 0xDBFF) {
        high = ch;
        return;
    }
    std::wstring w;
    if (ch >= 0xDC00 && ch <= 0xDFFF) {
        if (!high) return;
        w += high;
        high = 0;
    }
    w += ch == L'\b' ? L'\x7f' : ch;  // Backspace is DEL, as in other terminals
    if (ch == 0x7f) w = L"\x08";          // Ctrl+Backspace: delete a word
    std::string s = utf8(w);
    if (alt) s = "\x1b" + s;
    if (T.viewOff) {
        T.viewOff = 0;
        invalidate();
    }
    if (T.hasSel) {
        clearSel();
        invalidate();
    }
    send(s);
}

// ---- window

const UINT_PTR kBlinkTimer = 1;

LRESULT CALLBACK termProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_ERASEBKGND: return 1;
    case WM_PAINT: paint(); return 0;
    case WM_SIZE:
        if (T.font) applySize();
        invalidate();
        return 0;
    case WM_GETDLGCODE: return DLGC_WANTALLKEYS | DLGC_WANTCHARS | DLGC_WANTTAB | DLGC_WANTARROWS;
    case WM_SETFOCUS: {
        T.focused = true;
        T.cursorOn = true;
        UINT t = GetCaretBlinkTime();
        if (t != INFINITE && t) SetTimer(hwnd, kBlinkTimer, t, nullptr);
        if (T.focusEvents) send("\x1b[I");
        invalidate();
        return 0;
    }
    case WM_KILLFOCUS:
        T.focused = false;
        KillTimer(hwnd, kBlinkTimer);
        if (T.focusEvents) send("\x1b[O");
        invalidate();
        return 0;
    case WM_TIMER:
        if (wp == kBlinkTimer) {
            T.cursorOn = !T.cursorOn;
            invalidate();
        }
        return 0;
    case WM_KEYDOWN:
        if (key(wp)) return 0;
        return 0;  // characters arrive as WM_CHAR
    case WM_SYSKEYDOWN:
        if (wp != VK_MENU && wp != VK_F4 && key(wp)) return 0;
        return DefWindowProcW(hwnd, msg, wp, lp);
    case WM_CHAR: onChar((wchar_t)wp, false); return 0;
    case WM_SYSCHAR: onChar((wchar_t)wp, true); return 0;  // Alt+letter: ESC then the letter
    case WM_TERM_OUTPUT: {
        if ((unsigned)wp != T.gen) return 0;
        std::string data;
        {
            std::lock_guard<std::mutex> lk(T.mu);
            data.swap(T.pending);
        }
        feed(data);
        T.cursorOn = true;
        invalidate();
        return 0;
    }
    case WM_TERM_EXITED:
        if ((unsigned)wp != T.gen) return 0;
        stopShell();
        T.title.clear();
        feed("\r\n\x1b[90m[The shell has exited. Press Enter to start a new one.]\x1b[0m\r\n");
        invalidate();
        return 0;
    case WM_MOUSEWHEEL: {
        if (T.alt) {  // full-screen programs: the wheel is arrow keys
            const int n = GET_WHEEL_DELTA_WPARAM(wp) / WHEEL_DELTA * 3;
            for (int i = 0; i < std::abs(n); i++) send(n > 0 ? (T.appCursor ? "\x1bOA" : "\x1b[A") : (T.appCursor ? "\x1bOB" : "\x1b[B"));
            return 0;
        }
        int lines = 3;
        SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
        const int d = GET_WHEEL_DELTA_WPARAM(wp) * lines / WHEEL_DELTA;
        T.viewOff = std::clamp(T.viewOff + d, 0, (int)T.scrollback.size());
        invalidate();
        return 0;
    }
    case WM_SETCURSOR:
        if (LOWORD(lp) == HTCLIENT) {
            POINT pt;
            GetCursorPos(&pt);
            ScreenToClient(hwnd, &pt);
            const bool header = pt.y < T.headerH && !PtInRect(&T.closeRc, pt);
            SetCursor(LoadCursorW(nullptr, header || T.dragging ? IDC_SIZENS : pt.y < T.headerH ? IDC_ARROW : IDC_IBEAM));
            return TRUE;
        }
        break;
    case WM_LBUTTONDOWN: {
        SetFocus(hwnd);
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        if (pt.y < T.headerH) {
            if (PtInRect(&T.closeRc, pt)) {
                PostMessageW(T.parent, WM_APP_TERM, TERM_CLOSE, 0);
                return 0;
            }
            RECT wr;
            GetWindowRect(hwnd, &wr);
            POINT sp = pt;
            ClientToScreen(hwnd, &sp);
            T.dragging = true;
            T.dragStartY = sp.y;
            T.dragStartH = wr.bottom - wr.top;
            SetCapture(hwnd);
            return 0;
        }
        long long line;
        int col;
        cellAt(pt.x, pt.y, line, col);
        T.selA = T.selB = line;
        T.selAc = T.selBc = col;
        T.selecting = true;
        T.hasSel = false;
        SetCapture(hwnd);
        invalidate();
        return 0;
    }
    case WM_LBUTTONDBLCLK: {  // a word
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        if (pt.y < T.headerH) return 0;
        long long line;
        int col;
        cellAt(pt.x, pt.y, line, col);
        const Line* l = lineAt(line);
        if (!l) return 0;
        auto word = [&](int i) { wchar_t c = l->c[(size_t)i].ch; return c && c != L' ' && !wcschr(L"\"'()[]{}<>,;|", c); };
        int a = std::min(col, (int)l->c.size() - 1), b = a;
        if (a < 0 || !word(a)) return 0;
        while (a > 0 && word(a - 1)) a--;
        while (b < (int)l->c.size() && word(b)) b++;
        T.selA = T.selB = line;
        T.selAc = a;
        T.selBc = b;
        T.hasSel = true;
        invalidate();
        return 0;
    }
    case WM_MOUSEMOVE: {
        POINT pt = {GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
        if (T.dragging) {
            POINT sp = pt;
            ClientToScreen(hwnd, &sp);
            PostMessageW(T.parent, WM_APP_TERM, TERM_RESIZE, T.dragStartH + (T.dragStartY - sp.y));
            return 0;
        }
        if (T.selecting) {
            if (pt.y < T.headerH && T.viewOff < (int)T.scrollback.size() && !T.alt) T.viewOff++;
            long long line;
            int col;
            cellAt(pt.x, pt.y, line, col);
            T.selB = line;
            T.selBc = col;
            T.hasSel = T.selA != T.selB || T.selAc != T.selBc;
            invalidate();
            return 0;
        }
        const bool hot = PtInRect(&T.closeRc, pt) != 0;
        if (hot != T.closeHot) {
            T.closeHot = hot;
            invalidate();
            TRACKMOUSEEVENT tme = {sizeof(tme), TME_LEAVE, hwnd, 0};
            TrackMouseEvent(&tme);
        }
        return 0;
    }
    case WM_MOUSELEAVE:
        if (T.closeHot) {
            T.closeHot = false;
            invalidate();
        }
        return 0;
    case WM_LBUTTONUP:
    case WM_CAPTURECHANGED:
        T.dragging = false;
        T.selecting = false;
        if (msg == WM_LBUTTONUP && GetCapture() == hwnd) ReleaseCapture();
        return 0;
    case WM_RBUTTONUP:  // as in the Windows console: copy the selection, or paste
        if (T.hasSel) {
            copySelection();
            clearSel();
            invalidate();
        } else {
            paste();
        }
        return 0;
    case WM_DROPFILES: {
        // Files and folders dropped here are typed at the prompt, quoted for PowerShell.
        HDROP drop = (HDROP)wp;
        std::wstring text;
        const UINT n = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
        for (UINT i = 0; i < n; i++) {
            wchar_t path[32768];
            if (!DragQueryFileW(drop, i, path, 32768)) continue;
            std::wstring p = path;
            const bool plain = p.find_first_of(L" '\"`$&()[]{};,@#%") == std::wstring::npos;
            if (!text.empty()) text += L' ';
            if (plain) {
                text += p;
            } else {
                text += L'\'';
                for (wchar_t c : p) {
                    if (c == L'\'') text += c;
                    text += c;
                }
                text += L'\'';
            }
        }
        DragFinish(drop);
        if (!text.empty() && T.running) {
            T.viewOff = 0;
            send(utf8(text));
        }
        SetForegroundWindow(T.parent);
        SetFocus(hwnd);
        return 0;
    }
    case WM_DESTROY: termShutdown(); return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

}  // namespace

HWND termCreate(HWND parent, HINSTANCE inst) {
    if (T.wnd) return T.wnd;
    WNDCLASSEXW wc = {sizeof(wc)};
    wc.style = CS_DBLCLKS;
    wc.lpfnWndProc = termProc;
    wc.hInstance = inst;
    wc.hCursor = LoadCursorW(nullptr, IDC_IBEAM);
    wc.lpszClassName = kTermClass;
    RegisterClassExW(&wc);
    T.parent = parent;
    T.wnd = CreateWindowExW(0, kTermClass, L"", WS_CHILD | WS_CLIPSIBLINGS, 0, 0, 0, 0, parent, nullptr, inst, nullptr);
    resetScreen(T.mainScr);
    resetScreen(T.altScr);
    DragAcceptFiles(T.wnd, TRUE);
    return T.wnd;
}

void termSetStyle(HFONT font, HFONT uiFont, int cw, int ch, int headerH, const Palette& pal) {
    T.font = font;
    T.uiFont = uiFont;
    T.cw = std::max(1, cw);
    T.ch = std::max(1, ch);
    T.headerH = headerH;
    T.pal = pal;
    if (T.bold) DeleteObject(T.bold);
    LOGFONTW lf = {};
    GetObjectW(font, sizeof(lf), &lf);
    lf.lfWeight = FW_BOLD;
    T.bold = CreateFontIndirectW(&lf);
    if (T.wnd) {
        applySize();
        invalidate();
    }
}

void termStart(const std::wstring& folder) {
    if (T.running) return;
    T.folder = folder;
    startShell();
    invalidate();
}

bool termRunning() { return T.running; }

void termCd(const std::wstring& folder) {
    if (!T.running || folder.empty() || _wcsicmp(folder.c_str(), T.folder.c_str()) == 0) return;
    T.folder = folder;
    if (T.alt) return;  // a full-screen program is running: leave it alone
    std::wstring q;
    for (wchar_t c : folder) {
        if (c == L'\'' || c == L'\x2018' || c == L'\x2019') q += c;  // quotes are doubled in PowerShell
        q += c;
    }
    // Esc clears whatever is typed at the prompt first.
    T.viewOff = 0;
    // Clear-Host after it, so the switch leaves a clean screen with just the new prompt.
    send("\x1b" + utf8(L"Set-Location -LiteralPath '" + q + L"'; Clear-Host") + "\r");
}

void termShutdown() { stopShell(); }
