#include "layout.h"

#include <algorithm>

namespace jnp {

namespace {
std::vector<u8>& scratch() {
    static thread_local std::vector<u8> buf;
    return buf;
}

const u8* fetch(const Doc& d, u64 start, u64 n) {
    std::vector<u8>& b = scratch();
    if (b.size() < n + 8) b.resize((size_t)n + 8);
    d.copy(start, b.data(), n);
    return b.data();
}

inline int cls(u8 b) {
    if (b == '\n' || b == '\r') return 3;
    if (b == ' ' || b == '\t') return 0;
    if (b >= 0x80 || (b >= '0' && b <= '9') || (b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z') ||
        b == '_')
        return 2;
    return 1;
}
const u64 kWordScan = 1u << 20;
}  // namespace

Row rowAt(const Doc& d, u64 start, const LayoutCfg& cfg) {
    u64 size = d.size();
    Row r;
    if (start >= size) {
        r.start = r.end = r.next = size;
        r.hard = true;
        return r;
    }
    u64 lim = std::min(size, start + kMaxRowBytes);
    u64 nl = d.findFwd('\n', start, lim);
    u64 lineEnd, next;
    bool hard = true;
    if (nl != NPOS) {
        lineEnd = nl;
        next = nl + 1;
        if (lineEnd > start && d.at(lineEnd - 1) == '\r') lineEnd--;
    } else if (lim == size) {
        lineEnd = next = size;
    } else {
        // No newline within the cap: cut here, on a character boundary.
        lineEnd = lim;
        if (!cfg.ansi)
            while (lineEnd > start + 1 && (d.at(lineEnd) & 0xC0) == 0x80) lineEnd--;
        if (lineEnd > start + 1 && d.at(lineEnd - 1) == '\r' && d.at(lineEnd) == '\n') lineEnd--;
        next = lineEnd;
        hard = false;
    }
    r.start = start;
    r.end = lineEnd;
    r.next = next;
    r.hard = hard;
    if (cfg.wrapCols <= 0 || lineEnd == start) return r;

    // Word wrap: walk cells until the row is full.
    u64 full = lineEnd - start;
    u64 want = std::min<u64>(full, (u64)cfg.wrapCols * 4 + 16);
    bool truncated = want < full;
    const u8* b = fetch(d, start, want);
    int col = 0;
    u64 i = 0, lastBreak = 0;
    while (i < want) {
        if (truncated && i + 4 > want) {  // ran out of fetched bytes (many zero-width chars)
            r.end = r.next = start + i;
            r.hard = false;
            return r;
        }
        u32 cp;
        int len = decodeChar(b + i, want - i, cfg.ansi, cp);
        int w = cellsFor(cp, col, cfg.tabw);
        if (col + w > cfg.wrapCols && i > 0) {
            u64 at = (lastBreak > 0 && lastBreak <= i) ? lastBreak : i;
            r.end = r.next = start + at;
            r.hard = false;
            return r;
        }
        col += w;
        i += (u64)len;
        if (cp == ' ' || cp == '\t') lastBreak = i;
    }
    return r;
}

bool nextRow(const Doc& d, const Row& cur, const LayoutCfg& cfg, Row& out) {
    u64 size = d.size();
    if (cur.next < size) {
        out = rowAt(d, cur.next, cfg);
        return true;
    }
    if (cur.end < cur.next) {  // the last line ends with a newline: one empty row follows
        out.start = out.end = out.next = size;
        out.hard = true;
        return true;
    }
    return false;
}

u64 lineStartBefore(const Doc& d, u64 pos, const LayoutCfg& cfg) {
    u64 lb = pos > kMaxLineBack ? pos - kMaxLineBack : 0;
    u64 nl = d.findBack('\n', pos, lb);
    if (nl != NPOS) return nl + 1;
    if (lb == 0) return 0;
    if (!cfg.ansi)
        while (lb < pos && (d.at(lb) & 0xC0) == 0x80) lb++;
    return lb;
}

bool prevRow(const Doc& d, u64 rowStart, const LayoutCfg& cfg, Row& out) {
    if (rowStart == 0) return false;
    u64 size = d.size();
    if (rowStart > size) rowStart = size;
    u64 ls = lineStartBefore(d, rowStart - 1, cfg);
    Row r = rowAt(d, ls, cfg);
    while (r.next < rowStart && r.next > r.start) r = rowAt(d, r.next, cfg);
    out = r;
    return true;
}

bool rowHas(const Doc& d, const Row& r, u64 pos) {
    if (pos < r.start) return false;
    if (pos < r.next) return true;
    return pos == r.next && r.end == r.next && r.next == d.size() && r.hard;
}

Row rowContaining(const Doc& d, u64 pos, const LayoutCfg& cfg) {
    u64 size = d.size();
    if (pos > size) pos = size;
    Row r = rowAt(d, lineStartBefore(d, pos, cfg), cfg);
    while (pos >= r.next && r.next < size && r.next > r.start) r = rowAt(d, r.next, cfg);
    if (pos >= r.next && r.end < r.next) {  // caret after the final newline
        r.start = r.end = r.next = size;
        r.hard = true;
    }
    return r;
}

int colOf(const Doc& d, const Row& r, u64 pos, const LayoutCfg& cfg) {
    if (pos > r.end) pos = r.end;
    if (pos <= r.start) return 0;
    u64 n = pos - r.start;
    const u8* b = fetch(d, r.start, n);
    int col = 0;
    u64 i = 0;
    while (i < n) {
        u32 cp;
        i += (u64)decodeChar(b + i, n - i, cfg.ansi, cp);
        col += cellsFor(cp, col, cfg.tabw);
    }
    return col;
}

u64 posOfCol(const Doc& d, const Row& r, int target, const LayoutCfg& cfg) {
    if (target <= 0 || r.end <= r.start) return r.start;
    u64 n = r.end - r.start;
    const u8* b = fetch(d, r.start, n);
    int col = 0;
    u64 i = 0;
    while (i < n) {
        u32 cp;
        int len = decodeChar(b + i, n - i, cfg.ansi, cp);
        int w = cellsFor(cp, col, cfg.tabw);
        if (col >= target) return r.start + i;
        if (col + w > target) return (target - col) * 2 >= w ? r.start + i + (u64)len : r.start + i;
        col += w;
        i += (u64)len;
    }
    return r.end;
}

u64 nextChar(const Doc& d, u64 pos, const LayoutCfg& cfg) {
    u64 size = d.size();
    if (pos >= size) return size;
    u8 b[4];
    u64 n = d.copy(pos, b, 4);
    if (b[0] == '\r' && n > 1 && b[1] == '\n') return pos + 2;
    if (cfg.ansi) return pos + 1;
    u32 cp;
    u64 p = pos + (u64)utf8Decode(b, n, cp);
    // Keep combining marks attached to their base character.
    for (int guard = 0; guard < 8 && p < size; guard++) {
        u64 m = d.copy(p, b, 4);
        int len = utf8Decode(b, m, cp);
        if (cp < 0x300 || cellWidth(cp) != 0) break;
        p += (u64)len;
    }
    return p;
}

static u64 prevCodepoint(const Doc& d, u64 pos) {
    u64 k = 1;
    while (k < 4 && k < pos && (d.at(pos - k) & 0xC0) == 0x80) k++;
    u8 b[4];
    u64 n = d.copy(pos - k, b, k);
    u32 cp;
    if ((u64)utf8Decode(b, n, cp) != k) return pos - 1;
    return pos - k;
}

u64 prevChar(const Doc& d, u64 pos, const LayoutCfg& cfg) {
    if (pos == 0) return 0;
    if (pos >= 2 && d.at(pos - 1) == '\n' && d.at(pos - 2) == '\r') return pos - 2;
    if (cfg.ansi) return pos - 1;
    u64 p = prevCodepoint(d, pos);
    for (int guard = 0; guard < 8 && p > 0; guard++) {
        u8 b[4];
        u64 n = d.copy(p, b, 4);
        u32 cp;
        utf8Decode(b, n, cp);
        if (cp < 0x300 || cellWidth(cp) != 0) break;
        p = prevCodepoint(d, p);
    }
    return p;
}

u64 wordRight(const Doc& d, u64 pos, const LayoutCfg& cfg) {
    u64 size = d.size();
    if (pos >= size) return size;
    int c = cls(d.at(pos));
    if (c == 3) return nextChar(d, pos, cfg);
    u64 lim = std::min(size, pos + kWordScan);
    u64 p = pos;
    if (c != 0)
        while (p < lim && cls(d.at(p)) == c) p++;
    while (p < lim && cls(d.at(p)) == 0) p++;
    return p;
}

u64 wordLeft(const Doc& d, u64 pos, const LayoutCfg& cfg) {
    if (pos == 0) return 0;
    if (cls(d.at(pos - 1)) == 3) return prevChar(d, pos, cfg);
    u64 lim = pos > kWordScan ? pos - kWordScan : 0;
    u64 p = pos;
    while (p > lim && cls(d.at(p - 1)) == 0) p--;
    if (p > lim) {
        int c = cls(d.at(p - 1));
        if (c != 3)
            while (p > lim && cls(d.at(p - 1)) == c) p--;
    }
    return p;
}

void wordAt(const Doc& d, u64 pos, u64& start, u64& end) {
    u64 size = d.size();
    start = end = pos;
    if (size == 0) return;
    if (pos >= size) pos = size - 1;
    int c = cls(d.at(pos));
    if (c == 3) {
        start = end = pos;
        return;
    }
    u64 lo = pos > kWordScan ? pos - kWordScan : 0;
    u64 hi = std::min(size, pos + kWordScan);
    start = pos;
    while (start > lo && cls(d.at(start - 1)) == c) start--;
    end = pos;
    while (end < hi && cls(d.at(end)) == c) end++;
}

}  // namespace jnp
