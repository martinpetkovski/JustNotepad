#pragma once
// Turning bytes into rows of grid cells, working only on the part of the
// document that is on screen. Nothing here depends on the file size.
//
// The editor uses a fixed-pitch grid: every character is 0, 1 or 2 cells wide
// and a tab runs to the next tab stop. That makes column <-> pixel mapping pure
// arithmetic.
#include "doc.h"
#include "types.h"
#include "utf.h"

namespace jnp {

// A row never holds more than this many bytes; longer lines are cut into
// several rows even with word wrap off, so one giant line cannot stall a paint.
constexpr u64 kMaxRowBytes = 16384;
// How far back we look for the start of a line before giving up and treating
// an arbitrary point as a line start (only matters for lines over 1 MB).
constexpr u64 kMaxLineBack = 1u << 20;

struct LayoutCfg {
    int wrapCols = 0;             // 0 = word wrap off
    int tabw = 4;
    const u16* ansi = nullptr;    // 256-entry table for single-byte files, or null for UTF-8
};

struct Row {
    u64 start = 0;  // first byte
    u64 end = 0;    // one past the last visible byte (line ending excluded)
    u64 next = 0;   // where the following row starts
    bool hard = true;  // ends at a newline or the end of the file (not a wrap)
};

inline int decodeChar(const u8* p, u64 n, const u16* ansi, u32& cp) {
    if (ansi) {
        cp = ansi[*p];
        return 1;
    }
    return utf8Decode(p, n, cp);
}

inline int cellsFor(u32 cp, int col, int tabw) {
    if (cp == '\t') return tabw - (col % tabw);
    if (cp < 0x20 || cp == 0x7F) return 1;
    return cellWidth(cp);
}

// The row starting at `start` (which must be a row start).
Row rowAt(const Doc& d, u64 start, const LayoutCfg& cfg);
// The row after `cur`; false at the end of the document.
bool nextRow(const Doc& d, const Row& cur, const LayoutCfg& cfg, Row& out);
// The row before the one starting at `rowStart`; false at the top.
bool prevRow(const Doc& d, u64 rowStart, const LayoutCfg& cfg, Row& out);
// The row that shows byte offset `pos`.
Row rowContaining(const Doc& d, u64 pos, const LayoutCfg& cfg);
// Start of the line containing pos (search stops at the byte before `pos`).
u64 lineStartBefore(const Doc& d, u64 pos, const LayoutCfg& cfg);
bool rowHas(const Doc& d, const Row& r, u64 pos);

int colOf(const Doc& d, const Row& r, u64 pos, const LayoutCfg& cfg);
u64 posOfCol(const Doc& d, const Row& r, int col, const LayoutCfg& cfg);

u64 nextChar(const Doc& d, u64 pos, const LayoutCfg& cfg);
u64 prevChar(const Doc& d, u64 pos, const LayoutCfg& cfg);
u64 wordRight(const Doc& d, u64 pos, const LayoutCfg& cfg);
u64 wordLeft(const Doc& d, u64 pos, const LayoutCfg& cfg);
void wordAt(const Doc& d, u64 pos, u64& start, u64& end);

}  // namespace jnp
