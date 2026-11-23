#pragma once
// UTF-8 / UTF-16 helpers. No OS dependencies.
#include <string>
#include <vector>

#include "types.h"

namespace jnp {

inline int utf8ExpectedLen(u8 b) {
    if (b < 0x80) return 1;
    if (b >= 0xC2 && b <= 0xDF) return 2;
    if (b >= 0xE0 && b <= 0xEF) return 3;
    if (b >= 0xF0 && b <= 0xF4) return 4;
    return 1;
}

// Decodes one code point from p[0..n). Returns the bytes consumed (>= 1).
// An invalid or truncated sequence yields U+FFFD and consumes one byte.
int utf8Decode(const u8* p, u64 n, u32& cp);

// Writes up to 4 bytes, returns the count.
int utf8Encode(u32 cp, u8* out);

// True when the sample is valid UTF-8 (a sequence cut off at the end is fine).
bool looksUtf8(const u8* p, u64 n);

// Display width in grid cells: 0 (combining), 1 or 2 (East Asian wide, emoji).
int cellWidth(u32 cp);

// Converts UTF-16 bytes (without BOM) to UTF-8, appending to out.
void utf16ToUtf8(const u8* bytes, u64 nbytes, bool bigEndian, std::vector<u8>& out);

// Streaming UTF-8 -> UTF-16 converter. Sequences may be split across feeds.
class Utf8ToUtf16 {
public:
    void feed(const u8* p, u64 n, std::vector<u16>& out);
    void finish(std::vector<u16>& out);

private:
    u8 carry_[8] = {};
    int n_ = 0;
};

// Rewrites every line ending in s to CRLF or LF.
void normalizeEol(std::string& s, bool crlf);

}  // namespace jnp
