#include "utf.h"

namespace jnp {

int utf8Decode(const u8* p, u64 n, u32& cp) {
    u8 b = p[0];
    if (b < 0x80) {
        cp = b;
        return 1;
    }
    int len = utf8ExpectedLen(b);
    if (len == 1 || (u64)len > n) {
        cp = 0xFFFD;
        return 1;
    }
    u32 c = len == 2 ? (b & 0x1Fu) : len == 3 ? (b & 0x0Fu) : (b & 0x07u);
    for (int i = 1; i < len; i++) {
        if ((p[i] & 0xC0) != 0x80) {
            cp = 0xFFFD;
            return 1;
        }
        c = (c << 6) | (p[i] & 0x3Fu);
    }
    if ((len == 3 && (c < 0x800 || (c >= 0xD800 && c <= 0xDFFF))) ||
        (len == 4 && (c < 0x10000 || c > 0x10FFFF))) {
        cp = 0xFFFD;
        return 1;
    }
    cp = c;
    return len;
}

int utf8Encode(u32 cp, u8* out) {
    if (cp < 0x80) {
        out[0] = (u8)cp;
        return 1;
    }
    if (cp < 0x800) {
        out[0] = (u8)(0xC0 | (cp >> 6));
        out[1] = (u8)(0x80 | (cp & 0x3F));
        return 2;
    }
    if (cp < 0x10000) {
        out[0] = (u8)(0xE0 | (cp >> 12));
        out[1] = (u8)(0x80 | ((cp >> 6) & 0x3F));
        out[2] = (u8)(0x80 | (cp & 0x3F));
        return 3;
    }
    out[0] = (u8)(0xF0 | (cp >> 18));
    out[1] = (u8)(0x80 | ((cp >> 12) & 0x3F));
    out[2] = (u8)(0x80 | ((cp >> 6) & 0x3F));
    out[3] = (u8)(0x80 | (cp & 0x3F));
    return 4;
}

bool looksUtf8(const u8* p, u64 n) {
    u64 i = 0;
    while (i < n) {
        if (p[i] < 0x80) {
            i++;
            continue;
        }
        int len = utf8ExpectedLen(p[i]);
        if (len == 1) return false;
        if (i + (u64)len > n) {
            // Cut off by the end of the sample: accept if the tail is plausible.
            for (u64 k = i + 1; k < n; k++)
                if ((p[k] & 0xC0) != 0x80) return false;
            return true;
        }
        u32 cp;
        if (utf8Decode(p + i, n - i, cp) != len) return false;
        i += (u64)len;
    }
    return true;
}

int cellWidth(u32 cp) {
    if (cp < 0x300) return 1;
    // Combining marks and zero-width characters.
    if ((cp >= 0x0300 && cp <= 0x036F) || (cp >= 0x1AB0 && cp <= 0x1AFF) ||
        (cp >= 0x1DC0 && cp <= 0x1DFF) || (cp >= 0x200B && cp <= 0x200F) ||
        (cp >= 0x20D0 && cp <= 0x20FF) || (cp >= 0xFE00 && cp <= 0xFE0F) ||
        (cp >= 0xFE20 && cp <= 0xFE2F) || cp == 0xFEFF)
        return 0;
    if (cp < 0x1100) return 1;
    // East Asian wide and emoji blocks.
    if ((cp <= 0x115F) || (cp >= 0x2E80 && cp <= 0x303E) || (cp >= 0x3041 && cp <= 0x33FF) ||
        (cp >= 0x3400 && cp <= 0x4DBF) || (cp >= 0x4E00 && cp <= 0x9FFF) ||
        (cp >= 0xA000 && cp <= 0xA4CF) || (cp >= 0xAC00 && cp <= 0xD7A3) ||
        (cp >= 0xF900 && cp <= 0xFAFF) || (cp >= 0xFE30 && cp <= 0xFE4F) ||
        (cp >= 0xFF00 && cp <= 0xFF60) || (cp >= 0xFFE0 && cp <= 0xFFE6) ||
        (cp >= 0x1F300 && cp <= 0x1F64F) || (cp >= 0x1F900 && cp <= 0x1FAFF) ||
        (cp >= 0x20000 && cp <= 0x3FFFD))
        return 2;
    return 1;
}

void utf16ToUtf8(const u8* b, u64 nbytes, bool be, std::vector<u8>& out) {
    u64 units = nbytes / 2;
    out.reserve(out.size() + units + units / 2);
    auto unit = [&](u64 i) -> u32 {
        return be ? (u32)((b[2 * i] << 8) | b[2 * i + 1]) : (u32)((b[2 * i + 1] << 8) | b[2 * i]);
    };
    u8 tmp[4];
    for (u64 i = 0; i < units; i++) {
        u32 c = unit(i);
        if (c >= 0xD800 && c <= 0xDBFF && i + 1 < units) {
            u32 d = unit(i + 1);
            if (d >= 0xDC00 && d <= 0xDFFF) {
                c = 0x10000 + ((c - 0xD800) << 10) + (d - 0xDC00);
                i++;
            } else {
                c = 0xFFFD;
            }
        } else if (c >= 0xD800 && c <= 0xDFFF) {
            c = 0xFFFD;
        }
        if (c < 0x80) {
            out.push_back((u8)c);
        } else {
            int k = utf8Encode(c, tmp);
            out.insert(out.end(), tmp, tmp + k);
        }
    }
    if (nbytes & 1) {  // stray trailing byte
        out.push_back(0xEF);
        out.push_back(0xBF);
        out.push_back(0xBD);
    }
}

static inline void emit16(u32 cp, std::vector<u16>& out) {
    if (cp < 0x10000) {
        out.push_back((u16)cp);
    } else {
        cp -= 0x10000;
        out.push_back((u16)(0xD800 + (cp >> 10)));
        out.push_back((u16)(0xDC00 + (cp & 0x3FF)));
    }
}

void Utf8ToUtf16::feed(const u8* p, u64 n, std::vector<u16>& out) {
    // Finish whatever was left over from the previous feed.
    while (n_ > 0) {
        int need = utf8ExpectedLen(carry_[0]);
        while (n_ < need && n > 0) {
            carry_[n_++] = *p++;
            n--;
        }
        if (n_ < need) return;  // still incomplete; wait for more input
        u32 cp;
        int len = utf8Decode(carry_, (u64)n_, cp);
        emit16(cp, out);
        for (int i = len; i < n_; i++) carry_[i - len] = carry_[i];
        n_ -= len;
    }
    u64 i = 0;
    while (i < n) {
        u8 b = p[i];
        if (b < 0x80) {
            out.push_back(b);
            i++;
            continue;
        }
        int need = utf8ExpectedLen(b);
        if (i + (u64)need > n) {
            while (i < n) carry_[n_++] = p[i++];
            break;
        }
        u32 cp;
        i += (u64)utf8Decode(p + i, n - i, cp);
        emit16(cp, out);
    }
}

void Utf8ToUtf16::finish(std::vector<u16>& out) {
    int i = 0;
    while (i < n_) {
        u32 cp;
        i += utf8Decode(carry_ + i, (u64)(n_ - i), cp);
        emit16(cp, out);
    }
    n_ = 0;
}

void normalizeEol(std::string& s, bool crlf) {
    std::string o;
    o.reserve(s.size() + s.size() / 16);
    for (size_t i = 0; i < s.size(); i++) {
        char c = s[i];
        if (c == '\r') {
            if (i + 1 < s.size() && s[i + 1] == '\n') i++;
            if (crlf) o += '\r';
            o += '\n';
        } else if (c == '\n') {
            if (crlf) o += '\r';
            o += '\n';
        } else {
            o += c;
        }
    }
    s.swap(o);
}

}  // namespace jnp
