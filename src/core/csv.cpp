#include "csv.h"

namespace jnp {

char csvSniff(const char* p, size_t n, bool tsvName) {
    if (tsvName) return '\t';
    // Count candidates outside quotes over the first few lines; the most
    // frequent, and consistent, one wins.
    const char cand[4] = {',', ';', '\t', '|'};
    int count[4] = {0, 0, 0, 0};
    bool q = false;
    int lines = 0;
    for (size_t i = 0; i < n && i < (64u << 10) && lines < 20; i++) {
        char c = p[i];
        if (c == '"') q = !q;
        else if (!q && c == '\n') lines++;
        else if (!q)
            for (int k = 0; k < 4; k++)
                if (c == cand[k]) count[k]++;
    }
    int best = 0;
    for (int k = 1; k < 4; k++)
        if (count[k] > count[best]) best = k;
    return count[best] ? cand[best] : ',';
}

std::vector<u64> csvRows(const char* p, size_t n, size_t maxRows, bool& truncated) {
    std::vector<u64> rows;
    truncated = false;
    if (n == 0) {
        rows.push_back(0);
        return rows;
    }
    rows.push_back(0);
    bool q = false;
    for (size_t i = 0; i < n; i++) {
        char c = p[i];
        if (c == '"') q = !q;
        else if (c == '\n' && !q && i + 1 < n) {
            if (rows.size() >= maxRows) {
                truncated = true;
                break;
            }
            rows.push_back(i + 1);
        }
    }
    rows.push_back(truncated ? rows.back() : n);
    if (truncated) {
        // The end of the last row kept: the line break after it.
        size_t e = (size_t)rows[rows.size() - 2];
        bool qq = false;
        while (e < n && !(p[e] == '\n' && !qq)) {
            if (p[e] == '"') qq = !qq;
            e++;
        }
        rows.back() = e < n ? e + 1 : n;
    }
    return rows;
}

void csvFields(const char* p, size_t n, u64 base, char delim, std::vector<CsvField>& out) {
    out.clear();
    size_t i = 0;
    for (;;) {
        CsvField f;
        f.start = base + i;
        if (i < n && p[i] == '"') {
            i++;
            while (i < n) {
                if (p[i] == '"') {
                    if (i + 1 < n && p[i + 1] == '"') {
                        i += 2;
                        continue;
                    }
                    i++;
                    break;
                }
                i++;
            }
            while (i < n && p[i] != delim) i++;  // stray text after the closing quote
        } else {
            while (i < n && p[i] != delim) i++;
        }
        f.end = base + i;
        out.push_back(f);
        if (i >= n) break;
        i++;  // the delimiter
        if (i == n) {  // a trailing delimiter: one more, empty field
            CsvField e;
            e.start = e.end = base + i;
            out.push_back(e);
            break;
        }
    }
}

std::string csvValue(const char* p, size_t n) {
    while (n && (p[n - 1] == '\r')) n--;
    if (n >= 1 && p[0] == '"') {
        std::string v;
        size_t i = 1;
        while (i < n) {
            if (p[i] == '"') {
                if (i + 1 < n && p[i + 1] == '"') {
                    v += '"';
                    i += 2;
                    continue;
                }
                break;
            }
            v += p[i++];
        }
        return v;
    }
    return std::string(p, n);
}

std::string csvQuote(const std::string& v, char delim) {
    bool need = v.find(delim) != std::string::npos || v.find('"') != std::string::npos ||
                v.find('\n') != std::string::npos || v.find('\r') != std::string::npos ||
                (!v.empty() && (v.front() == ' ' || v.back() == ' '));
    if (!need) return v;
    std::string out = "\"";
    for (char c : v) {
        if (c == '"') out += '"';
        out += c;
    }
    return out + "\"";
}

}  // namespace jnp
