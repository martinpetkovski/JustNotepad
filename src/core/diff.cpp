#include "diff.h"

#include <algorithm>

namespace jnp {

namespace {
const u64 kFnvOffset = 1469598103934665603ull, kFnvPrime = 1099511628211ull;
}

void LineHasher::feed(const u8* p, u64 n) {
    u64 h = h_;
    for (u64 i = 0; i < n; i++) {
        u8 c = p[i];
        if (cr_) {
            cr_ = false;
            if (c != '\n') h = (h ^ '\r') * kFnvPrime;  // a lone CR is ordinary content
        }
        if (c == '\n') {
            lines.push_back(h);
            h = kFnvOffset;
            any_ = false;
        } else if (c == '\r') {
            cr_ = true;
            any_ = true;
        } else {
            h = (h ^ c) * kFnvPrime;
            any_ = true;
        }
    }
    h_ = h;
}

void LineHasher::finish() {
    if (cr_) h_ = (h_ ^ '\r') * kFnvPrime;
    if (any_) lines.push_back(h_);
    h_ = kFnvOffset;
    any_ = cr_ = false;
}

void hashDocLines(const Doc& d, std::vector<u64>& out) {
    LineHasher h;
    d.forEachPiece([&](const u8* p, u64 n) {
        h.feed(p, n);
        return true;
    });
    h.finish();
    out.swap(h.lines);
}

std::vector<Hunk> diffLines(const std::vector<u64>& a, const std::vector<u64>& b) {
    std::vector<Hunk> out;
    size_t n = a.size(), m = b.size();
    size_t pre = 0;
    while (pre < n && pre < m && a[pre] == b[pre]) pre++;
    size_t suf = 0;
    while (suf < n - pre && suf < m - pre && a[n - 1 - suf] == b[m - 1 - suf]) suf++;
    const i64 N = (i64)(n - pre - suf), M = (i64)(m - pre - suf);
    if (N == 0 && M == 0) return out;
    auto whole = [&]() {
        out.push_back(Hunk{(u32)pre, (u32)N, (u32)pre, (u32)M});
        return out;
    };
    if (N == 0 || M == 0) return whole();
    const u64* A = a.data() + pre;
    const u64* B = b.data() + pre;

    // Keep the worst case bounded: this runs on the UI thread.
    const i64 kMaxD = 1500;
    const i64 kBudget = 40ll * 1000 * 1000;
    const i64 maxD = std::min<i64>(N + M, kMaxD);
    const i64 off = maxD + 1;
    std::vector<i64> V((size_t)(2 * maxD + 3), 0);
    std::vector<std::vector<i64>> trace;
    i64 foundD = -1, work = 0;
    for (i64 d = 0; d <= maxD && foundD < 0; d++) {
        trace.emplace_back(V.begin() + (off - d), V.begin() + (off + d + 1));  // state before step d
        for (i64 k = -d; k <= d; k += 2) {
            i64 x = (k == -d || (k != d && V[(size_t)(off + k - 1)] < V[(size_t)(off + k + 1)]))
                        ? V[(size_t)(off + k + 1)]
                        : V[(size_t)(off + k - 1)] + 1;
            i64 y = x - k;
            while (x < N && y < M && A[x] == B[y]) {
                x++;
                y++;
                work++;
            }
            V[(size_t)(off + k)] = x;
            if (x >= N && y >= M) {
                foundD = d;
                break;
            }
        }
        work += 2 * d + 1;
        if (work > kBudget) break;
    }
    if (foundD < 0) return whole();

    // Walk back through the saved states to find which lines were removed/added.
    std::vector<char> delA((size_t)N, 0), insB((size_t)M, 0);
    i64 x = N, y = M;
    for (i64 d = foundD; d > 0; d--) {
        const std::vector<i64>& Vp = trace[(size_t)d];  // indices -d..d stored at [k + d]
        auto vp = [&](i64 k) { return Vp[(size_t)(k + d)]; };
        i64 k = x - y;
        i64 pk = (k == -d || (k != d && vp(k - 1) < vp(k + 1))) ? k + 1 : k - 1;
        i64 px = vp(pk), py = px - pk;
        while (x > px && y > py) {
            x--;
            y--;
        }
        if (x == px) insB[(size_t)py] = 1;
        else delA[(size_t)px] = 1;
        x = px;
        y = py;
    }
    i64 i = 0, j = 0;
    while (i < N || j < M) {
        bool da = i < N && delA[(size_t)i], ib = j < M && insB[(size_t)j];
        if (!da && !ib) {
            i++;
            j++;
            continue;
        }
        i64 as = i, bs = j;
        for (;;) {
            bool moved = false;
            while (i < N && delA[(size_t)i]) {
                i++;
                moved = true;
            }
            while (j < M && insB[(size_t)j]) {
                j++;
                moved = true;
            }
            if (!moved) break;
        }
        out.push_back(Hunk{(u32)(pre + as), (u32)(i - as), (u32)(pre + bs), (u32)(j - bs)});
    }
    return out;
}

int hunkAtB(const std::vector<Hunk>& hunks, u32 line) {
    // Last hunk starting at or before `line`.
    size_t lo = 0, hi = hunks.size();
    while (lo < hi) {
        size_t mid = (lo + hi) / 2;
        if (hunks[mid].bStart <= line) lo = mid + 1;
        else hi = mid;
    }
    if (lo == 0) return -1;
    const Hunk& h = hunks[lo - 1];
    if (h.bCount == 0) return h.bStart == line ? (int)(lo - 1) : -1;
    return line < h.bStart + h.bCount ? (int)(lo - 1) : -1;
}

int hunkAtA(const std::vector<Hunk>& hunks, u32 line) {
    size_t lo = 0, hi = hunks.size();
    while (lo < hi) {
        size_t mid = (lo + hi) / 2;
        if (hunks[mid].aStart <= line) lo = mid + 1;
        else hi = mid;
    }
    if (lo == 0) return -1;
    const Hunk& h = hunks[lo - 1];
    if (h.aCount == 0) return h.aStart == line ? (int)(lo - 1) : -1;
    return line < h.aStart + h.aCount ? (int)(lo - 1) : -1;
}

i64 mapBtoA(const std::vector<Hunk>& hunks, i64 line) {
    if (line < 0) return line;
    size_t lo = 0, hi = hunks.size();
    while (lo < hi) {
        size_t mid = (lo + hi) / 2;
        if ((i64)hunks[mid].bStart <= line) lo = mid + 1;
        else hi = mid;
    }
    if (lo == 0) return line;
    const Hunk& h = hunks[lo - 1];
    i64 inside = line - (i64)h.bStart;
    if (inside < (i64)h.bCount) return (i64)h.aStart + std::min<i64>(inside, h.aCount);
    return (i64)h.aStart + h.aCount + (inside - (i64)h.bCount);
}

}  // namespace jnp
