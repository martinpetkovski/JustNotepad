// Tests for the portable core. Builds anywhere (no Windows headers):
//   g++ -std=c++20 -O2 -pthread tests/core_tests.cpp src/core/*.cpp -o core_tests
//
// Most of this is differential testing: apply thousands of random edits to the
// piece tree and to a plain std::string, and require identical results.
#include <chrono>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <random>
#include <string>
#include <thread>
#include <vector>

#include "../src/core/diff.h"
#include "../src/core/doc.h"
#include "../src/core/layout.h"
#include "../src/core/lineindex.h"
#include "../src/core/utf.h"

using namespace jnp;

static int g_fail = 0;
static long g_checks = 0;
#define CHECK(c)                                                    \
    do {                                                            \
        g_checks++;                                                 \
        if (!(c)) {                                                 \
            if (g_fail++ < 20) printf("FAIL %s:%d  %s\n", __FILE__, __LINE__, #c); \
        }                                                           \
    } while (0)

static std::mt19937_64 rng(12345);
static u64 rnd(u64 n) { return n ? rng() % n : 0; }

static std::string randomText(u64 n) {
    static const char* bits[] = {"a", "b", "hello", " ", " ", "\n", "\r\n", "\t", "x_y", "(", ";",
                                 "\xC3\xA9", "\xE4\xB8\xAD", "\xF0\x9F\x98\x80", "e\xCC\x81", "Z"};
    std::string s;
    while (s.size() < n) s += bits[rnd(sizeof(bits) / sizeof(bits[0]))];
    return s;
}

static std::string docText(const Doc& d) {
    std::string s;
    d.forEachPiece([&](const u8* p, u64 n) {
        s.append((const char*)p, (size_t)n);
        return true;
    });
    return s;
}

static u64 refLineOf(const std::string& s, u64 pos) {
    u64 c = 0;
    for (u64 i = 0; i < pos; i++) c += s[i] == '\n';
    return c;
}
static u64 refLineStart(const std::string& s, u64 line) {
    if (line == 0) return 0;
    u64 c = 0;
    for (u64 i = 0; i < s.size(); i++)
        if (s[i] == '\n' && ++c == line) return i + 1;
    return s.size();
}

static void verify(const Doc& d, const std::string& ref, bool lines) {
    CHECK(d.size() == ref.size());
    CHECK(docText(d) == ref);
    for (int i = 0; i < 6; i++) {
        u64 pos = rnd(ref.size() + 1);
        u64 n = rnd(64);
        CHECK(d.text(pos, n) == ref.substr((size_t)pos, (size_t)n));
        if (pos < ref.size()) CHECK(d.at(pos) == (u8)ref[(size_t)pos]);
        // single-byte search both ways
        u64 lim = pos + rnd(ref.size() - pos + 1);
        size_t f = ref.find('\n', (size_t)pos);
        u64 expect = (f == std::string::npos || f >= lim) ? NPOS : (u64)f;
        CHECK(d.findFwd('\n', pos, lim) == expect);
        u64 lo = rnd(pos + 1);
        u64 eb = NPOS;
        for (u64 k = pos; k > lo; k--)
            if (ref[(size_t)k - 1] == '\n') {
                eb = k - 1;
                break;
            }
        CHECK(d.findBack('\n', pos, lo) == eb);
        if (lines && d.linesKnown()) {
            CHECK(d.lineOf(pos) == refLineOf(ref, pos));
            u64 total = refLineOf(ref, ref.size());
            CHECK(d.newlineCount() == total);
            u64 line = rnd(total + 2);
            CHECK(d.lineStart(line) == refLineStart(ref, line));
        }
    }
}

static void testEdits(const std::string& original, bool asyncIndex, int steps) {
    Doc d;
    LineIndex idx;
    std::string ref = original;
    if (asyncIndex) {
        // Index not ready yet: edits must work without line information.
        d.reset((const u8*)original.data(), original.size(), &idx);
        CHECK(original.empty() || !d.linesKnown());
    } else {
        idx.buildSync((const u8*)original.data(), original.size());
        d.reset((const u8*)original.data(), original.size(), &idx);
        CHECK(d.linesKnown());
    }
    std::vector<std::string> undoStack{ref}, redoStack;
    for (int step = 0; step < steps; step++) {
        if (asyncIndex && step == steps / 2) {
            idx.buildSync((const u8*)original.data(), original.size());
            d.indexReady();
            CHECK(d.linesKnown());
        }
        int op = (int)rnd(10);
        u64 caret = 0;
        if (op < 4) {
            u64 pos = rnd(ref.size() + 1);
            std::string t = randomText(1 + rnd(12));
            d.insert(pos, (const u8*)t.data(), t.size());
            ref.insert((size_t)pos, t);
            undoStack.push_back(ref);
            redoStack.clear();
        } else if (op < 7) {
            if (ref.empty()) continue;
            u64 pos = rnd(ref.size());
            u64 n = 1 + rnd(std::min<u64>(20, ref.size() - pos));
            d.erase(pos, n);
            ref.erase((size_t)pos, (size_t)n);
            undoStack.push_back(ref);
            redoStack.clear();
        } else if (op == 7) {
            // grouped replace = one undo step
            if (ref.empty()) continue;
            u64 pos = rnd(ref.size());
            u64 n = 1 + rnd(std::min<u64>(8, ref.size() - pos));
            std::string t = randomText(1 + rnd(8));
            d.beginGroup();
            d.erase(pos, n);
            d.insert(pos, (const u8*)t.data(), t.size());
            d.endGroup();
            ref.erase((size_t)pos, (size_t)n);
            ref.insert((size_t)pos, t);
            undoStack.push_back(ref);
            redoStack.clear();
        } else if (op == 8) {
            bool ok = d.undo(caret);
            CHECK(ok == (undoStack.size() > 1));
            if (ok) {
                redoStack.push_back(undoStack.back());
                undoStack.pop_back();
                ref = undoStack.back();
                CHECK(caret <= ref.size());
            }
        } else {
            bool ok = d.redo(caret);
            CHECK(ok == !redoStack.empty());
            if (ok) {
                ref = redoStack.back();
                undoStack.push_back(ref);
                redoStack.pop_back();
            }
        }
        if (step % 7 == 0) verify(d, ref, true);
    }
    verify(d, ref, true);
    // Undo everything: must be the original again, and unmodified.
    u64 caret;
    while (d.undo(caret)) {}
    CHECK(docText(d) == original);
    CHECK(!d.modified());
}

static void testTyping() {
    // Typing coalesces into one undo step and a tiny tree.
    Doc d;
    d.reset(nullptr, 0, nullptr);
    std::string ref;
    for (int i = 0; i < 5000; i++) {
        char c = (char)('a' + i % 26);
        d.insert(d.size(), (const u8*)&c, 1, true);
        ref += c;
    }
    CHECK(docText(d) == ref);
    CHECK(d.pieceCount() == 1);
    CHECK(d.modified());
    d.markSaved();
    CHECK(!d.modified());
    d.insert(d.size(), (const u8*)"!", 1, true);  // must not merge into the saved step
    CHECK(d.modified());
    u64 caret;
    CHECK(d.undo(caret));
    CHECK(!d.modified());
    CHECK(docText(d) == ref);
    CHECK(d.undo(caret));
    CHECK(d.size() == 0);
    CHECK(d.modified());
    CHECK(!d.undo(caret));
    CHECK(d.redo(caret) && docText(d) == ref && caret == ref.size());

    // Backspace run coalesces too.
    d.breakCoalesce();
    for (int i = 0; i < 10; i++) d.erase(d.size() - 1, 1, true);
    CHECK(d.size() == ref.size() - 10);
    CHECK(d.undo(caret));
    CHECK(docText(d) == ref && caret == ref.size());
}

static void testSearch() {
    std::string big;
    while (big.size() < 900000) big += randomText(200) + "Needle" + randomText(50) + "needle";
    Doc d;
    LineIndex idx;
    idx.buildSync((const u8*)big.data(), big.size());
    d.reset((const u8*)big.data(), big.size(), &idx);
    std::string ref = big;
    for (int i = 0; i < 300; i++) {  // fragment the tree
        u64 pos = rnd(ref.size() + 1);
        std::string t = (i % 3 == 0) ? std::string("nEEdle") : randomText(1 + rnd(5));
        d.insert(pos, (const u8*)t.data(), t.size());
        ref.insert((size_t)pos, t);
    }
    std::string low = ref;
    for (char& c : low)
        if (c >= 'A' && c <= 'Z') c = (char)(c + 32);
    for (int i = 0; i < 200; i++) {
        u64 from = rnd(ref.size());
        size_t f = ref.find("needle", (size_t)from);
        CHECK(d.searchFwd("needle", from, d.size(), true) == (f == std::string::npos ? NPOS : (u64)f));
        f = low.find("needle", (size_t)from);
        CHECK(d.searchFwd("NEEDLE", from, d.size(), false) == (f == std::string::npos ? NPOS : (u64)f));
        // backward: last match entirely below `from`
        size_t b = from >= 6 ? ref.rfind("Needle", (size_t)from - 6) : std::string::npos;
        CHECK(d.searchBack("Needle", 0, from, true) == (b == std::string::npos ? NPOS : (u64)b));
    }
    CHECK(d.searchFwd("this string is not present", 0, d.size(), true) == NPOS);
    CHECK(d.searchBack("this string is not present", 0, d.size(), false) == NPOS);
}

static void testLineIndex() {
    std::string s;
    while (s.size() < 700000) s += randomText(1 + rnd(300));
    LineIndex sync, async;
    sync.buildSync((const u8*)s.data(), s.size());
    bool done = false;
    async.buildAsync((const u8*)s.data(), s.size(), [&] { done = true; });
    for (int i = 0; i < 2000 && !async.ready(); i++) std::this_thread::sleep_for(std::chrono::milliseconds(1));
    async.cancel();
    CHECK(async.ready() && done);
    u64 total = refLineOf(s, s.size());
    CHECK(sync.countBefore(s.size()) == total);
    CHECK(async.countBefore(s.size()) == total);
    for (int i = 0; i < 300; i++) {
        u64 a = rnd(s.size() + 1), b = a + rnd(s.size() - a + 1);
        CHECK(sync.countRange(a, b) == refLineOf(s, b) - refLineOf(s, a));
        u64 k = 1 + rnd(50);
        u64 want = refLineStart(s, refLineOf(s, a) + k);
        CHECK(sync.findNth(a, k) == want);
    }
    // countNewlines on odd lengths and alignments
    for (u64 n = 0; n < 600; n += 7) CHECK(countNewlines((const u8*)s.data() + n % 13, n) == refLineOf(s.substr(n % 13, n), n));
}

static void testLayout(const std::string& text, int wrapCols) {
    Doc d;
    LineIndex idx;
    idx.buildSync((const u8*)text.data(), text.size());
    d.reset((const u8*)text.data(), text.size(), &idx);
    LayoutCfg cfg;
    cfg.wrapCols = wrapCols;
    // Walk forward collecting rows; the visible bytes plus line endings must be the document.
    std::vector<Row> rows;
    Row r = rowAt(d, 0, cfg);
    rows.push_back(r);
    Row nx;
    while (nextRow(d, r, cfg, nx)) {
        CHECK(nx.start == r.next);
        CHECK(nx.start > r.start || nx.start == d.size());
        r = nx;
        rows.push_back(r);
        if (rows.size() > text.size() + 10) break;
    }
    std::string rebuilt;
    for (const Row& x : rows) {
        CHECK(x.start <= x.end && x.end <= x.next);
        CHECK(x.end - x.start <= kMaxRowBytes);
        rebuilt += text.substr((size_t)x.start, (size_t)(x.next - x.start));
        if (wrapCols > 0) CHECK(colOf(d, x, x.end, cfg) <= std::max(wrapCols, 4));
        std::string ending = text.substr((size_t)x.end, (size_t)(x.next - x.end));
        CHECK(ending == "" || ending == "\n" || ending == "\r\n");
    }
    CHECK(rebuilt == text);
    // Walking backward visits the same rows.
    for (size_t i = rows.size(); i-- > 1;) {
        Row pv;
        CHECK(prevRow(d, rows[i].start, cfg, pv));
        CHECK(pv.start == rows[i - 1].start && pv.next == rows[i - 1].next);
    }
    Row pv;
    CHECK(!prevRow(d, 0, cfg, pv));
    // Every offset maps to the row that holds it, and columns round-trip.
    for (int i = 0; i < 400; i++) {
        u64 pos = rnd(text.size() + 1);
        if (pos < text.size() && pos > 0 && text[(size_t)pos] == '\n' && text[(size_t)pos - 1] == '\r') continue;
        Row c = rowContaining(d, pos, cfg);
        CHECK(rowHas(d, c, pos));
        bool found = false;
        for (const Row& x : rows)
            if (x.start == c.start) found = true;
        CHECK(found);
        // only test the round trip on character boundaries
        if (pos <= c.end && (pos == text.size() || ((u8)text[(size_t)pos] & 0xC0) != 0x80)) {
            int col = colOf(d, c, pos, cfg);
            u64 back = posOfCol(d, c, col, cfg);
            CHECK(colOf(d, c, back, cfg) == col);
        }
    }
    // Caret stepping covers the document in both directions without getting stuck.
    u64 p = 0, steps = 0;
    while (p < d.size() && steps < text.size() + 5) {
        u64 q = nextChar(d, p, cfg);
        CHECK(q > p);
        CHECK(prevChar(d, q, cfg) == p);
        p = q;
        steps++;
    }
    CHECK(p == d.size());
    p = 0;
    steps = 0;
    while (p < d.size() && steps < text.size() + 5) {
        u64 q = wordRight(d, p, cfg);
        CHECK(q > p);
        p = q;
        steps++;
    }
    p = d.size();
    while (p > 0) {
        u64 q = wordLeft(d, p, cfg);
        CHECK(q < p);
        p = q;
    }
}

static void testUtf() {
    std::string s = randomText(5000) + "\xFF\xC3" + "tail\xE4\xB8";
    // Reference: feed everything at once.
    std::vector<u16> all;
    {
        Utf8ToUtf16 c;
        c.feed((const u8*)s.data(), s.size(), all);
        c.finish(all);
    }
    // Feed in random slices; the result must be identical.
    for (int round = 0; round < 50; round++) {
        std::vector<u16> out;
        Utf8ToUtf16 c;
        size_t i = 0;
        while (i < s.size()) {
            size_t n = std::min<size_t>(1 + (size_t)rnd(9), s.size() - i);
            c.feed((const u8*)s.data() + i, n, out);
            i += n;
        }
        c.finish(out);
        CHECK(out == all);
    }
    // UTF-16 -> UTF-8 -> UTF-16 round trip of valid text (both byte orders).
    std::string valid = randomText(3000);
    std::vector<u16> w;
    {
        Utf8ToUtf16 c;
        c.feed((const u8*)valid.data(), valid.size(), w);
        c.finish(w);
    }
    for (int be = 0; be < 2; be++) {
        std::vector<u8> bytes;
        for (u16 u : w) {
            if (be) { bytes.push_back((u8)(u >> 8)); bytes.push_back((u8)u); }
            else { bytes.push_back((u8)u); bytes.push_back((u8)(u >> 8)); }
        }
        std::vector<u8> back;
        utf16ToUtf8(bytes.data(), bytes.size(), be != 0, back);
        CHECK(std::string(back.begin(), back.end()) == valid);
    }
    CHECK(looksUtf8((const u8*)valid.data(), valid.size()));
    CHECK(looksUtf8((const u8*)"abc\xE4\xB8", 5));   // cut off at the end is fine
    CHECK(!looksUtf8((const u8*)"caf\xE9 au lait", 12));  // Latin-1
    u32 cp;
    CHECK(utf8Decode((const u8*)"\xC0\x80", 2, cp) == 1 && cp == 0xFFFD);       // overlong
    CHECK(utf8Decode((const u8*)"\xED\xA0\x80", 3, cp) == 1 && cp == 0xFFFD);   // surrogate
    CHECK(utf8Decode((const u8*)"\xF0\x9F\x98\x80", 4, cp) == 4 && cp == 0x1F600);
    CHECK(cellWidth('a') == 1 && cellWidth(0x4E2D) == 2 && cellWidth(0x301) == 0 && cellWidth(0x1F600) == 2);
    std::string e = "a\nb\r\nc\rd";
    normalizeEol(e, true);
    CHECK(e == "a\r\nb\r\nc\r\nd");
    normalizeEol(e, false);
    CHECK(e == "a\nb\nc\nd");
}


static void testDiff() {
    auto hashOf = [](const std::string& s) {
        LineHasher h;
        // feed in two parts to exercise state carried across pieces
        size_t cut = s.size() / 2;
        h.feed((const u8*)s.data(), cut);
        h.feed((const u8*)s.data() + cut, s.size() - cut);
        h.finish();
        return h.lines;
    };
    CHECK(hashOf("").empty());
    CHECK(hashOf("a\n").size() == 1 && hashOf("a\nb").size() == 2 && hashOf("\n\n").size() == 2);
    CHECK(hashOf("a\r\nb\r\n") == hashOf("a\nb\n"));   // line endings do not matter
    CHECK(hashOf("a\rb\n") != hashOf("ab\n"));            // a lone CR does
    CHECK(hashOf("x\ny") == hashOf("x\ny\n"));
    for (int round = 0; round < 400; round++) {
        // Lines drawn from a small alphabet, so there are many repeats to confuse the matcher.
        std::vector<u64> a, b;
        size_t n = (size_t)rnd(round < 300 ? 60 : 3000);
        for (size_t i = 0; i < n; i++) a.push_back(rnd(round % 2 ? 6 : 1000));
        b = a;
        int edits = (int)rnd(12);
        for (int e = 0; e < edits; e++) {
            size_t pos = (size_t)rnd(b.size() + 1);
            if (rnd(2) && pos < b.size()) b.erase(b.begin() + (long)pos, b.begin() + (long)std::min(b.size(), pos + 1 + (size_t)rnd(4)));
            else b.insert(b.begin() + (long)pos, 1 + (size_t)rnd(4), 5000 + rnd(3));
        }
        std::vector<Hunk> hs = diffLines(a, b);
        // Rebuild b from a and the hunks; everything outside hunks must be equal lines.
        std::vector<u64> rebuilt;
        size_t ai = 0, bi = 0;
        bool ok = true;
        for (const Hunk& h : hs) {
            if (h.aStart < ai || h.bStart < bi || h.aStart - ai != h.bStart - bi) { ok = false; break; }
            if (h.aCount == 0 && h.bCount == 0) ok = false;
            while (ai < h.aStart) { if (a[ai] != b[bi]) ok = false; rebuilt.push_back(a[ai]); ai++; bi++; }
            for (u32 k = 0; k < h.bCount; k++) rebuilt.push_back(b[bi + k]);
            ai += h.aCount;
            bi += h.bCount;
        }
        while (ok && ai < a.size() && bi < b.size()) { if (a[ai] != b[bi]) ok = false; rebuilt.push_back(a[ai]); ai++; bi++; }
        CHECK(ok && ai == a.size() && bi == b.size());
        CHECK(rebuilt == b);
        if (a == b) CHECK(hs.empty());
        // Lookups agree with a linear scan.
        for (int q = 0; q < 20; q++) {
            u32 line = (u32)rnd(b.size() + 2);
            int want = -1;
            for (size_t k = 0; k < hs.size(); k++)
                if ((hs[k].bCount == 0 && hs[k].bStart == line) || (line >= hs[k].bStart && line < hs[k].bStart + hs[k].bCount)) want = (int)k;
            CHECK(hunkAtB(hs, line) == want);
            if (want < 0 && line < b.size()) {
                i64 m = mapBtoA(hs, line);
                CHECK(m >= 0 && (size_t)m < a.size() && a[(size_t)m] == b[line]);
            }
        }
    }
    // Two unrelated large inputs: must finish quickly and still be a valid (coarse) diff.
    std::vector<u64> a, b;
    for (int i = 0; i < 200000; i++) { a.push_back(rng()); b.push_back(rng()); }
    std::vector<Hunk> hs = diffLines(a, b);
    CHECK(hs.size() == 1 && hs[0].aCount == a.size() && hs[0].bCount == b.size());
}

static void bench() {
    // Not a pass/fail test: prints rough numbers so regressions are visible.
    using clk = std::chrono::steady_clock;
    auto ms = [](clk::duration d) { return std::chrono::duration<double, std::milli>(d).count(); };
    const u64 N = 512ull << 20;
    std::vector<u8> big((size_t)N);
    for (u64 i = 0; i < N; i++) big[(size_t)i] = (i % 61 == 60) ? '\n' : (u8)('a' + i % 26);

    auto t0 = clk::now();
    Doc d;
    LineIndex idx;
    d.reset(big.data(), N, &idx);
    auto t1 = clk::now();
    printf("  open 512 MB (no index yet):        %8.4f ms\n", ms(t1 - t0));

    LayoutCfg cfg;
    t0 = clk::now();
    u64 sink = 0;
    for (int i = 0; i < 1000; i++) {
        Row r = rowAt(d, lineStartBefore(d, rnd(N), cfg), cfg);
        for (int k = 0; k < 60; k++) {
            sink += r.end;
            Row nx;
            if (!nextRow(d, r, cfg, nx)) break;
            r = nx;
        }
    }
    t1 = clk::now();
    printf("  lay out a 60-row screen anywhere:  %8.4f ms\n", ms(t1 - t0) / 1000);

    t0 = clk::now();
    const int kEdits = 100000;
    for (int i = 0; i < kEdits; i++) d.insert(rnd(d.size()), (const u8*)"x", 1);
    t1 = clk::now();
    printf("  random insert into 512 MB:         %8.4f ms each (%d edits, %zu pieces)\n",
           ms(t1 - t0) / kEdits, kEdits, d.pieceCount());

    t0 = clk::now();
    idx.buildSync(big.data(), N);
    d.indexReady();
    t1 = clk::now();
    printf("  line index for 512 MB:             %8.1f ms (%.1f GB/s), %llu lines\n", ms(t1 - t0),
           (double)N / (1 << 30) / (ms(t1 - t0) / 1000), (unsigned long long)d.newlineCount() + 1);

    t0 = clk::now();
    for (int i = 0; i < 10000; i++) sink += d.lineStart(rnd(d.newlineCount()));
    t1 = clk::now();
    printf("  go to a random line:               %8.4f ms\n", ms(t1 - t0) / 10000);

    t0 = clk::now();
    for (int i = 0; i < 10000; i++) sink += d.lineOf(rnd(d.size()));
    t1 = clk::now();
    printf("  line number of a random offset:    %8.4f ms\n", ms(t1 - t0) / 10000);

    t0 = clk::now();
    u64 hit = d.searchFwd("not in the file at all", 0, d.size(), true);
    t1 = clk::now();
    printf("  search all 512 MB (%zu pieces):  %8.1f ms%s\n", d.pieceCount(), ms(t1 - t0),
           hit == NPOS ? "" : " ?");
    if (sink == 42) printf(" ");
}

int main(int argc, char** argv) {
    testUtf();
    testLineIndex();
    testTyping();
    testEdits("", false, 3000);
    testEdits(randomText(2000), false, 6000);
    testEdits(randomText(300000), false, 3000);  // spans several index blocks
    testEdits(randomText(300000), true, 3000);   // edits while the index is still missing
    testSearch();
    testDiff();
    for (int wrap : {0, 7, 40}) {
        testLayout("", wrap);
        testLayout("one line, no newline", wrap);
        testLayout("ends with newline\n", wrap);
        testLayout("\n\n\r\n\n", wrap);
        testLayout(randomText(6000), wrap);
        testLayout(std::string(40000, 'x') + "\n" + randomText(500) + std::string(33000, 'y'), wrap);
        testLayout(randomText(100) + std::string(20000, ' ') + "\r\n" + std::string(16383, 'q') + "\r\nend", wrap);
    }
    printf("%ld checks, %d failures\n", g_checks, g_fail);
    if (argc > 1 && std::string(argv[1]) == "bench") bench();
    return g_fail ? 1 : 0;
}
