#include "lineindex.h"

#include <algorithm>
#include <cstring>

#if defined(__SSE2__) || defined(_M_X64) || defined(_M_AMD64)
#include <emmintrin.h>
#define JN_SSE2 1
#endif

namespace jnp {

u64 countNewlines(const u8* p, u64 n) {
    u64 c = 0;
#ifdef JN_SSE2
    const __m128i nl = _mm_set1_epi8('\n');
    const __m128i zero = _mm_setzero_si128();
    while (n >= 16) {
        // Up to 255 iterations fit in the per-byte counters before they overflow.
        u64 iters = n / 16;
        if (iters > 255) iters = 255;
        __m128i acc = zero;
        for (u64 i = 0; i < iters; i++) {
            acc = _mm_sub_epi8(acc, _mm_cmpeq_epi8(_mm_loadu_si128((const __m128i*)p), nl));
            p += 16;
        }
        n -= iters * 16;
        __m128i s = _mm_sad_epu8(acc, zero);
        c += (u64)_mm_cvtsi128_si64(s) + (u64)_mm_cvtsi128_si64(_mm_srli_si128(s, 8));
    }
#endif
    for (u64 i = 0; i < n; i++) c += (p[i] == '\n');
    return c;
}

void LineIndex::run() {
    u64 blocks = (n_ + kBlock - 1) >> kShift;
    for (u64 i = 0; i < blocks; i++) {
        if (stop_.load(std::memory_order_relaxed)) return;
        u64 off = i << kShift;
        u64 len = std::min(kBlock, n_ - off);
        prefix_[i + 1] = prefix_[i] + countNewlines(d_ + off, len);
        progress_.store(off + len, std::memory_order_relaxed);
    }
    ready_.store(true, std::memory_order_release);
}

void LineIndex::buildSync(const u8* data, u64 size) {
    cancel();
    d_ = data;
    n_ = size;
    prefix_.assign(((n_ + kBlock - 1) >> kShift) + 1, 0);
    stop_ = false;
    ready_ = false;
    progress_ = 0;
    run();
}

void LineIndex::buildAsync(const u8* data, u64 size, std::function<void()> onDone) {
    cancel();
    d_ = data;
    n_ = size;
    prefix_.assign(((n_ + kBlock - 1) >> kShift) + 1, 0);
    stop_ = false;
    ready_ = false;
    progress_ = 0;
    th_ = std::thread([this, onDone]() {
        run();
        if (ready() && onDone) onDone();
    });
}

void LineIndex::cancel() {
    stop_ = true;
    if (th_.joinable()) th_.join();
}

u64 LineIndex::countBefore(u64 x) const {
    if (x > n_) x = n_;
    u64 b = x >> kShift;
    u64 start = b << kShift;
    return prefix_[b] + countNewlines(d_ + start, x - start);
}

u64 LineIndex::countRange(u64 a, u64 b) const {
    if (b > n_) b = n_;
    if (a >= b) return 0;
    if ((a >> kShift) == ((b - 1) >> kShift)) return countNewlines(d_ + a, b - a);
    return countBefore(b) - countBefore(a);
}

u64 LineIndex::findNth(u64 from, u64 k) const {
    if (k == 0) return from;
    u64 target = countBefore(from) + k;  // 1-based rank of the newline in the file
    if (target > prefix_.back()) return n_;
    // First block boundary whose prefix reaches the target, minus one.
    size_t i = (size_t)(std::lower_bound(prefix_.begin(), prefix_.end(), target) - prefix_.begin());
    u64 b = (u64)i - 1;
    u64 need = target - prefix_[b];
    u64 pos = b << kShift;
    u64 end = std::min(n_, pos + kBlock);
    while (pos < end) {
        const void* q = memchr(d_ + pos, '\n', (size_t)(end - pos));
        if (!q) break;
        pos = (u64)((const u8*)q - d_) + 1;
        if (--need == 0) return pos;
    }
    return n_;
}

}  // namespace jnp
