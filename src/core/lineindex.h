#pragma once
// Newline index over the immutable original file contents.
//
// The file is cut into 64 KB blocks and we store a running newline count per
// block. That costs 8 bytes per 64 KB (about 1.3 MB for a 10 GB file) and
// answers "how many newlines before offset X" and "where is newline number N"
// with one array lookup plus a scan of at most one block.
//
// The index can be built on a background thread; the editor never waits for it.
#include <atomic>
#include <functional>
#include <thread>
#include <vector>

#include "types.h"

namespace jnp {

u64 countNewlines(const u8* p, u64 n);

class LineIndex {
public:
    static constexpr u64 kShift = 16;
    static constexpr u64 kBlock = 1ull << kShift;

    LineIndex() = default;
    ~LineIndex() { cancel(); }
    LineIndex(const LineIndex&) = delete;
    LineIndex& operator=(const LineIndex&) = delete;

    void buildSync(const u8* data, u64 size);
    // onDone runs on the worker thread once the index is complete (not when cancelled).
    void buildAsync(const u8* data, u64 size, std::function<void()> onDone);
    void cancel();

    bool ready() const { return ready_.load(std::memory_order_acquire); }
    u64 progressBytes() const { return progress_.load(std::memory_order_relaxed); }
    u64 size() const { return n_; }

    // The following require ready().
    u64 countRange(u64 a, u64 b) const;  // newlines in [a, b)
    u64 countBefore(u64 x) const;        // newlines in [0, x)
    // Offset just past the k-th newline (k >= 1) at or after `from`; size() if none.
    u64 findNth(u64 from, u64 k) const;

private:
    void run();

    const u8* d_ = nullptr;
    u64 n_ = 0;
    std::vector<u64> prefix_;  // prefix_[i] = newlines in blocks [0, i)
    std::atomic<bool> ready_{false};
    std::atomic<bool> stop_{false};
    std::atomic<u64> progress_{0};
    std::thread th_;
};

}  // namespace jnp
