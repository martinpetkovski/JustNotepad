#pragma once
// Line diff between the base revision (a) and the text being edited (b).
//
// Lines are compared by a 64-bit hash of their content without the line
// ending, so a CRLF working copy still matches an LF copy in the repository.
#include <vector>

#include "doc.h"
#include "types.h"

namespace jnp {

struct Hunk {
    u32 aStart, aCount;  // lines removed from the base
    u32 bStart, bCount;  // lines added in the current text
};

// Splits a byte stream into lines and hashes each one. Feed pieces in order.
class LineHasher {
public:
    void feed(const u8* p, u64 n);
    void finish();
    std::vector<u64> lines;

private:
    u64 h_ = 1469598103934665603ull;
    bool any_ = false, cr_ = false;
};

void hashDocLines(const Doc& d, std::vector<u64>& out);

// Myers diff. Returns hunks in order. If the two sides differ too much for the
// work budget, the differing middle is reported as one big hunk.
std::vector<Hunk> diffLines(const std::vector<u64>& a, const std::vector<u64>& b);

// Index of the hunk covering current line `line` (or a pure deletion sitting
// right at that line), or -1.
int hunkAtB(const std::vector<Hunk>& hunks, u32 line);
// Index of the hunk covering base line `line`, or -1.
int hunkAtA(const std::vector<Hunk>& hunks, u32 line);
// Base line that corresponds to current line `line`.
i64 mapBtoA(const std::vector<Hunk>& hunks, i64 line);

}  // namespace jnp
