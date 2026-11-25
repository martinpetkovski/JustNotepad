#include "doc.h"

#include <algorithm>
#include <cstring>

namespace jnp {

struct Doc::Node {
    Piece p;
    u64 sumLen;  // bytes in this subtree
    u64 sumNl;   // known newlines in this subtree
    u32 prio;
    bool unk;  // some piece in this subtree has an unknown newline count
    Node* l;
    Node* r;
};

#define JN_LEN(t) ((t) ? (t)->sumLen : 0)
#define JN_NL(t) ((t) ? (t)->sumNl : 0)
#define JN_UNK(t) ((t) ? (t)->unk : false)

Doc::Doc() = default;

Doc::~Doc() = default;

// ---------------------------------------------------------------- nodes

Doc::Node* Doc::newNode(const Piece& p) {
    Node* n;
    if (free_) {
        n = free_;
        free_ = n->l;
    } else {
        const size_t kPool = 1024;
        if (pools_.empty() || poolUsed_ == kPool) {
            pools_.emplace_back(new Node[kPool]);
            poolUsed_ = 0;
        }
        n = &pools_.back()[poolUsed_++];
    }
    rng_ ^= rng_ << 13;
    rng_ ^= rng_ >> 17;
    rng_ ^= rng_ << 5;
    n->p = p;
    n->prio = rng_;
    n->l = n->r = nullptr;
    pull(n);
    return n;
}

void Doc::freeNode(Node* n) {
    n->l = free_;
    free_ = n;
}

void Doc::freeTree(Node* t) {
    if (!t) return;
    freeTree(t->l);
    freeTree(t->r);
    freeNode(t);
}

void Doc::pull(Node* t) {
    t->sumLen = t->p.len + JN_LEN(t->l) + JN_LEN(t->r);
    t->sumNl = (t->p.nl < 0 ? 0 : (u64)t->p.nl) + JN_NL(t->l) + JN_NL(t->r);
    t->unk = t->p.nl < 0 || JN_UNK(t->l) || JN_UNK(t->r);
}

const u8* Doc::dataOf(const Piece& p) const {
    return p.src == 0 ? orig_ + p.off : chunks_[p.src - 1].data.get() + p.off;
}

i64 Doc::pieceNl(const Piece& p) const {
    if (p.src == 0) {
        if (!idx_ || !idx_->ready()) return -1;
        return (i64)idx_->countRange(p.off, p.off + p.len);
    }
    return (i64)countNewlines(dataOf(p), p.len);
}

u64 Doc::nlInPrefix(const Piece& p, u64 k) const {
    if (k == 0) return 0;
    if (p.src == 0) return idx_->countRange(p.off, p.off + k);
    return countNewlines(dataOf(p), k);
}

u64 Doc::afterNthNl(const Piece& p, u64 k) const {
    if (p.src == 0) return idx_->findNth(p.off, k) - p.off;
    const u8* d = dataOf(p);
    u64 pos = 0;
    while (pos < p.len) {
        const void* q = memchr(d + pos, '\n', (size_t)(p.len - pos));
        if (!q) break;
        pos = (u64)((const u8*)q - d) + 1;
        if (--k == 0) return pos;
    }
    return p.len;
}

void Doc::split(Node* t, u64 pos, Node*& a, Node*& b) {
    if (!t) {
        a = b = nullptr;
        return;
    }
    u64 ls = JN_LEN(t->l);
    if (pos <= ls) {
        Node* tmp;
        split(t->l, pos, a, tmp);
        t->l = tmp;
        pull(t);
        b = t;
    } else if (pos >= ls + t->p.len) {
        Node* tmp;
        split(t->r, pos - ls - t->p.len, tmp, b);
        t->r = tmp;
        pull(t);
        a = t;
    } else {
        // The split point is inside this piece: cut it in two.
        u64 k = pos - ls;
        Piece pa = t->p, pb = t->p;
        pa.len = k;
        pb.off += k;
        pb.len -= k;
        if (t->p.nl < 0) {
            pa.nl = pieceNl(pa);
            pb.nl = pieceNl(pb);
        } else if (pa.len <= pb.len) {  // count the shorter half, derive the other
            pa.nl = pieceNl(pa);
            if (pa.nl < 0) pb.nl = -1;
            else pb.nl = t->p.nl - pa.nl;
        } else {
            pb.nl = pieceNl(pb);
            if (pb.nl < 0) pa.nl = -1;
            else pa.nl = t->p.nl - pb.nl;
        }
        Node* right = t->r;
        t->r = nullptr;
        t->p = pa;
        pull(t);
        a = t;
        b = merge(newNode(pb), right);
    }
}

Doc::Node* Doc::merge(Node* a, Node* b) {
    if (!a) return b;
    if (!b) return a;
    if (a->prio >= b->prio) {
        a->r = merge(a->r, b);
        pull(a);
        return a;
    }
    b->l = merge(a, b->l);
    pull(b);
    return b;
}

void Doc::collect(Node* t, std::vector<Piece>* out) const {
    if (!t) return;
    collect(t->l, out);
    out->push_back(t->p);
    collect(t->r, out);
}

void Doc::fixNl(Node* t) {
    if (!t || !t->unk) return;
    fixNl(t->l);
    fixNl(t->r);
    if (t->p.nl < 0) t->p.nl = pieceNl(t->p);
    pull(t);
}

// ---------------------------------------------------------------- setup

void Doc::reset(const u8* orig, u64 size, const LineIndex* idx) {
    freeTree(root_);
    root_ = nullptr;
    orig_ = orig;
    idx_ = idx;
    chunks_.clear();
    undo_.clear();
    redo_.clear();
    group_ = 0;
    savedId_ = 0;
    coalesceOk_ = false;
    version_++;
    if (size) {
        Piece p{0, 0, size, 0};
        p.nl = pieceNl(p);
        root_ = newNode(p);
    }
}

void Doc::rebaseOriginal(const u8* orig, const LineIndex* idx) {
    orig_ = orig;
    idx_ = idx;
    version_++;
}

void Doc::indexReady() {
    fixNl(root_);
    version_++;
}

u64 Doc::size() const { return JN_LEN(root_); }

// ---------------------------------------------------------------- reading

bool Doc::locate(u64 pos, const u8*& ptr, u64& start, u64& len) const {
    if (cVer_ == version_ && pos >= cStart_ && pos < cStart_ + cLen_) {
        ptr = cPtr_;
        start = cStart_;
        len = cLen_;
        return true;
    }
    Node* t = root_;
    u64 base = 0;
    while (t) {
        u64 ls = JN_LEN(t->l);
        if (pos < base + ls) {
            t = t->l;
        } else if (pos < base + ls + t->p.len) {
            start = cStart_ = base + ls;
            len = cLen_ = t->p.len;
            ptr = cPtr_ = dataOf(t->p);
            cVer_ = version_;
            return true;
        } else {
            base += ls + t->p.len;
            t = t->r;
        }
    }
    return false;
}

bool Doc::spanAt(u64 pos, const u8*& p, u64& n) const {
    const u8* ptr;
    u64 start, len;
    if (!locate(pos, ptr, start, len)) return false;
    p = ptr + (pos - start);
    n = len - (pos - start);
    return true;
}

u64 Doc::copy(u64 pos, u8* dst, u64 n) const {
    u64 done = 0;
    while (done < n) {
        const u8* p;
        u64 avail;
        if (!spanAt(pos + done, p, avail)) break;
        u64 k = std::min(avail, n - done);
        memcpy(dst + done, p, (size_t)k);
        done += k;
    }
    return done;
}

u8 Doc::at(u64 pos) const {
    const u8* p;
    u64 n;
    return spanAt(pos, p, n) ? *p : 0;
}

std::string Doc::text(u64 pos, u64 n) const {
    u64 sz = size();
    if (pos > sz) pos = sz;
    if (n > sz - pos) n = sz - pos;
    std::string s((size_t)n, '\0');
    if (n) copy(pos, (u8*)&s[0], n);
    return s;
}

u64 Doc::findFwd(u8 c, u64 from, u64 limit) const {
    u64 sz = size();
    if (limit > sz) limit = sz;
    while (from < limit) {
        const u8* p;
        u64 n;
        if (!spanAt(from, p, n)) break;
        n = std::min(n, limit - from);
        const void* q = memchr(p, c, (size_t)n);
        if (q) return from + (u64)((const u8*)q - p);
        from += n;
    }
    return NPOS;
}

u64 Doc::findBack(u8 c, u64 from, u64 limit) const {
    u64 sz = size();
    if (from > sz) from = sz;
    while (from > limit) {
        const u8* ptr;
        u64 start, len;
        if (!locate(from - 1, ptr, start, len)) break;
        u64 lo = std::max(start, limit);
        for (u64 i = from; i > lo; i--)
            if (ptr[i - 1 - start] == c) return i - 1;
        from = lo;
    }
    return NPOS;
}

void Doc::forEachPiece(const std::function<bool(const u8*, u64)>& f) const {
    // Iterative in-order walk.
    std::vector<Node*> st;
    Node* t = root_;
    while (t || !st.empty()) {
        while (t) {
            st.push_back(t);
            t = t->l;
        }
        t = st.back();
        st.pop_back();
        if (!f(dataOf(t->p), t->p.len)) return;
        t = t->r;
    }
}

size_t Doc::pieceCount() const {
    size_t n = 0;
    forEachPiece([&](const u8*, u64) {
        n++;
        return true;
    });
    return n;
}

// ---------------------------------------------------------------- lines

bool Doc::linesKnown() const { return !JN_UNK(root_); }

u64 Doc::newlineCount() const { return JN_NL(root_); }

u64 Doc::lineOf(u64 pos) const {
    u64 acc = 0;
    Node* t = root_;
    while (t) {
        u64 ls = JN_LEN(t->l);
        if (pos < ls) {
            t = t->l;
            continue;
        }
        acc += JN_NL(t->l);
        pos -= ls;
        if (pos < t->p.len) {
            acc += nlInPrefix(t->p, pos);
            break;
        }
        acc += (u64)t->p.nl;
        pos -= t->p.len;
        t = t->r;
    }
    return acc;
}

u64 Doc::lineStart(u64 line) const {
    if (line == 0) return 0;
    u64 k = line, base = 0;
    Node* t = root_;
    while (t) {
        u64 lnl = JN_NL(t->l);
        if (k <= lnl) {
            t = t->l;
            continue;
        }
        k -= lnl;
        base += JN_LEN(t->l);
        if (k <= (u64)t->p.nl) return base + afterNthNl(t->p, k);
        k -= (u64)t->p.nl;
        base += t->p.len;
        t = t->r;
    }
    return size();
}

// ---------------------------------------------------------------- editing

Piece Doc::appendAdd(const u8* s, u64 n) {
    if (chunks_.empty() || chunks_.back().fill + n > chunks_.back().cap) {
        Chunk c;
        c.cap = std::max<u64>(n, 1u << 20);
        c.data.reset(new u8[(size_t)c.cap]);
        chunks_.push_back(std::move(c));
    }
    Chunk& c = chunks_.back();
    memcpy(c.data.get() + c.fill, s, (size_t)n);
    Piece p{(u32)chunks_.size(), c.fill, n, 0};
    c.fill += n;
    p.nl = (i64)countNewlines(s, n);
    return p;
}

void Doc::rawInsert(u64 pos, const Piece* ps, size_t count) {
    Node *a, *b;
    split(root_, pos, a, b);
    for (size_t i = 0; i < count; i++) {
        Piece pc = ps[i];
        if (pc.len == 0) continue;
        if (pc.nl < 0) pc.nl = pieceNl(pc);
        // If this continues the piece right before it, just make that piece
        // longer. This keeps the tree tiny while typing.
        Node* last = a;
        while (last && last->r) last = last->r;
        if (last && last->p.src == pc.src && last->p.off + last->p.len == pc.off && last->p.nl >= 0 &&
            pc.nl >= 0) {
            for (Node* t = a; t; t = t->r) {
                t->sumLen += pc.len;
                t->sumNl += (u64)pc.nl;
            }
            last->p.len += pc.len;
            last->p.nl += pc.nl;
        } else {
            a = merge(a, newNode(pc));
        }
    }
    root_ = merge(a, b);
    version_++;
    if (onEdit) {
        u64 total = 0;
        for (size_t i = 0; i < count; i++) total += ps[i].len;
        if (total) onEdit(pos, 0, total);
    }
}

void Doc::rawErase(u64 pos, u64 n, std::vector<Piece>* out) {
    if (n == 0) return;
    Node *a, *bc, *m, *c;
    split(root_, pos, a, bc);
    split(bc, n, m, c);
    if (out) collect(m, out);
    freeTree(m);
    root_ = merge(a, c);
    version_++;
    if (onEdit) onEdit(pos, n, 0);
}

void Doc::pushEdit(Edit&& e) {
    e.group = group_;
    e.id = ++idSeq_;
    undo_.push_back(std::move(e));
    redo_.clear();
    coalesceOk_ = true;
}

void Doc::insert(u64 pos, const u8* s, u64 n, bool coalesce) {
    if (n == 0) return;
    if (pos > size()) pos = size();
    Piece p = appendAdd(s, n);
    rawInsert(pos, &p, 1);
    if (coalesce && coalesceOk_ && group_ == 0 && !undo_.empty()) {
        Edit& e = undo_.back();
        if (e.group == 0 && e.id != savedId_ && e.removedLen == 0 && !e.inserted.empty() &&
            e.pos + e.insertedLen == pos) {
            Piece& lastp = e.inserted.back();
            if (lastp.src == p.src && lastp.off + lastp.len == p.off) {
                lastp.len += n;
                lastp.nl += p.nl;
                e.insertedLen += n;
                redo_.clear();
                return;
            }
        }
    }
    Edit e;
    e.pos = pos;
    e.inserted.push_back(p);
    e.insertedLen = n;
    pushEdit(std::move(e));
}

void Doc::erase(u64 pos, u64 n, bool coalesce) {
    u64 sz = size();
    if (pos >= sz || n == 0) return;
    if (n > sz - pos) n = sz - pos;
    std::vector<Piece> removed;
    rawErase(pos, n, &removed);
    if (coalesce && coalesceOk_ && group_ == 0 && !undo_.empty()) {
        Edit& e = undo_.back();
        if (e.group == 0 && e.id != savedId_ && e.insertedLen == 0 && e.removedLen > 0) {
            if (pos + n == e.pos) {  // backspace
                e.removed.insert(e.removed.begin(), removed.begin(), removed.end());
                e.pos = pos;
                e.removedLen += n;
                redo_.clear();
                return;
            }
            if (pos == e.pos) {  // forward delete
                e.removed.insert(e.removed.end(), removed.begin(), removed.end());
                e.removedLen += n;
                redo_.clear();
                return;
            }
        }
    }
    Edit e;
    e.pos = pos;
    e.removed = std::move(removed);
    e.removedLen = n;
    pushEdit(std::move(e));
}

void Doc::beginGroup() {
    group_ = ++groupSeq_;
    if (group_ == 0) group_ = ++groupSeq_;
}

void Doc::endGroup() {
    group_ = 0;
    coalesceOk_ = false;
}

bool Doc::undo(u64& caret) {
    if (undo_.empty()) return false;
    u32 g = undo_.back().group;
    do {
        Edit e = std::move(undo_.back());
        undo_.pop_back();
        rawErase(e.pos, e.insertedLen, nullptr);
        rawInsert(e.pos, e.removed.data(), e.removed.size());
        caret = e.pos + e.removedLen;
        redo_.push_back(std::move(e));
    } while (g != 0 && !undo_.empty() && undo_.back().group == g);
    coalesceOk_ = false;
    return true;
}

bool Doc::redo(u64& caret) {
    if (redo_.empty()) return false;
    u32 g = redo_.back().group;
    do {
        Edit e = std::move(redo_.back());
        redo_.pop_back();
        rawErase(e.pos, e.removedLen, nullptr);
        rawInsert(e.pos, e.inserted.data(), e.inserted.size());
        caret = e.pos + e.insertedLen;
        undo_.push_back(std::move(e));
    } while (g != 0 && !redo_.empty() && redo_.back().group == g);
    coalesceOk_ = false;
    return true;
}

bool Doc::modified() const { return (undo_.empty() ? 0 : undo_.back().id) != savedId_; }

void Doc::markSaved() {
    savedId_ = undo_.empty() ? 0 : undo_.back().id;
    coalesceOk_ = false;
}

// ---------------------------------------------------------------- search

namespace {
const u64 kWin = 1u << 18;

inline void foldAscii(u8* p, u64 n) {
    for (u64 i = 0; i < n; i++)
        if (p[i] >= 'A' && p[i] <= 'Z') p[i] = (u8)(p[i] + 32);
}

u64 findIn(const u8* h, u64 n, const u8* nd, u64 m) {
    if (m == 0 || n < m) return NPOS;
    u64 last = n - m;
    u64 i = 0;
    while (i <= last) {
        const void* q = memchr(h + i, nd[0], (size_t)(last - i + 1));
        if (!q) return NPOS;
        i = (u64)((const u8*)q - h);
        if (memcmp(h + i, nd, (size_t)m) == 0) return i;
        i++;
    }
    return NPOS;
}

u64 rfindIn(const u8* h, u64 n, const u8* nd, u64 m) {
    if (m == 0 || n < m) return NPOS;
    for (u64 i = n - m + 1; i > 0; i--)
        if (h[i - 1] == nd[0] && memcmp(h + i - 1, nd, (size_t)m) == 0) return i - 1;
    return NPOS;
}
}  // namespace

u64 Doc::searchFwd(const std::string& needle, u64 from, u64 to, bool matchCase,
                   const Progress& progress) const {
    u64 m = needle.size();
    u64 sz = size();
    if (to > sz) to = sz;
    if (m == 0 || from >= to || to - from < m) return NPOS;
    std::string nd = needle;
    if (!matchCase) foldAscii((u8*)&nd[0], m);
    std::vector<u8> buf((size_t)(kWin + m));
    u64 pos = from;
    while (pos + m <= to) {
        u64 n = std::min<u64>(kWin + m, to - pos);
        copy(pos, buf.data(), n);
        if (!matchCase) foldAscii(buf.data(), n);
        u64 i = findIn(buf.data(), n, (const u8*)nd.data(), m);
        if (i != NPOS) return pos + i;
        if (pos + n >= to) break;
        pos += n - (m - 1);
        if (progress && !progress(pos)) return NPOS;
    }
    return NPOS;
}

u64 Doc::searchBack(const std::string& needle, u64 lo, u64 hi, bool matchCase,
                    const Progress& progress) const {
    u64 m = needle.size();
    u64 sz = size();
    if (hi > sz) hi = sz;
    if (m == 0 || lo >= hi || hi - lo < m) return NPOS;
    std::string nd = needle;
    if (!matchCase) foldAscii((u8*)&nd[0], m);
    std::vector<u8> buf((size_t)(kWin + m));
    u64 end = hi;
    while (end >= lo + m) {
        u64 n = std::min<u64>(kWin + m, end - lo);
        u64 start = end - n;
        copy(start, buf.data(), n);
        if (!matchCase) foldAscii(buf.data(), n);
        u64 i = rfindIn(buf.data(), n, (const u8*)nd.data(), m);
        if (i != NPOS) return start + i;
        if (start == lo) break;
        end = start + (m - 1);
        if (progress && !progress(end)) return NPOS;
    }
    return NPOS;
}

}  // namespace jnp
