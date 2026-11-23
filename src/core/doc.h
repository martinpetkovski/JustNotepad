#pragma once
// The document: a piece tree.
//
// The text is a sequence of "pieces", each pointing at a run of bytes that is
// never modified: either part of the original file (memory-mapped or loaded),
// or part of an append-only buffer that receives everything the user types.
// The pieces live in a balanced tree (a treap) keyed by position, so inserting
// or deleting anywhere costs O(log pieces) regardless of the file size, and
// opening a file copies nothing.
//
// Undo and redo just move piece descriptors around, so they are cheap too.
#include <functional>
#include <memory>
#include <string>
#include <vector>

#include "lineindex.h"
#include "types.h"

namespace jnp {

struct Piece {
    u32 src;  // 0 = original file, otherwise add-buffer chunk number + 1
    u64 off;
    u64 len;
    i64 nl;  // newlines inside, or -1 while the line index is still being built
};

class Doc {
public:
    Doc();
    ~Doc();
    Doc(const Doc&) = delete;
    Doc& operator=(const Doc&) = delete;

    // Replaces the whole document with `size` bytes at `orig` (not copied; must
    // stay valid). `idx` may still be building; call indexReady() when it is done.
    void reset(const u8* orig, u64 size, const LineIndex* idx);
    // Same content at a new address (the file was re-mapped).
    void rebaseOriginal(const u8* orig, const LineIndex* idx);
    void indexReady();

    u64 size() const;
    u64 version() const { return version_; }

    // ---- reading ----
    // Contiguous bytes starting at pos. Returns false at or past the end.
    bool spanAt(u64 pos, const u8*& p, u64& n) const;
    u64 copy(u64 pos, u8* dst, u64 n) const;
    u8 at(u64 pos) const;
    std::string text(u64 pos, u64 n) const;
    // First `c` in [from, limit), or NPOS.
    u64 findFwd(u8 c, u64 from, u64 limit) const;
    // Last `c` in [limit, from), or NPOS.
    u64 findBack(u8 c, u64 from, u64 limit) const;
    // Calls f(ptr, len) for every piece in order; stop by returning false.
    void forEachPiece(const std::function<bool(const u8*, u64)>& f) const;
    size_t pieceCount() const;

    // ---- lines (0-based) ----
    bool linesKnown() const;
    u64 newlineCount() const;       // requires linesKnown()
    u64 lineOf(u64 pos) const;      // newlines before pos
    u64 lineStart(u64 line) const;  // offset where that line begins (size() if past the end)

    // ---- editing ----
    // `coalesce` lets consecutive typing or deleting merge into one undo step.
    void insert(u64 pos, const u8* s, u64 n, bool coalesce = false);
    void erase(u64 pos, u64 n, bool coalesce = false);
    void beginGroup();  // edits until endGroup() undo as one step
    void endGroup();
    void breakCoalesce() { coalesceOk_ = false; }

    bool canUndo() const { return !undo_.empty(); }
    bool canRedo() const { return !redo_.empty(); }
    bool undo(u64& caret);
    bool redo(u64& caret);

    bool modified() const;
    void markSaved();

    // ---- search (bytes; case-insensitivity is ASCII only) ----
    // `progress` is called now and then with the current offset; return false to abort.
    using Progress = std::function<bool(u64)>;
    u64 searchFwd(const std::string& needle, u64 from, u64 to, bool matchCase,
                  const Progress& progress = nullptr) const;
    // Last match lying entirely inside [lo, hi).
    u64 searchBack(const std::string& needle, u64 lo, u64 hi, bool matchCase,
                   const Progress& progress = nullptr) const;

private:
    struct Node;
    struct Edit {
        u64 pos = 0;
        std::vector<Piece> removed;
        u64 removedLen = 0;
        std::vector<Piece> inserted;
        u64 insertedLen = 0;
        u32 group = 0;
        u64 id = 0;
    };
    struct Chunk {
        std::unique_ptr<u8[]> data;
        u64 cap = 0, fill = 0;
    };

    Node* newNode(const Piece& p);
    void freeNode(Node* n);
    void freeTree(Node* t);
    static void pull(Node* t);
    void split(Node* t, u64 pos, Node*& a, Node*& b);
    Node* merge(Node* a, Node* b);
    void collect(Node* t, std::vector<Piece>* out) const;
    void fixNl(Node* t);

    const u8* dataOf(const Piece& p) const;
    i64 pieceNl(const Piece& p) const;
    u64 nlInPrefix(const Piece& p, u64 k) const;
    u64 afterNthNl(const Piece& p, u64 k) const;
    Piece appendAdd(const u8* s, u64 n);
    void rawInsert(u64 pos, const Piece* ps, size_t count);
    void rawErase(u64 pos, u64 n, std::vector<Piece>* out);
    bool locate(u64 pos, const u8*& ptr, u64& start, u64& len) const;
    void pushEdit(Edit&& e);

    Node* root_ = nullptr;
    Node* free_ = nullptr;
    std::vector<std::unique_ptr<Node[]>> pools_;
    size_t poolUsed_ = 0;
    u32 rng_ = 0x9E3779B9u;

    const u8* orig_ = nullptr;
    const LineIndex* idx_ = nullptr;
    std::vector<Chunk> chunks_;

    std::vector<Edit> undo_, redo_;
    u32 group_ = 0, groupSeq_ = 0;
    u64 idSeq_ = 0, savedId_ = 0;
    bool coalesceOk_ = false;
    u64 version_ = 0;

    // One-entry cache of the last piece looked up; most reads are sequential.
    mutable u64 cVer_ = NPOS, cStart_ = 0, cLen_ = 0;
    mutable const u8* cPtr_ = nullptr;
};

}  // namespace jnp
