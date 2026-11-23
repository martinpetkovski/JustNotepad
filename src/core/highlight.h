#pragma once
// Syntax highlighting: one small tokenizer driven by a table of languages.
//
// It colours one line at a time. What carries over from one line to the next
// (inside a block comment, a multi-line string, a markup tag, ...) is packed
// into a 32-bit state, so the editor can keep the state at the start of every
// few hundred lines and colour any visible line without reading the file from
// the top. No OS headers: this is part of the portable core.
#include <cstddef>
#include <string>

#include "types.h"

namespace jnp {

enum HlKind : u8 {
    HL_TEXT = 0,
    HL_KEYWORD,
    HL_TYPE,
    HL_LITERAL,   // true, false, null, constants
    HL_FUNCTION,  // a name followed by "("
    HL_STRING,
    HL_NUMBER,
    HL_COMMENT,
    HL_PREPROC,   // #include, annotations, Unreal macros
    HL_TAG,       // markup tags, headings
    HL_ATTR,      // markup attributes, keys in JSON / YAML / INI
    HL_INSERTED,  // diff "+" lines
    HL_DELETED,   // diff "-" lines
    HL_COUNT
};

struct Lang;

// Picks the language for a file name (extension or well-known names such as
// "Makefile"), falling back to a "#!" first line. nullptr: plain text.
const Lang* langForFile(const std::string& fileNameUtf8, const std::string& firstLine);
const char* langName(const Lang* l);

// Colours `n` bytes of one line (without its line break). `state` is the
// state at the start of the line; the state at its end is returned.
// `kinds`, if not null, receives one HlKind per byte.
u32 highlightLine(const Lang* l, const u8* p, size_t n, u32 state, u8* kinds);

}  // namespace jnp
