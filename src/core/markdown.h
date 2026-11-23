#pragma once
// Markdown to a list of blocks with styled runs, for the preview. Covers what
// READMEs and notes use: headings, paragraphs, emphasis, code (inline and
// fenced), lists (nested, numbered, task lists), quotes, tables, rules,
// links and images (shown as their text). Each block knows its source lines,
// so the preview can follow the editor. No OS headers: part of the portable core.
#include <string>
#include <vector>

#include "types.h"

namespace jnp {

struct MdRun {
    enum Style : u8 { Bold = 1, Italic = 2, Code = 4, Strike = 8, Link = 16 };
    std::string text;
    u8 style = 0;
    std::string url;  // for links
};

struct MdBlock {
    enum Kind : u8 { Paragraph, Heading, CodeBlock, ListItem, Rule, Table };
    Kind kind = Paragraph;
    int level = 0;        // heading 1..6; list nesting 0..
    int quote = 0;        // inside how many ">" quotes
    bool ordered = false; // list: "1." rather than "-"
    int number = 0;       // list: its number
    int task = -1;        // list: -1 not a task, 0 "[ ]", 1 "[x]"
    std::vector<MdRun> runs;       // paragraph, heading, list item
    std::vector<std::string> code; // code block lines
    std::string lang;              // code block language
    std::vector<std::vector<std::vector<MdRun>>> rows;  // table: rows of cells of runs; rows[0] is the header
    std::vector<int> align;        // table: per column -1 left, 0 centre, 1 right
    u32 line = 0, lineEnd = 0;     // source lines [line, lineEnd)
};

std::vector<MdBlock> parseMarkdown(const std::string& text);
std::vector<MdRun> parseInline(const std::string& s);

}  // namespace jnp
