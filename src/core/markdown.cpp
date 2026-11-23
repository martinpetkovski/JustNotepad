#include "markdown.h"

#include <cctype>
#include <algorithm>
#include <cstdlib>
#include <cstring>

namespace jnp {

namespace {

std::string trim(const std::string& s) {
    size_t a = 0, b = s.size();
    while (a < b && (s[a] == ' ' || s[a] == '\t')) a++;
    while (b > a && (s[b - 1] == ' ' || s[b - 1] == '\t' || s[b - 1] == '\r')) b--;
    return s.substr(a, b - a);
}

size_t indentOf(const std::string& s) {
    size_t n = 0;
    for (char c : s) {
        if (c == ' ') n++;
        else if (c == '\t') n += 4;
        else break;
    }
    return n;
}

bool isRule(const std::string& t) {
    if (t.size() < 3) return false;
    char c = t[0];
    if (c != '-' && c != '*' && c != '_') return false;
    int count = 0;
    for (char x : t) {
        if (x == c) count++;
        else if (x != ' ') return false;
    }
    return count >= 3;
}

// "- ", "* ", "+ ", "1. ", "1) " -> length of the marker including the space.
size_t listMarker(const std::string& t, bool& ordered, int& number) {
    if (t.size() >= 2 && (t[0] == '-' || t[0] == '*' || t[0] == '+') && t[1] == ' ') {
        ordered = false;
        return 2;
    }
    size_t i = 0;
    while (i < t.size() && i < 9 && isdigit((unsigned char)t[i])) i++;
    if (i > 0 && i + 1 < t.size() && (t[i] == '.' || t[i] == ')') && t[i + 1] == ' ') {
        ordered = true;
        number = atoi(t.substr(0, i).c_str());
        return i + 2;
    }
    return 0;
}

std::vector<std::string> splitRow(const std::string& line) {
    std::string t = trim(line);
    if (!t.empty() && t[0] == '|') t.erase(0, 1);
    if (!t.empty() && t.back() == '|') t.pop_back();
    std::vector<std::string> cells;
    std::string cur;
    for (size_t i = 0; i < t.size(); i++) {
        if (t[i] == '\\' && i + 1 < t.size() && t[i + 1] == '|') {
            cur += '|';
            i++;
        } else if (t[i] == '|') {
            cells.push_back(trim(cur));
            cur.clear();
        } else {
            cur += t[i];
        }
    }
    cells.push_back(trim(cur));
    return cells;
}

bool isTableSeparator(const std::string& line, std::vector<int>& align) {
    std::string t = trim(line);
    if (t.find('-') == std::string::npos || (t.find('|') == std::string::npos)) return false;
    align.clear();
    for (const std::string& c : splitRow(t)) {
        if (c.empty()) return false;
        for (char x : c)
            if (x != '-' && x != ':' && x != ' ') return false;
        bool l = c.front() == ':', r = c.back() == ':';
        align.push_back(l && r ? 0 : r ? 1 : -1);
    }
    return !align.empty();
}

void addText(std::vector<MdRun>& out, const std::string& s, u8 style, const std::string& url = std::string()) {
    if (s.empty()) return;
    if (!out.empty() && out.back().style == style && out.back().url == url) {
        out.back().text += s;
        return;
    }
    MdRun r;
    r.text = s;
    r.style = style;
    r.url = url;
    out.push_back(std::move(r));
}

void inlineInto(std::vector<MdRun>& out, const std::string& s, u8 style, int depth) {
    std::string text;
    auto flush = [&]() {
        addText(out, text, style);
        text.clear();
    };
    size_t i = 0;
    while (i < s.size()) {
        char c = s[i];
        if (c == '\\' && i + 1 < s.size() && ispunct((unsigned char)s[i + 1])) {
            text += s[i + 1];
            i += 2;
            continue;
        }
        if (c == '`') {  // code span: `x` or ``x``
            size_t ticks = 0;
            while (i + ticks < s.size() && s[i + ticks] == '`') ticks++;
            std::string fence(ticks, '`');
            size_t e = s.find(fence, i + ticks);
            if (e != std::string::npos) {
                flush();
                addText(out, trim(s.substr(i + ticks, e - i - ticks)), (u8)(style | MdRun::Code));
                i = e + ticks;
                continue;
            }
        }
        if (depth < 8 && (c == '*' || c == '_' || c == '~')) {
            size_t k = 0;
            while (i + k < s.size() && s[i + k] == c && k < 3) k++;
            const bool strike = c == '~';
            if (!(strike && k < 2) && i + k < s.size() && s[i + k] != ' ') {
                std::string mark(strike ? 2 : k, c);
                if (strike) k = 2;
                // Underscores inside words (snake_case) are not emphasis.
                bool wordInside = c == '_' && i > 0 && isalnum((unsigned char)s[i - 1]);
                size_t e = wordInside ? std::string::npos : s.find(mark, i + k);
                if (e != std::string::npos && s[e - 1] != ' ') {
                    flush();
                    u8 st = style;
                    if (strike) st |= MdRun::Strike;
                    else if (k == 1) st |= MdRun::Italic;
                    else if (k == 2) st |= MdRun::Bold;
                    else st |= MdRun::Bold | MdRun::Italic;
                    inlineInto(out, s.substr(i + k, e - i - k), st, depth + 1);
                    i = e + k;
                    continue;
                }
            }
        }
        if (c == '[' || (c == '!' && i + 1 < s.size() && s[i + 1] == '[')) {
            const bool image = c == '!';
            size_t a = i + (image ? 2 : 1);
            int nest = 1;
            size_t e = a;
            while (e < s.size() && nest) {
                if (s[e] == '[') nest++;
                else if (s[e] == ']') nest--;
                if (nest) e++;
            }
            if (e < s.size() && e + 1 < s.size() && s[e + 1] == '(') {
                size_t u = s.find(')', e + 2);
                if (u != std::string::npos) {
                    flush();
                    std::string label = s.substr(a, e - a), url = trim(s.substr(e + 2, u - e - 2));
                    size_t sp = url.find(' ');  // [x](url "title")
                    if (sp != std::string::npos) url.resize(sp);
                    if (image) {
                        addText(out, "[image: " + (label.empty() ? url : label) + "]", (u8)(style | MdRun::Italic));
                    } else {
                        std::vector<MdRun> inner;
                        inlineInto(inner, label, (u8)(style | MdRun::Link), depth + 1);
                        for (MdRun& r : inner) {
                            r.url = url;
                            out.push_back(std::move(r));
                        }
                    }
                    i = u + 1;
                    continue;
                }
            }
        }
        if (c == '<') {  // <https://autolink>
            size_t e = s.find('>', i);
            if (e != std::string::npos) {
                std::string u = s.substr(i + 1, e - i - 1);
                if (u.rfind("http://", 0) == 0 || u.rfind("https://", 0) == 0 || u.rfind("mailto:", 0) == 0) {
                    flush();
                    addText(out, u, (u8)(style | MdRun::Link), u);
                    i = e + 1;
                    continue;
                }
            }
        }
        text += c;
        i++;
    }
    flush();
}

}  // namespace

std::vector<MdRun> parseInline(const std::string& s) {
    std::vector<MdRun> out;
    inlineInto(out, s, 0, 0);
    return out;
}

std::vector<MdBlock> parseMarkdown(const std::string& text) {
    std::vector<std::string> lines;
    size_t a = 0;
    while (a <= text.size()) {
        size_t b = text.find('\n', a);
        if (b == std::string::npos) b = text.size();
        std::string l = text.substr(a, b - a);
        if (!l.empty() && l.back() == '\r') l.pop_back();
        lines.push_back(std::move(l));
        if (b == text.size()) break;
        a = b + 1;
    }
    std::vector<MdBlock> out;
    size_t i = 0;
    std::string para;
    u32 paraLine = 0;
    int paraQuote = 0;
    auto flushPara = [&](u32 end) {
        if (para.empty()) return;
        MdBlock b;
        b.kind = MdBlock::Paragraph;
        b.runs = parseInline(para);
        b.line = paraLine;
        b.lineEnd = end;
        b.quote = paraQuote;
        out.push_back(std::move(b));
        para.clear();
    };
    while (i < lines.size()) {
        std::string raw = lines[i];
        // Quote markers in front: "> > text"
        int quote = 0;
        {
            size_t k = 0;
            for (;;) {
                while (k < raw.size() && k < 3 && raw[k] == ' ') k++;
                if (k < raw.size() && raw[k] == '>') {
                    quote++;
                    k++;
                    if (k < raw.size() && raw[k] == ' ') k++;
                } else {
                    break;
                }
            }
            raw = raw.substr(std::min(k, raw.size()));
            if (quote == 0) raw = lines[i];
        }
        const std::string t = trim(raw);
        const u32 ln = (u32)i;

        if (t.empty()) {
            flushPara(ln);
            i++;
            continue;
        }
        // Fenced code
        if (t.rfind("```", 0) == 0 || t.rfind("~~~", 0) == 0) {
            flushPara(ln);
            const std::string fence = t.substr(0, 3);
            MdBlock b;
            b.kind = MdBlock::CodeBlock;
            b.lang = trim(t.substr(3));
            b.quote = quote;
            b.line = ln;
            i++;
            while (i < lines.size() && trim(lines[i]).rfind(fence, 0) != 0) b.code.push_back(lines[i++]);
            if (i < lines.size()) i++;
            b.lineEnd = (u32)i;
            out.push_back(std::move(b));
            continue;
        }
        // Indented code (4 spaces), not inside a paragraph
        if (para.empty() && quote == 0 && indentOf(raw) >= 4 && (out.empty() || out.back().kind != MdBlock::ListItem)) {
            MdBlock b;
            b.kind = MdBlock::CodeBlock;
            b.line = ln;
            while (i < lines.size() && (indentOf(lines[i]) >= 4 || trim(lines[i]).empty())) {
                std::string l = lines[i++];
                b.code.push_back(l.size() >= 4 ? l.substr(l[0] == '\t' ? 1 : 4) : std::string());
            }
            while (!b.code.empty() && trim(b.code.back()).empty()) b.code.pop_back();
            b.lineEnd = (u32)i;
            out.push_back(std::move(b));
            continue;
        }
        // ATX heading
        if (t[0] == '#') {
            size_t h = 0;
            while (h < t.size() && t[h] == '#') h++;
            if (h <= 6 && (h == t.size() || t[h] == ' ')) {
                flushPara(ln);
                std::string body = trim(t.substr(h));
                while (!body.empty() && body.back() == '#') body.pop_back();
                MdBlock b;
                b.kind = MdBlock::Heading;
                b.level = (int)h;
                b.runs = parseInline(trim(body));
                b.quote = quote;
                b.line = ln;
                b.lineEnd = ln + 1;
                out.push_back(std::move(b));
                i++;
                continue;
            }
        }
        // Setext heading: a paragraph line followed by === or ---
        if (!para.empty() && (t.find_first_not_of('=') == std::string::npos ||
                              (t.find_first_not_of('-') == std::string::npos && t.size() >= 2))) {
            MdBlock b;
            b.kind = MdBlock::Heading;
            b.level = t[0] == '=' ? 1 : 2;
            b.runs = parseInline(para);
            b.quote = paraQuote;
            b.line = paraLine;
            b.lineEnd = ln + 1;
            out.push_back(std::move(b));
            para.clear();
            i++;
            continue;
        }
        if (isRule(t)) {
            flushPara(ln);
            MdBlock b;
            b.kind = MdBlock::Rule;
            b.quote = quote;
            b.line = ln;
            b.lineEnd = ln + 1;
            out.push_back(std::move(b));
            i++;
            continue;
        }
        // Table: header row, separator row, body rows
        std::vector<int> align;
        if (t.find('|') != std::string::npos && i + 1 < lines.size() && isTableSeparator(lines[i + 1], align)) {
            flushPara(ln);
            MdBlock b;
            b.kind = MdBlock::Table;
            b.align = align;
            b.quote = quote;
            b.line = ln;
            auto addRow = [&](const std::string& l) {
                std::vector<std::vector<MdRun>> row;
                for (const std::string& c : splitRow(l)) row.push_back(parseInline(c));
                row.resize(align.size());
                b.rows.push_back(std::move(row));
            };
            addRow(t);
            i += 2;
            while (i < lines.size() && trim(lines[i]).find('|') != std::string::npos) addRow(lines[i++]);
            b.lineEnd = (u32)i;
            out.push_back(std::move(b));
            continue;
        }
        // List item
        bool ordered = false;
        int number = 0;
        size_t mk = listMarker(t, ordered, number);
        if (mk) {
            flushPara(ln);
            MdBlock b;
            b.kind = MdBlock::ListItem;
            b.ordered = ordered;
            b.number = number;
            b.level = (int)(indentOf(raw) / 2);
            b.quote = quote;
            b.line = ln;
            std::string body = t.substr(mk);
            if (body.size() >= 3 && body[0] == '[' && (body[1] == ' ' || body[1] == 'x' || body[1] == 'X') && body[2] == ']') {
                b.task = body[1] == ' ' ? 0 : 1;
                body = trim(body.substr(3));
            }
            i++;
            // Lazy continuation lines belong to the item.
            while (i < lines.size()) {
                std::string nt = trim(lines[i]);
                bool o2;
                int n2;
                if (nt.empty() || listMarker(nt, o2, n2) || nt[0] == '#' || nt[0] == '>' || nt.rfind("```", 0) == 0 || isRule(nt))
                    break;
                body += " " + nt;
                i++;
            }
            b.runs = parseInline(body);
            b.lineEnd = (u32)i;
            out.push_back(std::move(b));
            continue;
        }
        // Paragraph text (lines join; two trailing spaces or "\" make a line break)
        if (para.empty()) {
            paraLine = ln;
            paraQuote = quote;
        } else if (quote != paraQuote) {
            flushPara(ln);
            paraLine = ln;
            paraQuote = quote;
        }
        bool hardBreak = raw.size() >= 2 && raw.compare(raw.size() - 2, 2, "  ") == 0;
        std::string piece = t;
        if (!piece.empty() && piece.back() == '\\') {
            piece.pop_back();
            hardBreak = true;
        }
        if (!para.empty() && para.back() != '\n') para += ' ';
        para += piece;
        if (hardBreak) para += '\n';
        i++;
    }
    flushPara((u32)lines.size());
    return out;
}

}  // namespace jnp
