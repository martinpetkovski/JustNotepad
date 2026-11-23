#include "datatree.h"

#include <cctype>
#include <cstdio>
#include <cstring>

namespace jnp {

namespace {

void appendUtf8(std::string& out, u32 cp) {
    if (cp < 0x80) {
        out += (char)cp;
    } else if (cp < 0x800) {
        out += (char)(0xC0 | (cp >> 6));
        out += (char)(0x80 | (cp & 0x3F));
    } else if (cp < 0x10000) {
        out += (char)(0xE0 | (cp >> 12));
        out += (char)(0x80 | ((cp >> 6) & 0x3F));
        out += (char)(0x80 | (cp & 0x3F));
    } else {
        out += (char)(0xF0 | (cp >> 18));
        out += (char)(0x80 | ((cp >> 12) & 0x3F));
        out += (char)(0x80 | ((cp >> 6) & 0x3F));
        out += (char)(0x80 | (cp & 0x3F));
    }
}

int hexv(char c) {
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return c - 'a' + 10;
    if (c >= 'A' && c <= 'F') return c - 'A' + 10;
    return -1;
}

// ---------------------------------------------------------------- JSON

struct JsonParser {
    const char* p;
    size_t n, i = 0, maxNodes;
    DataTree t;

    bool fail(const char* msg) {
        if (t.ok) {
            t.ok = false;
            t.error = msg;
            t.errorPos = i;
        }
        return false;
    }
    void space() {
        for (;;) {
            while (i < n && (p[i] == ' ' || p[i] == '\t' || p[i] == '\r' || p[i] == '\n')) i++;
            // Comments (JSONC, tsconfig, VS Code settings)
            if (i + 1 < n && p[i] == '/' && p[i + 1] == '/') {
                while (i < n && p[i] != '\n') i++;
            } else if (i + 1 < n && p[i] == '/' && p[i + 1] == '*') {
                i += 2;
                while (i + 1 < n && !(p[i] == '*' && p[i + 1] == '/')) i++;
                i = i + 2 <= n ? i + 2 : n;
            } else {
                return;
            }
        }
    }
    int add(DataNode::Kind k, int parent, const std::string& label) {
        if (t.nodes.size() >= maxNodes) {
            t.truncated = true;
            return -1;
        }
        DataNode d;
        d.kind = k;
        d.parent = parent;
        d.label = label;
        t.nodes.push_back(std::move(d));
        int id = (int)t.nodes.size() - 1;
        if (parent >= 0) t.nodes[(size_t)parent].kids.push_back(id);
        return id;
    }
    bool str(std::string& out) {  // at the opening quote
        char q = p[i++];
        while (i < n && p[i] != q) {
            char c = p[i++];
            if (c == '\n') return fail("Line break inside a string");
            if (c != '\\') {
                out += c;
                continue;
            }
            if (i >= n) break;
            char e = p[i++];
            switch (e) {
            case 'n': out += '\n'; break;
            case 't': out += '\t'; break;
            case 'r': out += '\r'; break;
            case 'b': out += '\b'; break;
            case 'f': out += '\f'; break;
            case 'u': {
                u32 cp = 0;
                for (int k = 0; k < 4; k++) {
                    int v = i < n ? hexv(p[i]) : -1;
                    if (v < 0) return fail("Bad \\u escape");
                    cp = cp * 16 + (u32)v;
                    i++;
                }
                if (cp >= 0xD800 && cp < 0xDC00 && i + 5 < n && p[i] == '\\' && p[i + 1] == 'u') {
                    u32 lo = 0;
                    bool good = true;
                    for (int k = 0; k < 4; k++) {
                        int v = hexv(p[i + 2 + k]);
                        if (v < 0) good = false;
                        lo = lo * 16 + (u32)(v < 0 ? 0 : v);
                    }
                    if (good && lo >= 0xDC00 && lo < 0xE000) {
                        cp = 0x10000 + ((cp - 0xD800) << 10) + (lo - 0xDC00);
                        i += 6;
                    }
                }
                appendUtf8(out, cp);
                break;
            }
            default: out += e; break;
            }
        }
        if (i >= n) return fail("Unclosed string");
        i++;
        return true;
    }
    bool value(int parent, const std::string& label, int depth) {
        if (depth > 512) return fail("Nested too deeply");
        space();
        if (i >= n) return fail("Expected a value");
        const size_t s = i;
        char c = p[i];
        if (c == '{' || c == '[') {
            const bool obj = c == '{';
            int id = add(obj ? DataNode::Object : DataNode::Array, parent, label);
            i++;
            size_t count = 0;
            for (;;) {
                space();
                if (i < n && p[i] == (obj ? '}' : ']')) {
                    i++;
                    break;
                }
                std::string key;
                if (obj) {
                    if (i >= n || (p[i] != '"' && p[i] != '\'')) return fail("Expected a \"key\"");
                    if (!str(key)) return false;
                    space();
                    if (i >= n || p[i] != ':') return fail("Expected ':' after the key");
                    i++;
                } else {
                    key = "[" + std::to_string(count) + "]";
                }
                if (!value(id, key, depth + 1)) return false;
                count++;
                space();
                if (i < n && p[i] == ',') {
                    i++;
                    continue;  // a trailing comma before } or ] is accepted
                }
                if (i < n && p[i] == (obj ? '}' : ']')) {
                    i++;
                    break;
                }
                return fail(obj ? "Expected ',' or '}'" : "Expected ',' or ']'");
            }
            if (id >= 0) {
                DataNode& d = t.nodes[(size_t)id];
                d.start = s;
                d.end = i;
                d.value = (obj ? "{" : "[") + std::to_string(count) + (obj ? "}" : "]");
            }
            return true;
        }
        if (c == '"' || c == '\'') {
            std::string v;
            if (!str(v)) return false;
            int id = add(DataNode::String, parent, label);
            if (id >= 0) {
                DataNode& d = t.nodes[(size_t)id];
                d.value = std::move(v);
                d.start = d.vstart = s;
                d.end = d.vend = i;
            }
            return true;
        }
        // number / true / false / null
        while (i < n && (isalnum((unsigned char)p[i]) || p[i] == '-' || p[i] == '+' || p[i] == '.')) i++;
        std::string w(p + s, i - s);
        if (w.empty()) return fail("Unexpected character");
        DataNode::Kind k;
        if (w == "true" || w == "false") k = DataNode::Bool;
        else if (w == "null") k = DataNode::Null;
        else if (jsonIsLiteral(w)) k = DataNode::Number;
        else {
            i = s;
            return fail("Not a valid value");
        }
        int id = add(k, parent, label);
        if (id >= 0) {
            DataNode& d = t.nodes[(size_t)id];
            d.value = w;
            d.start = d.vstart = s;
            d.end = d.vend = i;
        }
        return true;
    }
};

// ---------------------------------------------------------------- XML

std::string xmlDecode(const char* p, size_t n) {
    std::string out;
    out.reserve(n);
    for (size_t i = 0; i < n; i++) {
        if (p[i] != '&') {
            out += p[i];
            continue;
        }
        size_t e = i + 1;
        while (e < n && e - i < 12 && p[e] != ';') e++;
        if (e >= n || p[e] != ';') {
            out += '&';
            continue;
        }
        std::string ent(p + i + 1, e - i - 1);
        if (ent == "amp") out += '&';
        else if (ent == "lt") out += '<';
        else if (ent == "gt") out += '>';
        else if (ent == "quot") out += '"';
        else if (ent == "apos") out += '\'';
        else if (!ent.empty() && ent[0] == '#') {
            u32 cp = 0;
            bool hex = ent.size() > 1 && (ent[1] == 'x' || ent[1] == 'X');
            for (size_t k = hex ? 2 : 1; k < ent.size(); k++) cp = hex ? cp * 16 + (u32)(hexv(ent[k]) < 0 ? 0 : hexv(ent[k])) : cp * 10 + (u32)(ent[k] - '0');
            appendUtf8(out, cp);
        } else {
            out.append(p + i, e - i + 1);
        }
        i = e;
    }
    return out;
}

struct XmlParser {
    const char* p;
    size_t n, i = 0, maxNodes;
    DataTree t;

    void fail(const char* msg, size_t at) {
        if (t.ok) {
            t.ok = false;
            t.error = msg;
            t.errorPos = at;
        }
    }
    int add(DataNode::Kind k, int parent, std::string label) {
        if (t.nodes.size() >= maxNodes) {
            t.truncated = true;
            return -1;
        }
        DataNode d;
        d.kind = k;
        d.parent = parent;
        d.label = std::move(label);
        t.nodes.push_back(std::move(d));
        int id = (int)t.nodes.size() - 1;
        if (parent >= 0) t.nodes[(size_t)parent].kids.push_back(id);
        return id;
    }
    static bool nameChar(char c) {
        return isalnum((unsigned char)c) || c == '_' || c == ':' || c == '-' || c == '.' || (unsigned char)c >= 0x80;
    }

    void run() {
        // A document root holding the top-level nodes.
        int root = add(DataNode::Element, -1, "document");
        t.nodes[0].start = 0;
        t.nodes[0].end = n;
        std::vector<int> stack{root};
        while (i < n && !t.truncated) {
            int cur = stack.back();
            if (p[i] != '<') {
                size_t s = i;
                while (i < n && p[i] != '<') i++;
                // Text: trimmed; whitespace-only text between tags is skipped.
                size_t a = s, b = i;
                while (a < b && isspace((unsigned char)p[a])) a++;
                while (b > a && isspace((unsigned char)p[b - 1])) b--;
                if (b > a && cur != root) {
                    int id = add(DataNode::Text, cur, "#text");
                    if (id >= 0) {
                        DataNode& d = t.nodes[(size_t)id];
                        d.value = xmlDecode(p + a, b - a);
                        d.start = d.vstart = a;
                        d.end = d.vend = b;
                    }
                }
                continue;
            }
            const size_t s = i;
            if (n - i >= 4 && memcmp(p + i, "<!--", 4) == 0) {
                const char* e = nullptr;
                for (size_t k = i + 4; k + 2 < n; k++)
                    if (p[k] == '-' && p[k + 1] == '-' && p[k + 2] == '>') {
                        e = p + k;
                        break;
                    }
                if (!e) {
                    fail("Unclosed comment", s);
                    return;
                }
                size_t ce = (size_t)(e - p);
                int id = add(DataNode::Comment, cur, "#comment");
                if (id >= 0) {
                    DataNode& d = t.nodes[(size_t)id];
                    d.value = std::string(p + i + 4, ce - i - 4);
                    d.start = s;
                    d.end = ce + 3;
                    d.vstart = i + 4;
                    d.vend = ce;
                }
                i = ce + 3;
                continue;
            }
            if (n - i >= 9 && memcmp(p + i, "<![CDATA[", 9) == 0) {
                size_t k = i + 9;
                while (k + 2 < n && !(p[k] == ']' && p[k + 1] == ']' && p[k + 2] == '>')) k++;
                int id = add(DataNode::Text, cur, "#cdata");
                if (id >= 0) {
                    DataNode& d = t.nodes[(size_t)id];
                    d.value = std::string(p + i + 9, (k < n ? k : n) - i - 9);
                    d.start = s;
                    d.end = k + 3 <= n ? k + 3 : n;
                }
                i = k + 3 <= n ? k + 3 : n;
                continue;
            }
            if (i + 1 < n && (p[i + 1] == '?' || p[i + 1] == '!')) {  // <?xml ...?>, <!DOCTYPE ...>
                while (i < n && p[i] != '>') i++;
                i = i < n ? i + 1 : n;
                continue;
            }
            if (i + 1 < n && p[i + 1] == '/') {  // </name>
                size_t a = i + 2, b = a;
                while (b < n && nameChar(p[b])) b++;
                std::string name(p + a, b - a);
                while (b < n && p[b] != '>') b++;
                i = b < n ? b + 1 : n;
                // Close the matching open element (and anything left open inside it).
                int k = (int)stack.size() - 1;
                while (k > 0 && t.nodes[(size_t)stack[(size_t)k]].label != name) k--;
                if (k == 0) {
                    fail("Closing tag without an opening one", s);
                    continue;
                }
                if (k != (int)stack.size() - 1) fail("Element not closed", t.nodes[(size_t)stack.back()].start);
                while ((int)stack.size() > k) {
                    t.nodes[(size_t)stack.back()].end = i;
                    stack.pop_back();
                }
                continue;
            }
            // <name attr="v" ...> or <name .../>
            size_t a = i + 1, b = a;
            while (b < n && nameChar(p[b])) b++;
            if (b == a) {
                fail("Expected a tag name after '<'", s);
                i++;
                continue;
            }
            int el = add(DataNode::Element, cur, std::string(p + a, b - a));
            i = b;
            bool selfClosed = false;
            size_t attrs = 0;
            while (i < n) {
                while (i < n && isspace((unsigned char)p[i])) i++;
                if (i < n && p[i] == '>') {
                    i++;
                    break;
                }
                if (i + 1 < n && p[i] == '/' && p[i + 1] == '>') {
                    i += 2;
                    selfClosed = true;
                    break;
                }
                size_t na = i;
                while (i < n && nameChar(p[i])) i++;
                if (i == na) {
                    fail("Unexpected character in a tag", i);
                    i++;
                    continue;
                }
                size_t nb = i;
                while (i < n && isspace((unsigned char)p[i])) i++;
                size_t va = i, vb = i;
                std::string v;
                if (i < n && p[i] == '=') {
                    i++;
                    while (i < n && isspace((unsigned char)p[i])) i++;
                    if (i < n && (p[i] == '"' || p[i] == '\'')) {
                        char q = p[i++];
                        va = i;
                        while (i < n && p[i] != q) i++;
                        vb = i;
                        if (i < n) i++;
                        else fail("Unclosed attribute value", va);
                    } else {
                        va = i;
                        while (i < n && !isspace((unsigned char)p[i]) && p[i] != '>') i++;
                        vb = i;
                    }
                    v = xmlDecode(p + va, vb - va);
                }
                if (el >= 0) {
                    int id = add(DataNode::Attribute, el, "@" + std::string(p + na, nb - na));
                    if (id >= 0) {
                        DataNode& d = t.nodes[(size_t)id];
                        d.value = std::move(v);
                        d.start = na;
                        d.end = i;
                        d.vstart = va;
                        d.vend = vb;
                    }
                }
                attrs++;
            }
            if (el >= 0) {
                t.nodes[(size_t)el].start = s;
                t.nodes[(size_t)el].end = i;
                if (!selfClosed) stack.push_back(el);
            }
        }
        if (stack.size() > 1) fail("Element not closed", t.nodes[(size_t)stack.back()].start);
        while (stack.size() > 1) {
            t.nodes[(size_t)stack.back()].end = n;
            stack.pop_back();
        }
        // Element summaries: its text if it only holds text, else how many children.
        for (DataNode& d : t.nodes) {
            if (d.kind != DataNode::Element) continue;
            size_t elements = 0, texts = 0;
            int onlyText = -1;
            for (int k : d.kids) {
                const DataNode& c = t.nodes[(size_t)k];
                if (c.kind == DataNode::Element) elements++;
                if (c.kind == DataNode::Text) {
                    texts++;
                    onlyText = k;
                }
            }
            if (elements == 0 && texts == 1) {
                const DataNode& c = t.nodes[(size_t)onlyText];
                d.value = c.value;
                d.vstart = c.vstart;
                d.vend = c.vend;
            } else if (elements) {
                d.value = "<" + std::to_string(elements) + ">";
            }
        }
    }
};

}  // namespace

DataTree parseJson(const char* p, size_t n, size_t maxNodes) {
    JsonParser j{p, n, 0, maxNodes, {}};
    // Skip a UTF-8 byte order mark.
    if (n >= 3 && (unsigned char)p[0] == 0xEF && (unsigned char)p[1] == 0xBB && (unsigned char)p[2] == 0xBF) j.i = 3;
    j.space();
    if (j.i >= n) {
        j.t.ok = false;
        j.t.error = "The document is empty";
        return std::move(j.t);
    }
    if (j.value(-1, "root", 0)) {
        j.space();
        if (j.i < n && !j.t.truncated) j.fail("Unexpected text after the end");
    }
    return std::move(j.t);
}

DataTree parseXml(const char* p, size_t n, size_t maxNodes) {
    XmlParser x{p, n, 0, maxNodes, {}};
    x.run();
    return std::move(x.t);
}

std::string jsonQuote(const std::string& s) {
    std::string out = "\"";
    for (unsigned char c : s) {
        switch (c) {
        case '"': out += "\\\""; break;
        case '\\': out += "\\\\"; break;
        case '\n': out += "\\n"; break;
        case '\r': out += "\\r"; break;
        case '\t': out += "\\t"; break;
        default:
            if (c < 0x20) {
                char b[8];
                snprintf(b, sizeof(b), "\\u%04x", c);
                out += b;
            } else {
                out += (char)c;
            }
        }
    }
    return out + "\"";
}

bool jsonIsLiteral(const std::string& s) {
    if (s == "true" || s == "false" || s == "null") return true;
    // -?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?
    size_t i = 0, n = s.size();
    if (i < n && s[i] == '-') i++;
    if (i >= n || !isdigit((unsigned char)s[i])) return false;
    if (s[i] == '0') i++;
    else
        while (i < n && isdigit((unsigned char)s[i])) i++;
    if (i < n && s[i] == '.') {
        i++;
        if (i >= n || !isdigit((unsigned char)s[i])) return false;
        while (i < n && isdigit((unsigned char)s[i])) i++;
    }
    if (i < n && (s[i] == 'e' || s[i] == 'E')) {
        i++;
        if (i < n && (s[i] == '+' || s[i] == '-')) i++;
        if (i >= n || !isdigit((unsigned char)s[i])) return false;
        while (i < n && isdigit((unsigned char)s[i])) i++;
    }
    return i == n;
}

std::string xmlEscape(const std::string& s, bool attribute) {
    std::string out;
    for (char c : s) {
        switch (c) {
        case '&': out += "&amp;"; break;
        case '<': out += "&lt;"; break;
        case '>': out += "&gt;"; break;
        case '"': out += attribute ? "&quot;" : "\""; break;
        default: out += c;
        }
    }
    return out;
}

}  // namespace jnp
