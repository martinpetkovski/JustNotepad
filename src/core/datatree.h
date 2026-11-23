#pragma once
// JSON and XML as a tree, for the data view: every node knows where it sits
// in the source, so the view can select it in the editor and edit its value
// in place. The parsers are forgiving (JSON comments and trailing commas,
// unclosed XML tags) and report the first error with its byte position.
// No OS headers: part of the portable core.
#include <cstddef>
#include <string>
#include <vector>

#include "types.h"

namespace jnp {

struct DataNode {
    enum Kind : u8 {
        Object,     // JSON {}
        Array,      // JSON []
        String,
        Number,
        Bool,
        Null,
        Element,    // XML <tag>
        Attribute,  // XML name="value"
        Text,       // XML text between tags
        Comment,    // XML <!-- -->
    };
    Kind kind = Null;
    std::string label;  // key, element / attribute name, or "[3]" for array items
    std::string value;  // decoded value of leaves; a short summary for containers
    u64 start = 0, end = 0;    // the whole node in the source
    u64 vstart = 0, vend = 0;  // the editable part (raw, as written); vend == vstart: none
    int parent = -1;
    std::vector<int> kids;
};

struct DataTree {
    std::vector<DataNode> nodes;  // nodes[0] is the root when !nodes.empty()
    bool ok = true;
    u64 errorPos = 0;
    std::string error;
    bool truncated = false;  // stopped at the node limit
};

DataTree parseJson(const char* p, size_t n, size_t maxNodes);
DataTree parseXml(const char* p, size_t n, size_t maxNodes);

// Helpers for editing a value in place.
std::string jsonQuote(const std::string& utf8);              // "..." with escapes
bool jsonIsLiteral(const std::string& s);                     // a number, true, false or null
std::string xmlEscape(const std::string& s, bool attribute);  // &amp; &lt; ... (and &quot; in attributes)

}  // namespace jnp
