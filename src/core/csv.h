#pragma once
// CSV / TSV: rows and fields with their byte positions, for the table view.
// Quoted fields may hold the delimiter, quotes ("") and line breaks.
// No OS headers: part of the portable core.
#include <cstddef>
#include <string>
#include <vector>

#include "types.h"

namespace jnp {

struct CsvField {
    u64 start = 0, end = 0;  // raw span, quotes included
};

// The delimiter a text most likely uses: ',', ';', '\t' or '|'.
char csvSniff(const char* p, size_t n, bool tsvName);
// Start offset of every row, plus one final entry: the end of the text.
// Line breaks inside quotes do not end a row. Stops after maxRows.
std::vector<u64> csvRows(const char* p, size_t n, size_t maxRows, bool& truncated);
// The fields of one row (p .. p+n, without its line break); offsets relative to `base`.
void csvFields(const char* p, size_t n, u64 base, char delim, std::vector<CsvField>& out);
// A raw field's value (quotes removed, "" -> ").
std::string csvValue(const char* p, size_t n);
// A value written as a field: quoted when it has to be.
std::string csvQuote(const std::string& v, char delim);

}  // namespace jnp
