#pragma once
#include <windows.h>
#include <tchar.h>
#include <string>
#include <vector>

// Text manipulation functions
std::basic_string<TCHAR> SortLines(const std::basic_string<TCHAR>& input);
// JoinLines removed
std::basic_string<TCHAR> IndentLines(const std::basic_string<TCHAR>& input);
std::basic_string<TCHAR> UnindentLines(const std::basic_string<TCHAR>& input);
std::basic_string<TCHAR> ToUpperCase(const std::basic_string<TCHAR>& input);
std::basic_string<TCHAR> ToLowerCase(const std::basic_string<TCHAR>& input);
std::basic_string<TCHAR> ToCapitalize(const std::basic_string<TCHAR>& input);
std::basic_string<TCHAR> ToSentenceCase(const std::basic_string<TCHAR>& input);

// Navigation helpers
int CalculateNextSubword(const TCHAR* text, int length);
int CalculatePrevSubword(const TCHAR* text, int length);

struct Range {
    long start;
    long end;
};
Range FindEnclosingBrackets(const TCHAR* text, int length, long selStart, long selEnd);

// Selection Hierarchy
enum SelectionLevel {
    SEL_NONE,
    SEL_CHARACTER,
    SEL_SUBWORD,
    SEL_WORD,
    SEL_PHRASE,
    SEL_SENTENCE,
    SEL_LINE,
    SEL_PARAGRAPH,
    SEL_SECTION,
    SEL_DOCUMENT
};

Range GetSelectionRange(const TCHAR* text, int length, long selStart, long selEnd, SelectionLevel level);
