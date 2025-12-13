#include "../src/TextHelpers.h"
#include <iostream>
#include <cassert>
#include <string>

// Simple test framework
#define TEST(name) void name()
#define RUN_TEST(name) do { std::cout << "Running " << #name << "... "; name(); std::cout << "PASSED" << std::endl; } while(0)

// Helper for string literals
#ifdef UNICODE
#define S(x) L##x
#define TO_STRING(x) std::wstring(x)
#else
#define S(x) x
#define TO_STRING(x) std::string(x)
#endif

// Custom assert to print values
template<typename T>
void AssertEq(const T& a, const T& b, const char* a_str, const char* b_str, int line)
{
    if (a != b)
    {
        std::cerr << "FAILED at line " << line << ": " << a_str << " != " << b_str << "\n";
        // std::cerr << "Expected: " << b << "\nActual:   " << a << std::endl; // Printing wstring is tricky with cerr
        exit(1);
    }
}

#define ASSERT_EQ(a, b) AssertEq(a, b, #a, #b, __LINE__)

TEST(TestSortLines)
{
    std::basic_string<TCHAR> input = S("b\r\na\r\nc");
    std::basic_string<TCHAR> expected = S("a\r\nb\r\nc");
    ASSERT_EQ(SortLines(input), expected);
}

// JoinLines removed

TEST(TestIndentLines)
{
    std::basic_string<TCHAR> input = S("Line1\r\nLine2");
    std::basic_string<TCHAR> expected = S("\tLine1\r\n\tLine2");
    std::basic_string<TCHAR> result = IndentLines(input);
    ASSERT_EQ(result, expected);

    // Multiple indents
    std::basic_string<TCHAR> expected2 = S("\t\tLine1\r\n\t\tLine2");
    ASSERT_EQ(IndentLines(result), expected2);
}

TEST(TestUnindentLines)
{
    std::basic_string<TCHAR> input = S("\t\tLine1\r\n        Line2\r\n  Line3");
    
    // First unindent
    // \t\tLine1 -> \tLine1
    //         Line2 (8 spaces) ->     Line2 (4 spaces)
    //   Line3 (2 spaces) ->  Line3 (1 space)
    std::basic_string<TCHAR> expected1 = S("\tLine1\r\n    Line2\r\n Line3");
    std::basic_string<TCHAR> result1 = UnindentLines(input);
    ASSERT_EQ(result1, expected1);

    // Second unindent
    // \tLine1 -> Line1
    //     Line2 (4 spaces) -> Line2
    //  Line3 (1 space) -> Line3
    std::basic_string<TCHAR> expected2 = S("Line1\r\nLine2\r\nLine3");
    ASSERT_EQ(UnindentLines(result1), expected2);
}

TEST(TestUpperCase)
{
    std::basic_string<TCHAR> input = S("Hello World");
    std::basic_string<TCHAR> expected = S("HELLO WORLD");
    ASSERT_EQ(ToUpperCase(input), expected);
}

TEST(TestLowerCase)
{
    std::basic_string<TCHAR> input = S("Hello World");
    std::basic_string<TCHAR> expected = S("hello world");
    ASSERT_EQ(ToLowerCase(input), expected);
}

TEST(TestCapitalize)
{
    std::basic_string<TCHAR> input = S("hello world");
    std::basic_string<TCHAR> expected = S("Hello World");
    ASSERT_EQ(ToCapitalize(input), expected);
}

TEST(TestSentenceCase)
{
    std::basic_string<TCHAR> input = S("hello world. how are you?");
    std::basic_string<TCHAR> expected = S("Hello world. How are you?");
    ASSERT_EQ(ToSentenceCase(input), expected);
}

TEST(TestSubwordNavigation)
{
    // camelCase
    std::basic_string<TCHAR> s1 = S("camelCase");
    ASSERT_EQ(CalculateNextSubword(s1.c_str(), (int)s1.length()), 5); // "camel" -> 5
    
    // snake_case
    std::basic_string<TCHAR> s2 = S("snake_case");
    ASSERT_EQ(CalculateNextSubword(s2.c_str(), (int)s2.length()), 5); // "snake" -> 5
    
    // Prev
    ASSERT_EQ(CalculatePrevSubword(s1.c_str(), (int)s1.length()), 5); // "Case" -> start at 5
    ASSERT_EQ(CalculatePrevSubword(s2.c_str(), (int)s2.length()), 6); // "case" -> start at 6 (after _)
}

TEST(TestFindEnclosingBrackets)
{
    std::basic_string<TCHAR> text = S("func(arg1, arg2)");
    // Select "arg1" -> 5, 9
    Range r = FindEnclosingBrackets(text.c_str(), (int)text.length(), 5, 9);
    ASSERT_EQ(r.start, 4L); // (
    ASSERT_EQ(r.end, 16L); // ) + 1
}

TEST(TestIndentLinesVector)
{
    std::vector<TCHAR> buf = { 'L', 'i', 'n', 'e', '1', '\r', '\n', 'L', 'i', 'n', 'e', '2', '\0' };
    std::basic_string<TCHAR> input(buf.data());
    std::basic_string<TCHAR> expected = S("\tLine1\r\n\tLine2");
    ASSERT_EQ(IndentLines(input), expected);
}

TEST(TestSelectionHierarchy)
{
    std::basic_string<TCHAR> text = S("This is a sentence. This is another one.");
    // Length: 19 + 21 = 40
    // "This is a sentence." is 0-19.
    // "This is another one." is 20-40.
    
    // Select "sentence" (10-18)
    long selStart = 10;
    long selEnd = 18;
    
    // SEL_WORD -> "sentence" (10-18)
    Range rWord = GetSelectionRange(text.c_str(), (int)text.length(), selStart, selEnd, SEL_WORD);
    ASSERT_EQ(rWord.start, 10L);
    ASSERT_EQ(rWord.end, 18L);
    
    // SEL_PHRASE -> "This is a sentence" (0-18) (assuming no commas)
    // Actually, phrase stops at sentence separators too?
    // My implementation stops at phrase separators OR newlines.
    // It does NOT stop at sentence separators unless they are also phrase separators (which they aren't in my list).
    // But it stops at newlines.
    // Wait, "This is a sentence."
    // If I select "sentence", p1 goes back to 0. p2 goes forward to 18 (before dot).
    // So 0-18.
    Range rPhrase = GetSelectionRange(text.c_str(), (int)text.length(), selStart, selEnd, SEL_PHRASE);
    ASSERT_EQ(rPhrase.start, 0L);
    ASSERT_EQ(rPhrase.end, 18L);
    
    // SEL_SENTENCE -> "This is a sentence." (0-19)
    Range rSentence = GetSelectionRange(text.c_str(), (int)text.length(), selStart, selEnd, SEL_SENTENCE);
    ASSERT_EQ(rSentence.start, 0L);
    ASSERT_EQ(rSentence.end, 19L);
    
    // SEL_LINE -> Whole text (0-40) if no newlines
    Range rLine = GetSelectionRange(text.c_str(), (int)text.length(), selStart, selEnd, SEL_LINE);
    ASSERT_EQ(rLine.start, 0L);
    ASSERT_EQ(rLine.end, 40L);
}

TEST(TestSelectionHierarchyMultiLine)
{
    std::basic_string<TCHAR> text = S("Line 1\nLine 2\nLine 3");
    // Line 1: 0-6 (\n at 6) -> 0-7
    // Line 2: 7-13 (\n at 13) -> 7-14
    // Line 3: 14-20
    
    // Select "Line 2" (7-13)
    Range r = GetSelectionRange(text.c_str(), (int)text.length(), 7, 13, SEL_LINE);
    ASSERT_EQ(r.start, 7L);
    ASSERT_EQ(r.end, 14L);
    
    // Idempotency: If 7-14 is selected, SEL_LINE should stay 7-14
    Range r2 = GetSelectionRange(text.c_str(), (int)text.length(), 7, 14, SEL_LINE);
    ASSERT_EQ(r2.start, 7L);
    ASSERT_EQ(r2.end, 14L);
}

TEST(TestExpandSelectionToLine)
{
    std::basic_string<TCHAR> text = S("Line 1\r\nLine 2\r\nLine 3");
    // Line 1: 0-8 (including \r\n)
    // Line 2: 8-16
    // Line 3: 16-22
    
    // Case 1: Cursor in Line 1
    Range r1 = GetSelectionRange(text.c_str(), (int)text.length(), 2, 2, SEL_LINE);
    ASSERT_EQ(r1.start, 0L);
    ASSERT_EQ(r1.end, 8L);
    
    // Case 2: Selection spanning Line 1 and part of Line 2
    Range r2 = GetSelectionRange(text.c_str(), (int)text.length(), 2, 10, SEL_LINE);
    ASSERT_EQ(r2.start, 0L);
    ASSERT_EQ(r2.end, 16L);
    
    // Case 3: Selection of full Line 1 (idempotency)
    Range r3 = GetSelectionRange(text.c_str(), (int)text.length(), 0, 8, SEL_LINE);
    ASSERT_EQ(r3.start, 0L);
    ASSERT_EQ(r3.end, 8L);
    
    // Case 4: Cursor at start of Line 2
    Range r4 = GetSelectionRange(text.c_str(), (int)text.length(), 8, 8, SEL_LINE);
    ASSERT_EQ(r4.start, 8L);
    ASSERT_EQ(r4.end, 16L);
}

TEST(TestSelectionHierarchyCR)
{
    // Test with CR only (RichEdit internal format often)
    std::basic_string<TCHAR> text = S("Line 1\rLine 2\rLine 3");
    // Line 1: 0-6 (\r at 6) -> 0-7
    // Line 2: 7-13 (\r at 13) -> 7-14
    // Line 3: 14-20
    
    // Select "Line 2" (7-13)
    Range r = GetSelectionRange(text.c_str(), (int)text.length(), 7, 13, SEL_LINE);
    ASSERT_EQ(r.start, 7L);
    ASSERT_EQ(r.end, 14L);
    
    // Idempotency
    Range r2 = GetSelectionRange(text.c_str(), (int)text.length(), 7, 14, SEL_LINE);
    ASSERT_EQ(r2.start, 7L);
    ASSERT_EQ(r2.end, 14L);
    
    // Cursor in Line 2
    Range r3 = GetSelectionRange(text.c_str(), (int)text.length(), 9, 9, SEL_LINE);
    ASSERT_EQ(r3.start, 7L);
    ASSERT_EQ(r3.end, 14L);
}

TEST(TestSortLinesEdgeCases)
{
    // Empty
    ASSERT_EQ(SortLines(S("")), TO_STRING(S("")));
    
    // Single line
    ASSERT_EQ(SortLines(S("abc")), TO_STRING(S("abc")));
    
    // Already sorted
    ASSERT_EQ(SortLines(S("a\r\nb\r\nc")), TO_STRING(S("a\r\nb\r\nc")));
    
    // Reverse sorted
    ASSERT_EQ(SortLines(S("c\r\nb\r\na")), TO_STRING(S("a\r\nb\r\nc")));
    
    // Duplicates
    ASSERT_EQ(SortLines(S("b\r\na\r\nb")), TO_STRING(S("a\r\nb\r\nb")));
}

TEST(TestIndentUnindentEdgeCases)
{
    // Empty
    ASSERT_EQ(IndentLines(S("")), TO_STRING(S("")));
    ASSERT_EQ(UnindentLines(S("")), TO_STRING(S("")));
    
    // Whitespace only lines
    // Indent should add tab
    ASSERT_EQ(IndentLines(S("  ")), TO_STRING(S("\t  ")));
    // Unindent should remove indentation if present
    ASSERT_EQ(UnindentLines(S("\t  ")), TO_STRING(S("  ")));
    // Unindent on spaces
    ASSERT_EQ(UnindentLines(S("      ")), TO_STRING(S("  "))); // 6 spaces -> 2 spaces (removes 4)
}

TEST(TestCaseConversionEdgeCases)
{
    // Empty
    ASSERT_EQ(ToUpperCase(S("")), TO_STRING(S("")));
    ASSERT_EQ(ToLowerCase(S("")), TO_STRING(S("")));
    ASSERT_EQ(ToCapitalize(S("")), TO_STRING(S("")));
    ASSERT_EQ(ToSentenceCase(S("")), TO_STRING(S("")));
    
    // Numbers and symbols
    std::basic_string<TCHAR> mixed = S("123!@#");
    ASSERT_EQ(ToUpperCase(mixed), mixed);
    ASSERT_EQ(ToLowerCase(mixed), mixed);
    
    // Mixed case for Capitalize
    ASSERT_EQ(ToCapitalize(S("hElLo wOrLd")), TO_STRING(S("Hello World")));
    
    // Sentence case with multiple sentences
    ASSERT_EQ(ToSentenceCase(S("hello. world.")), TO_STRING(S("Hello. World.")));
    ASSERT_EQ(ToSentenceCase(S("hello! world?")), TO_STRING(S("Hello! World?")));
}

TEST(TestFindEnclosingBracketsMore)
{
    std::basic_string<TCHAR> text = S("((inner)) [box] {brace} \"quote\" 'char'");
    int len = (int)text.length();
    
    // Nested ((inner)) - select "inner"
    // text: ( ( i n n e r ) )
    // indices: 0 1 2 3 4 5 6 7 8
    // Select "inner" -> 2-7
    Range r1 = FindEnclosingBrackets(text.c_str(), len, 2, 7);
    ASSERT_EQ(r1.start, 1L); // inner (
    ASSERT_EQ(r1.end, 8L);   // inner ) + 1
    
    // Outer ((inner)) - select "(inner)" -> 1-8
    Range r2 = FindEnclosingBrackets(text.c_str(), len, 1, 8);
    ASSERT_EQ(r2.start, 0L); // outer (
    ASSERT_EQ(r2.end, 9L);   // outer ) + 1
    
    // [box] - select "box"
    // " [box] " starts at 10
    // [ at 10, box at 11-14, ] at 14
    Range r3 = FindEnclosingBrackets(text.c_str(), len, 11, 14);
    ASSERT_EQ(r3.start, 10L);
    ASSERT_EQ(r3.end, 15L);
    
    // {brace}
    // " {brace} " starts at 16
    // { at 16, brace at 17-22, } at 22
    Range r4 = FindEnclosingBrackets(text.c_str(), len, 17, 22);
    ASSERT_EQ(r4.start, 16L);
    ASSERT_EQ(r4.end, 23L);
    
    // "quote"
    // " "quote" " starts at 24
    // " at 24, quote at 25-30, " at 30
    Range r5 = FindEnclosingBrackets(text.c_str(), len, 25, 30);
    ASSERT_EQ(r5.start, 24L);
    ASSERT_EQ(r5.end, 31L);
    
    // 'char'
    // " 'char'" starts at 32
    // ' at 32, char at 33-37, ' at 37
    Range r6 = FindEnclosingBrackets(text.c_str(), len, 33, 37);
    ASSERT_EQ(r6.start, 32L);
    ASSERT_EQ(r6.end, 38L);
    
    // No brackets
    Range r7 = FindEnclosingBrackets(text.c_str(), len, 0, len);
    ASSERT_EQ(r7.start, -1L);
    ASSERT_EQ(r7.end, -1L);
}

TEST(TestSelectionHierarchyExtended)
{
    // Paragraphs
    std::basic_string<TCHAR> text = S("Para1 Line1\r\nPara1 Line2\r\n\r\nPara2 Line1\r\nPara2 Line2");
    // Para1: 0-24 (including \r\n\r\n ?)
    // Para1 Line1\r\n (13)
    // Para1 Line2\r\n (13) -> 26
    // \r\n (2) -> 28
    // Para2 Line1\r\n (13) -> 41
    // Para2 Line2 (11) -> 52
    
    // Select inside Para1
    Range rPara1 = GetSelectionRange(text.c_str(), (int)text.length(), 5, 5, SEL_PARAGRAPH);
    ASSERT_EQ(rPara1.start, 0L);
    ASSERT_EQ(rPara1.end, 26L); 
    
    // Select inside Para2
    Range rPara2 = GetSelectionRange(text.c_str(), (int)text.length(), 30, 30, SEL_PARAGRAPH);
    ASSERT_EQ(rPara2.start, 28L);
    ASSERT_EQ(rPara2.end, (long)text.length());
    
    // Sections
    std::basic_string<TCHAR> doc = S("# Section 1\r\nContent 1\r\n# Section 2\r\nContent 2");
    
    // Select inside Section 1
    Range rSec1 = GetSelectionRange(doc.c_str(), (int)doc.length(), 15, 15, SEL_SECTION);
    ASSERT_EQ(rSec1.start, 0L);
    ASSERT_EQ(rSec1.end, 24L); // Start of Section 2
    
    // Select inside Section 2
    Range rSec2 = GetSelectionRange(doc.c_str(), (int)doc.length(), 40, 40, SEL_SECTION);
    ASSERT_EQ(rSec2.start, 24L);
    ASSERT_EQ(rSec2.end, (long)doc.length());
    
    // Document
    Range rDoc = GetSelectionRange(doc.c_str(), (int)doc.length(), 10, 20, SEL_DOCUMENT);
    ASSERT_EQ(rDoc.start, 0L);
    ASSERT_EQ(rDoc.end, (long)doc.length());
}

TEST(TestSubwordNavigationEdgeCases)
{
    // Empty
    ASSERT_EQ(CalculateNextSubword(S(""), 0), 0);
    ASSERT_EQ(CalculatePrevSubword(S(""), 0), 0);
    
    // Single char
    ASSERT_EQ(CalculateNextSubword(S("a"), 1), 1);
    ASSERT_EQ(CalculatePrevSubword(S("a"), 1), 0);
    
    // Underscores
    std::basic_string<TCHAR> s = S("__init__");
    // 0: _ -> 2 (start of init)
    ASSERT_EQ(CalculateNextSubword(s.c_str(), (int)s.length()), 2);
    
    // From end
    // 8: _ -> 6 (end of init)
    ASSERT_EQ(CalculatePrevSubword(s.c_str(), (int)s.length()), 6);
    
    // Mixed types
    std::basic_string<TCHAR> m = S("camelCase_snake");
    
    const TCHAR* ptr = m.c_str();
    int len = (int)m.length();
    int current = 0;
    
    // "camel"
    int next = CalculateNextSubword(ptr + current, len - current);
    current += next;
    ASSERT_EQ(current, 5); // At 'C'
    
    // "Case"
    next = CalculateNextSubword(ptr + current, len - current);
    current += next;
    ASSERT_EQ(current, 9); // At '_'
    
    // "_"
    next = CalculateNextSubword(ptr + current, len - current);
    current += next;
    ASSERT_EQ(current, 10); // At 's'
    
    // "snake"
    next = CalculateNextSubword(ptr + current, len - current);
    current += next;
    ASSERT_EQ(current, 15); // At end
}

int main()
{
    RUN_TEST(TestSortLines);
    RUN_TEST(TestSortLinesEdgeCases);
    RUN_TEST(TestIndentLines);
    RUN_TEST(TestIndentLinesVector);
    RUN_TEST(TestIndentUnindentEdgeCases);
    RUN_TEST(TestUnindentLines);
    RUN_TEST(TestUpperCase);
    RUN_TEST(TestLowerCase);
    RUN_TEST(TestCapitalize);
    RUN_TEST(TestCaseConversionEdgeCases);
    RUN_TEST(TestSentenceCase);
    RUN_TEST(TestSubwordNavigation);
    RUN_TEST(TestSubwordNavigationEdgeCases);
    RUN_TEST(TestFindEnclosingBrackets);
    RUN_TEST(TestFindEnclosingBracketsMore);
    RUN_TEST(TestSelectionHierarchy);
    RUN_TEST(TestSelectionHierarchyMultiLine);
    RUN_TEST(TestExpandSelectionToLine);
    RUN_TEST(TestSelectionHierarchyCR);
    RUN_TEST(TestSelectionHierarchyExtended);
    
    std::cout << "All tests passed!" << std::endl;
    return 0;
}
