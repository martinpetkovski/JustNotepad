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
    std::basic_string<TCHAR> expected = S("a\r\nb\r\nc\r\n");
    ASSERT_EQ(SortLines(input), expected);
}

TEST(TestJoinLines)
{
    std::basic_string<TCHAR> input = S("Hello\r\nWorld");
    std::basic_string<TCHAR> expected = S("Hello World");
    ASSERT_EQ(JoinLines(input), expected);
}

TEST(TestIndentLines)
{
    std::basic_string<TCHAR> input = S("Line1\r\nLine2");
    std::basic_string<TCHAR> expected = S("\tLine1\r\n\tLine2\r\n");
    ASSERT_EQ(IndentLines(input), expected);
}

TEST(TestUnindentLines)
{
    std::basic_string<TCHAR> input = S("\tLine1\r\n    Line2\r\n Line3");
    std::basic_string<TCHAR> expected = S("Line1\r\nLine2\r\nLine3\r\n");
    ASSERT_EQ(UnindentLines(input), expected);
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

TEST(TestFormatHexLine)
{
    BYTE data[] = { 0x48, 0x65, 0x6C, 0x6C, 0x6F }; // Hello
    std::string out;
    FormatHexLine(0, data, 5, out);
    
    // Expected format: "00000000  48 65 6C 6C 6F                                     |Hello|\r\n"
    // Let's verify start and end or full string.
    // 00000000  48 65 6C 6C 6F                                     |Hello|
    // 8 chars offset + 2 spaces = 10
    // 16 bytes * 3 chars = 48 chars.
    // + " |" = 2 chars
    // + 16 chars ascii
    // + "|\r\n" = 3 chars
    
    // My FormatHexLine implementation:
    // pos += sprintf(buf + pos, "%08X  ", offset); // 10 chars
    // Loop 16 times:
    //   if i < len: "%02X " (3 chars)
    //   else: "   " (3 chars)
    //   if i == 7: " " (1 char)
    // Total hex part: 16*3 + 1 = 49 chars?
    // Let's check code:
    /*
    for (DWORD i = 0; i < 16; i++)
    {
        if (i < len)
            pos += sprintf(buf + pos, "%02X ", data[i]);
        else
            pos += sprintf(buf + pos, "   ");
        
        if (i == 7) pos += sprintf(buf + pos, " ");
    }
    */
    // 16 * 3 = 48. Plus extra space at 7. Total 49.
    // Then " |" = 2 chars.
    // Then 16 chars (ascii).
    // Then "|\r\n" = 3 chars.
    // Total length: 10 + 49 + 2 + 16 + 3 = 80 chars?
    
    // Let's just check if it contains "Hello" and "48 65".
    
    bool containsHex = out.find("48 65 6C 6C 6F") != std::string::npos;
    bool containsAscii = out.find("|Hello|") != std::string::npos;
    
    if (!containsHex || !containsAscii)
    {
        std::cerr << "FAILED TestFormatHexLine\nOutput: " << out << std::endl;
        exit(1);
    }
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

TEST(TestParseHexLine)
{
    std::string line = "00000000  48 65 6C 6C 6F                                     |Hello|\r\n";
    std::vector<BYTE> bytes;
    bool res = ParseHexLine(line, bytes);
    
    if (!res) { std::cerr << "ParseHexLine failed" << std::endl; exit(1); }
    if (bytes.size() != 5) { std::cerr << "ParseHexLine size mismatch" << std::endl; exit(1); }
    if (bytes[0] != 0x48) { std::cerr << "ParseHexLine content mismatch" << std::endl; exit(1); }
}

TEST(TestFindEnclosingBrackets)
{
    std::basic_string<TCHAR> text = S("func(arg1, arg2)");
    // Select "arg1" -> 5, 9
    Range r = FindEnclosingBrackets(text.c_str(), (int)text.length(), 5, 9);
    ASSERT_EQ(r.start, 4L); // (
    ASSERT_EQ(r.end, 16L); // ) + 1
}

int main()
{
    RUN_TEST(TestSortLines);
    RUN_TEST(TestJoinLines);
    RUN_TEST(TestIndentLines);
    RUN_TEST(TestUnindentLines);
    RUN_TEST(TestUpperCase);
    RUN_TEST(TestLowerCase);
    RUN_TEST(TestCapitalize);
    RUN_TEST(TestSentenceCase);
    RUN_TEST(TestFormatHexLine);
    RUN_TEST(TestSubwordNavigation);
    RUN_TEST(TestParseHexLine);
    RUN_TEST(TestFindEnclosingBrackets);
    
    std::cout << "All tests passed!" << std::endl;
    return 0;
}
