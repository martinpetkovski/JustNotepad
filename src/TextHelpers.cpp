#include "TextHelpers.h"
#include <sstream>
#include <algorithm>
#include <vector>

std::vector<std::basic_string<TCHAR>> SplitLines(const std::basic_string<TCHAR>& input)
{
    std::vector<std::basic_string<TCHAR>> lines;
    std::basic_string<TCHAR> current;
    for (size_t i = 0; i < input.length(); i++)
    {
        if (input[i] == '\r')
        {
            lines.push_back(current);
            current.clear();
            if (i + 1 < input.length() && input[i+1] == '\n') i++;
        }
        else if (input[i] == '\n')
        {
            lines.push_back(current);
            current.clear();
        }
        else
        {
            current += input[i];
        }
    }
    lines.push_back(current);
    return lines;
}

std::basic_string<TCHAR> SortLines(const std::basic_string<TCHAR>& input)
{
    bool endsWithNewline = false;
    if (!input.empty())
    {
        if (input.back() == '\n') endsWithNewline = true;
    }

    std::vector<std::basic_string<TCHAR>> lines = SplitLines(input);
    
    // Separate non-empty and empty/whitespace-only lines so we can
    // retain the exact number of trailing empty lines after sorting.
    std::vector<std::basic_string<TCHAR>> nonEmpty;
    int emptyCount = 0;
    nonEmpty.reserve(lines.size());
    for (const auto& line : lines)
    {
        bool allWs = true;
        for (TCHAR ch : line)
        {
            if (!(ch == ' ' || ch == '\t')) { allWs = false; break; }
        }
        if (!line.empty() && !allWs)
            nonEmpty.push_back(line);
        else
            emptyCount++;
    }
    
    std::sort(nonEmpty.begin(), nonEmpty.end());
    
    // Rebuild: sorted non-empty lines followed by the original count of empty lines
    std::vector<std::basic_string<TCHAR>> result;
    result.reserve(nonEmpty.size() + emptyCount);
    for (auto& s : nonEmpty) result.push_back(s);
    for (int i = 0; i < emptyCount; ++i) result.push_back(std::basic_string<TCHAR>());
    
    std::basic_string<TCHAR> out;
    // Emit sorted lines and preserved empty lines; keep trailing newline behavior
    for (size_t i = 0; i < result.size(); i++)
    {
        out += result[i];
        if (i < result.size() - 1 || endsWithNewline)
        {
            out += _T("\r\n");
        }
    }
    return out;
}

// JoinLines removed

std::basic_string<TCHAR> IndentLines(const std::basic_string<TCHAR>& input)
{
    bool endsWithNewline = false;
    if (!input.empty())
    {
        if (input.back() == '\n') endsWithNewline = true;
    }

    std::vector<std::basic_string<TCHAR>> lines = SplitLines(input);
    std::basic_string<TCHAR> out;
    for (size_t i = 0; i < lines.size(); i++)
    {
        out += _T("\t");
        out += lines[i];
        if (i < lines.size() - 1 || endsWithNewline)
        {
            out += _T("\r\n");
        }
    }
    return out;
}

std::basic_string<TCHAR> UnindentLines(const std::basic_string<TCHAR>& input)
{
    bool endsWithNewline = false;
    if (!input.empty())
    {
        if (input.back() == '\n') endsWithNewline = true;
    }

    std::vector<std::basic_string<TCHAR>> lines = SplitLines(input);
    std::basic_string<TCHAR> out;
    for (size_t i = 0; i < lines.size(); i++)
    {
        auto& line = lines[i];
        if (!line.empty())
        {
            if (line[0] == '\t') line.erase(0, 1);
            else if (line.size() >= 4 && line.substr(0, 4) == _T("    ")) line.erase(0, 4);
            else if (line[0] == ' ') line.erase(0, 1);
        }
        out += line;
        if (i < lines.size() - 1 || endsWithNewline)
        {
            out += _T("\r\n");
        }
    }
    return out;
}

std::basic_string<TCHAR> ToUpperCase(const std::basic_string<TCHAR>& input)
{
    std::basic_string<TCHAR> out = input;
    if (!out.empty())
    {
        CharUpperBuff(&out[0], (DWORD)out.length());
    }
    return out;
}

std::basic_string<TCHAR> ToLowerCase(const std::basic_string<TCHAR>& input)
{
    std::basic_string<TCHAR> out = input;
    if (!out.empty())
    {
        CharLowerBuff(&out[0], (DWORD)out.length());
    }
    return out;
}

std::basic_string<TCHAR> ToCapitalize(const std::basic_string<TCHAR>& input)
{
    std::basic_string<TCHAR> out = input;
    BOOL bNewWord = TRUE;
    for (size_t i = 0; i < out.length(); i++)
    {
        if (IsCharAlpha(out[i]))
        {
            if (bNewWord)
            {
                out[i] = (TCHAR)CharUpper((LPTSTR)(LONG_PTR)out[i]);
                bNewWord = FALSE;
            }
            else
            {
                out[i] = (TCHAR)CharLower((LPTSTR)(LONG_PTR)out[i]);
            }
        }
        else
        {
            bNewWord = TRUE;
        }
    }
    return out;
}

std::basic_string<TCHAR> ToSentenceCase(const std::basic_string<TCHAR>& input)
{
    std::basic_string<TCHAR> out = input;
    BOOL bNewSentence = TRUE;
    for (size_t i = 0; i < out.length(); i++)
    {
        if (IsCharAlpha(out[i]))
        {
            if (bNewSentence)
            {
                out[i] = (TCHAR)CharUpper((LPTSTR)(LONG_PTR)out[i]);
                bNewSentence = FALSE;
            }
            else
            {
                out[i] = (TCHAR)CharLower((LPTSTR)(LONG_PTR)out[i]);
            }
        }
        else if (out[i] == '.' || out[i] == '!' || out[i] == '?')
        {
            bNewSentence = TRUE;
        }
    }
    return out;
}

void FormatHexLine(DWORD offset, const BYTE* data, DWORD len, std::string& out)
{
    char buf[256];
    int pos = 0;
    pos += sprintf(buf + pos, "%08X  ", offset);
    
    for (DWORD i = 0; i < 16; i++)
    {
        if (i < len)
            pos += sprintf(buf + pos, "%02X ", data[i]);
        else
            pos += sprintf(buf + pos, "   ");
        
        if (i == 7) pos += sprintf(buf + pos, " ");
    }
    
    pos += sprintf(buf + pos, " |");
    for (DWORD i = 0; i < len; i++)
    {
        unsigned char c = data[i];
        pos += sprintf(buf + pos, "%c", (c >= 32 && c <= 126) ? c : '.');
    }
    pos += sprintf(buf + pos, "|\r\n");
    out.append(buf);
}

bool ParseHexLine(const std::string& line, std::vector<BYTE>& bytes)
{
    // Format: XXXXXXXX  HH HH ...
    if (line.length() <= 10) return false;

    // Skip offset and spaces
    size_t hexStart = 10;
    
    // Read up to 16 bytes
    for (int i = 0; i < 16; i++)
    {
        // Each byte is 3 chars: "HH "
        if (hexStart + 2 >= line.length()) break;
        
        char h1 = line[hexStart];
        char h2 = line[hexStart + 1];
        
        if (h1 == ' ' || h1 == '|') break; // End of hex part
        
        // Simple hex parse
        auto fromHex = [](char c) -> int {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'A' && c <= 'F') return c - 'A' + 10;
            if (c >= 'a' && c <= 'f') return c - 'a' + 10;
            return 0;
        };
        
        BYTE b = (fromHex(h1) << 4) | fromHex(h2);
        bytes.push_back(b);
        
        hexStart += 3;
        if (i == 7) hexStart++; // Extra space
    }
    return !bytes.empty();
}

static int GetCharType(TCHAR c) {
    if (c >= 'a' && c <= 'z') return 1;
    if (c >= 'A' && c <= 'Z') return 2;
    if (c >= '0' && c <= '9') return 3;
    if (c == '_' || c == '-') return 4;
    return 0;
}

int CalculateNextSubword(const TCHAR* text, int length)
{
    if (length == 0) return 0;
    int i = 0;
    int startType = GetCharType(text[0]);
    i++;
    while (i < length)
    {
        int type = GetCharType(text[i]);
        if (type != startType)
        {
            if (startType == 1 && type == 2) break; // lower -> Upper
            if (type == 4) break; // -> _ or -
            if (startType == 4) break; // _ or - ->
            if (startType != 0 && type == 0) break; // word -> non-word
            if (startType == 0 && type != 0) break; // non-word -> word
            
            if (startType == 2 && type == 2) {} 
            else startType = type;
        }
        i++;
    }
    return i;
}

int CalculatePrevSubword(const TCHAR* text, int length)
{
    if (length == 0) return 0;
    int i = length - 1;
    int startType = GetCharType(text[i]);
    i--;
    while (i >= 0)
    {
        int type = GetCharType(text[i]);
        if (type != startType)
        {
            if (startType == 2 && type == 1) { i++; break; } // Upper -> lower (backwards)
            if (type == 4) { i++; break; }
            if (startType == 4) { i++; break; }
            if (startType != 0 && type == 0) { i++; break; }
            if (startType == 0 && type != 0) { i++; break; }
            startType = type;
        }
        i--;
    }
    if (i < 0) i = 0;
    return i;
}

Range FindEnclosingBrackets(const TCHAR* text, int length, long selStart, long selEnd)
{
    Range result = { -1, -1 };
    
    // Map selection to buffer indices
    // selStart and selEnd are relative to the start of 'text'
    
    const TCHAR* pairs = _T("()[]{}<>\"\"''");
    int bestDist = -1;

    for (int i = 0; pairs[i]; i += 2)
    {
        TCHAR open = pairs[i];
        TCHAR close = pairs[i+1];
        
        // Search backwards for open
        long p1 = selStart - 1;
        int balance = 0;
        while (p1 >= 0)
        {
            if (text[p1] == close && open != close) balance++;
            else if (text[p1] == open)
            {
                if (balance == 0) break;
                balance--;
            }
            p1--;
        }
        
        if (p1 >= 0)
        {
            // Search forwards for close
            long p2 = selEnd;
            balance = 0;
            while (p2 < length)
            {
                if (text[p2] == open && open != close) balance++;
                else if (text[p2] == close)
                {
                    if (balance == 0) break;
                    balance--;
                }
                p2++;
            }
            
            if (p2 < length)
            {
                long realStart = p1;
                long realEnd = p2 + 1;
                
                if (realStart <= selStart && realEnd >= selEnd && (realStart < selStart || realEnd > selEnd))
                {
                    int dist = (selStart - realStart) + (realEnd - selEnd);
                    if (bestDist == -1 || dist < bestDist)
                    {
                        bestDist = dist;
                        result.start = realStart;
                        result.end = realEnd;
                    }
                }
            }
        }
    }
    return result;
}

Range GetSelectionRange(const TCHAR* text, int length, long selStart, long selEnd, SelectionLevel level)
{
    Range r = { selStart, selEnd };
    if (length == 0) return r;

    switch (level)
    {
    case SEL_CHARACTER:
        if (selStart == selEnd)
        {
            if (selStart < length) r.end++;
        }
        break;
        
    case SEL_SUBWORD:
        {
            long p1 = selStart;
            if (p1 > 0)
            {
                int prev = CalculatePrevSubword(text, p1);
                p1 = prev;
            }
            
            long p2 = selEnd;
            if (p2 < length)
            {
                int next = CalculateNextSubword(text + p2, length - p2);
                p2 += next;
            }
            r.start = p1;
            r.end = p2;
        }
        break;
        
    case SEL_WORD:
        {
            auto IsWordChar = [](TCHAR c) { return IsCharAlphaNumeric(c) || c == '_'; };
            
            long p1 = selStart;
            while (p1 > 0 && IsWordChar(text[p1-1])) p1--;
            
            long p2 = selEnd;
            while (p2 < length && IsWordChar(text[p2])) p2++;
            
            r.start = p1;
            r.end = p2;
        }
        break;
        
    case SEL_PHRASE:
        {
            auto IsPhraseSep = [](TCHAR c) { return c == ',' || c == ';' || c == ':' || c == '-' || c == 0x2013 || c == 0x2014 || c == '.' || c == '!' || c == '?'; };
            
            long p1 = selStart;
            while (p1 > 0 && !IsPhraseSep(text[p1-1]) && text[p1-1] != '\n' && text[p1-1] != '\r') p1--;
            
            long p2 = selEnd;
            while (p2 < length && !IsPhraseSep(text[p2]) && text[p2] != '\n' && text[p2] != '\r') p2++;
            
            r.start = p1;
            r.end = p2;
        }
        break;
        
    case SEL_SENTENCE:
        {
            auto IsSentenceSep = [](TCHAR c) { return c == '.' || c == '!' || c == '?'; };
            auto IsAbbrev = [&](long pos) {
                if (pos <= 0) return false;
                long start = pos - 1;
                while (start >= 0 && IsCharAlpha(text[start])) start--;
                start++; // First char of word
                if (start >= pos) return false;
                
                std::basic_string<TCHAR> word(text + start, pos - start);
                static const TCHAR* abbrevs[] = { 
                    _T("Mr"), _T("Mrs"), _T("Ms"), _T("Dr"), _T("Prof"), _T("Sr"), _T("Jr"), _T("St"), 
                    _T("vs"), _T("etc"), _T("eg"), _T("ie"), _T("Fig")
                };
                for (const auto* a : abbrevs) {
                    if (_tcsicmp(word.c_str(), a) == 0) return true;
                }
                return false;
            };

            long p1 = selStart;
            while (p1 > 0)
            {
                TCHAR c = text[p1-1];
                if (IsSentenceSep(c))
                {
                    if (c != '.' || !IsAbbrev(p1-1)) break;
                }
                if (c == '\n' || c == '\r') 
                {
                    // Stop at paragraph breaks (double newline) or just newline?
                    // User said "Line" is next level. So Sentence should probably be within a line or paragraph.
                    // Let's stop at any newline to be safe and consistent with "Line" being larger.
                    // If we stop at newline, a sentence cannot span lines.
                    // But "Line" level selects the *whole* line.
                    // If a sentence is part of a line, it's smaller.
                    // If a sentence spans 2 lines, it's larger than 1 line?
                    // Let's stick to the previous logic: stop at newline.
                    break; 
                }
                p1--;
            }
            
            // Skip leading whitespace
            while (p1 < selStart && (text[p1] == ' ' || text[p1] == '\t')) p1++;
            
            long p2 = selEnd;
            while (p2 < length)
            {
                TCHAR c = text[p2];
                if (IsSentenceSep(c))
                {
                    if (c != '.' || !IsAbbrev(p2)) 
                    {
                        p2++; // Include separator
                        break;
                    }
                }
                if (c == '\n' || c == '\r') break;
                p2++;
            }
            
            r.start = p1;
            r.end = p2;
        }
        break;
        
    case SEL_LINE:
        {
            long p1 = selStart;
            while (p1 > 0 && text[p1-1] != '\n' && text[p1-1] != '\r') p1--;
            
            long p2 = selEnd;
            // If we have a selection and we are at the start of a line (meaning we ended at the previous newline),
            // do not expand to the next line.
            bool atLineStart = (p2 > 0 && (text[p2-1] == '\n' || text[p2-1] == '\r'));
            if (p2 > p1 && atLineStart)
            {
                // Already at end of line
            }
            else
            {
                while (p2 < length && text[p2] != '\r' && text[p2] != '\n') p2++;
                if (p2 < length && text[p2] == '\r') p2++;
                if (p2 < length && text[p2] == '\n') p2++;
            }
            
            r.start = p1;
            r.end = p2;
        }
        break;
        
    case SEL_PARAGRAPH:
        {
            long p1 = selStart;
            while (p1 > 0)
            {
                if (text[p1] == '\n' || text[p1] == '\r')
                {
                    if (p1 > 1 && text[p1-1] == '\r' && text[p1-2] == '\n') { p1++; break; } // \n\r? Unlikely.
                    // Check for double newline
                    // \r\n\r\n
                    // \n\n
                    // \r\r
                    
                    // If we found a newline, check if previous was also newline
                    bool isNewline = (text[p1] == '\n' || text[p1] == '\r'); // Wait, p1 is index.
                    // We are scanning backwards.
                    // If text[p1] is newline? No, we check text[p1].
                    // The loop decrements p1.
                    // Let's rewrite to be clearer.
                }
                p1--;
            }
            // Re-implementing Paragraph scan to be robust with \r
            p1 = selStart;
            while (p1 > 0)
            {
                // Check for blank line before p1
                // A blank line is \n\n or \r\n\r\n or \r\r
                
                // Let's look for 2 consecutive line breaks.
                // But we are going backwards.
                
                // Simplified: Stop if we see 2 newlines sequences.
                // Or just stop at the start of the paragraph.
                // Paragraph start is after a blank line.
                
                // If we are at p1.
                // If p1-1 is \n or \r.
                // And p1-2 (or p1-3) is also \n or \r.
                
                bool isLF = (text[p1-1] == '\n');
                bool isCR = (text[p1-1] == '\r');
                
                if (isLF || isCR)
                {
                    // Found one newline. Check for another before it.
                    long prev = p1 - 1;
                    if (isLF && prev > 0 && text[prev-1] == '\r') prev--; // Skip \r of \r\n
                    
                    if (prev > 0)
                    {
                        if (text[prev-1] == '\n' || text[prev-1] == '\r')
                        {
                            // Found double newline. p1 is the start of the paragraph.
                            break;
                        }
                    }
                }
                p1--;
            }
            
            long p2 = selEnd;
            
            // Check if we are already at the end of a paragraph (start of blank line)
            bool atParagraphEnd = false;
            if (p2 > p1 && p2 < length && (text[p2-1] == '\n' || text[p2-1] == '\r'))
            {
                // Check if next is newline
                if (text[p2] == '\r' || text[p2] == '\n') atParagraphEnd = true;
            }

            if (!atParagraphEnd)
            {
                while (p2 < length)
                {
                    if (text[p2] == '\r' || text[p2] == '\n')
                    {
                        // Found newline. Check if next is also newline.
                        long next = p2 + 1;
                        if (text[p2] == '\r' && next < length && text[next] == '\n') next++;
                        
                        if (next < length && (text[next] == '\r' || text[next] == '\n'))
                        {
                            // Double newline found. p2 should include the first newline sequence but stop before the second?
                            // "Paragraph/block (group of lines separated by blank lines)"
                            // Usually includes the newline of the last line.
                            p2 = next; 
                            break; 
                        }
                    }
                    p2++;
                }
            }
            r.start = p1;
            r.end = p2;
        }
        break;
        
    case SEL_SECTION:
        {
            long p1 = selStart;
            while (p1 > 0)
            {
                long lineStart = p1;
                while (lineStart > 0 && text[lineStart-1] != '\n' && text[lineStart-1] != '\r') lineStart--;
                
                if (text[lineStart] == '#') { p1 = lineStart; break; }
                if (length - lineStart >= 7 && _tcsncmp(text + lineStart, _T("Chapter"), 7) == 0) { p1 = lineStart; break; }
                
                p1 = lineStart - 1;
            }
            if (p1 < 0) p1 = 0;
            
            long p2 = selEnd;
            
            // Check if we are at start of a new section
            bool atSectionEnd = false;
            if (p2 > p1 && p2 < length && (text[p2-1] == '\n' || text[p2-1] == '\r'))
            {
                if (text[p2] == '#') atSectionEnd = true;
                else if (length - p2 >= 7 && _tcsncmp(text + p2, _T("Chapter"), 7) == 0) atSectionEnd = true;
            }

            if (!atSectionEnd)
            {
                while (p2 < length)
                {
                    if (text[p2] == '\n' || text[p2] == '\r')
                    {
                        long nextLine = p2 + 1;
                        if (text[p2] == '\r' && nextLine < length && text[nextLine] == '\n') nextLine++;
                        
                        if (nextLine < length)
                        {
                            if (text[nextLine] == '#') { p2 = nextLine; break; }
                            if (length - nextLine >= 7 && _tcsncmp(text + nextLine, _T("Chapter"), 7) == 0) { p2 = nextLine; break; }
                        }
                    }
                    p2++;
                }
            }
            r.start = p1;
            r.end = p2;
        }
        break;
        
    case SEL_DOCUMENT:
        r.start = 0;
        r.end = length;
        break;
    }
    
    if (r.start < 0) r.start = 0;
    if (r.end > length) r.end = length;
    return r;
}
