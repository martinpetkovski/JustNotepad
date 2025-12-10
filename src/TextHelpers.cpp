#include "TextHelpers.h"
#include <sstream>
#include <algorithm>
#include <vector>

std::basic_string<TCHAR> SortLines(const std::basic_string<TCHAR>& input)
{
    std::vector<std::basic_string<TCHAR>> lines;
    std::basic_stringstream<TCHAR> ss(input);
    std::basic_string<TCHAR> line;
    while (std::getline(ss, line))
    {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        lines.push_back(line);
    }
    
    std::sort(lines.begin(), lines.end());
    
    std::basic_string<TCHAR> out;
    for (size_t i = 0; i < lines.size(); i++)
    {
        out += lines[i];
        out += _T("\r\n");
    }
    // Note: This always adds a newline at the end, even if input didn't have one.
    // For strict correctness we might want to check input, but for this feature it's acceptable.
    return out;
}

std::basic_string<TCHAR> JoinLines(const std::basic_string<TCHAR>& input)
{
    std::basic_string<TCHAR> out;
    for (size_t i = 0; i < input.length(); i++)
    {
        if (input[i] == '\r' || input[i] == '\n')
        {
            if (out.empty() || out.back() != ' ') out += ' ';
        }
        else
        {
            out += input[i];
        }
    }
    return out;
}

std::basic_string<TCHAR> IndentLines(const std::basic_string<TCHAR>& input)
{
    std::basic_string<TCHAR> out;
    std::basic_stringstream<TCHAR> ss(input);
    std::basic_string<TCHAR> line;
    while (std::getline(ss, line))
    {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        out += _T("\t");
        out += line;
        out += _T("\r\n");
    }
    return out;
}

std::basic_string<TCHAR> UnindentLines(const std::basic_string<TCHAR>& input)
{
    std::basic_string<TCHAR> out;
    std::basic_stringstream<TCHAR> ss(input);
    std::basic_string<TCHAR> line;
    while (std::getline(ss, line))
    {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        if (!line.empty())
        {
            if (line[0] == '\t') line.erase(0, 1);
            else if (line.size() >= 4 && line.substr(0, 4) == _T("    ")) line.erase(0, 4);
            else if (line[0] == ' ') line.erase(0, 1);
        }
        out += line;
        out += _T("\r\n");
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

int CalculateNextSubword(const TCHAR* text, int length)
{
    if (length == 0) return 0;
    
    int i = 0;
    auto GetType = [](TCHAR c) -> int {
        if (c >= 'a' && c <= 'z') return 1;
        if (c >= 'A' && c <= 'Z') return 2;
        if (c >= '0' && c <= '9') return 3;
        if (c == '_') return 4;
        return 0;
    };
    
    int startType = GetType(text[0]);
    i++;
    while (i < length)
    {
        int type = GetType(text[i]);
        if (type != startType)
        {
            if (startType == 1 && type == 2) break; // lower -> Upper
            if (type == 4) break; // -> _
            if (startType == 4) break; // _ ->
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
    auto GetType = [](TCHAR c) -> int {
        if (c >= 'a' && c <= 'z') return 1;
        if (c >= 'A' && c <= 'Z') return 2;
        if (c >= '0' && c <= '9') return 3;
        if (c == '_') return 4;
        return 0;
    };
    
    int startType = GetType(text[i]);
    i--;
    while (i >= 0)
    {
        int type = GetType(text[i]);
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
