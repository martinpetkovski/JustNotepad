#include "../../src/PluginInterface.h"
#include <windows.h>
#include <tchar.h>
#include <string>
#include <vector>
#include <commdlg.h>
#include <fstream>
#include <sstream>
#include <thread>
#include <atomic>
#include <richedit.h>

static const char g_HexLookup[] = "0123456789ABCDEF";

void FormatHexLine(DWORD offset, const BYTE* data, DWORD len, std::string& out)
{
    char buf[128];
    char* p = buf;

    // Offset: %08X
    *p++ = g_HexLookup[(offset >> 28) & 0xF];
    *p++ = g_HexLookup[(offset >> 24) & 0xF];
    *p++ = g_HexLookup[(offset >> 20) & 0xF];
    *p++ = g_HexLookup[(offset >> 16) & 0xF];
    *p++ = g_HexLookup[(offset >> 12) & 0xF];
    *p++ = g_HexLookup[(offset >> 8) & 0xF];
    *p++ = g_HexLookup[(offset >> 4) & 0xF];
    *p++ = g_HexLookup[offset & 0xF];
    *p++ = ' ';
    *p++ = ' ';

    // Hex bytes
    for (DWORD i = 0; i < 16; i++)
    {
        if (i < len)
        {
            BYTE b = data[i];
            *p++ = g_HexLookup[(b >> 4) & 0xF];
            *p++ = g_HexLookup[b & 0xF];
            *p++ = ' ';
        }
        else
        {
            *p++ = ' ';
            *p++ = ' ';
            *p++ = ' ';
        }
        
        if (i == 7) *p++ = ' ';
    }

    *p++ = ' ';
    *p++ = '|';

    // ASCII
    for (DWORD i = 0; i < len; i++)
    {
        unsigned char c = data[i];
        *p++ = (c >= 32 && c <= 126) ? c : '.';
    }
    
    *p++ = '|';
    *p++ = '\r';
    *p++ = '\n';
    
    out.append(buf, p - buf);
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

// Track which files are in Hex mode
// Since we only have one editor in this app, we can just use a global flag
// But to be safe, let's map HWND to bool
// Actually, the app is SDI, so one HWND.
static bool g_bHexMode = false;

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Hex Editor";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"View and edit files in Hex format.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.0";
    }

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = 0;
        return NULL;
    }

    PLUGIN_API void OnFileEvent(const wchar_t* filePath, HWND hEditor, const wchar_t* eventType) {
        if (_wcsicmp(eventType, L"Loaded") == 0) {
            g_bHexMode = false;
            
            // Check if file is binary
            HANDLE hFile = CreateFile(filePath, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
            if (hFile == INVALID_HANDLE_VALUE) return;

            DWORD size = GetFileSize(hFile, NULL);
            if (size == 0) {
                CloseHandle(hFile);
                return;
            }

            // Read first 1KB to check for binary
            BYTE buf[1024];
            DWORD read;
            if (ReadFile(hFile, buf, sizeof(buf), &read, NULL) && read > 0) {
                bool isBinary = false;
                for (DWORD i = 0; i < read; i++) {
                    BYTE b = buf[i];
                    if (b < 0x20 && b != 0x09 && b != 0x0A && b != 0x0D) {
                        isBinary = true;
                        break;
                    }
                }

                if (isBinary) {
                    g_bHexMode = true;
                    
                    // Re-read full file and convert to hex
                    SetFilePointer(hFile, 0, NULL, FILE_BEGIN);
                    std::vector<BYTE> buffer(size);
                    ReadFile(hFile, buffer.data(), size, &read, NULL);
                    
                    std::string hexOutput;
                    hexOutput.reserve(size * 4 + size / 16 * 10);

                    for (DWORD i = 0; i < size; i += 16) {
                        DWORD chunk = min(16, size - i);
                        FormatHexLine(i, buffer.data() + i, chunk, hexOutput);
                    }

                    // Convert to wide char
                    int wlen = MultiByteToWideChar(CP_ACP, 0, hexOutput.c_str(), -1, NULL, 0);
                    std::vector<wchar_t> wbuf(wlen + 1);
                    MultiByteToWideChar(CP_ACP, 0, hexOutput.c_str(), -1, wbuf.data(), wlen);
                    wbuf[wlen] = 0;

                    SetWindowText(hEditor, wbuf.data());
                    
                    // Reset dirty flag since we just loaded it
                    SendMessage(hEditor, EM_SETMODIFY, FALSE, 0);
                }
            }
            CloseHandle(hFile);
        }
    }

    PLUGIN_API bool OnSaveFile(const wchar_t* filePath, HWND hEditor) {
        if (!g_bHexMode) return false;

        int len = GetWindowTextLength(hEditor);
        std::vector<wchar_t> wbuf(len + 1);
        GetWindowText(hEditor, wbuf.data(), len + 1);

        // Convert back to multibyte
        int mlen = WideCharToMultiByte(CP_ACP, 0, wbuf.data(), -1, NULL, 0, NULL, NULL);
        std::vector<char> mbuf(mlen);
        WideCharToMultiByte(CP_ACP, 0, wbuf.data(), -1, mbuf.data(), mlen, NULL, NULL);

        std::string content(mbuf.data());
        std::vector<BYTE> binaryData;
        
        // Split lines
        std::string line;
        std::stringstream ss(content);
        while (std::getline(ss, line))
        {
            if (!line.empty() && line.back() == '\r') line.pop_back();
            ParseHexLine(line, binaryData);
        }

        HANDLE hFile = CreateFile(filePath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE)
        {
            DWORD written;
            WriteFile(hFile, binaryData.data(), (DWORD)binaryData.size(), &written, NULL);
            CloseHandle(hFile);
            return true; // Handled
        }
        return false;
    }
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    return TRUE;
}
