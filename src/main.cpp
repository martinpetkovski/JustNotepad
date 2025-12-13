#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commdlg.h>
#include <richedit.h>
#include <commctrl.h>
#include <tchar.h>
#include <strsafe.h>
#include <vector>
#include <string>
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <string_view>
#include <shlwapi.h>
#include <shellapi.h>
#include <shlobj.h>
#include <deque>
#include <dwmapi.h>
#include <uxtheme.h>
#include <ole2.h>
#include <thread>
#include <atomic>
#include <map>
#include "resource.h"
#include "TextHelpers.h"
#include "PluginManager.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "ole32.lib")

#define IDC_EDITOR 1001
#define IDC_STATUSBAR 1002

enum Encoding {
    ENC_ANSI = 0,
    ENC_UTF8 = 1,
    ENC_UTF16LE = 2,
    ENC_UTF16BE = 3,
    ENC_ASCII,
    ENC_ISO_8859_1,
    ENC_WINDOWS_1252,
    ENC_UTF32,
    ENC_ISO_8859_15,
    ENC_SHIFT_JIS,
    ENC_EUC_JP,
    ENC_BIG5,
    ENC_GB18030,
    ENC_GBK,
    ENC_GB2312,
    ENC_EUC_KR,
    ENC_ISO_8859_2,
    ENC_WINDOWS_1250,
    ENC_KOI8_R,
    ENC_ISO_8859_5,
    ENC_WINDOWS_1251,
    ENC_ISO_8859_6,
    ENC_ISO_8859_7,
    ENC_ISO_8859_8,
    ENC_ISO_8859_9
};

struct EncodingInfo {
    Encoding id;
    UINT codePage;
    const TCHAR* name;
    int cmdId;
};

const EncodingInfo g_Encodings[] = {
    { ENC_ANSI, CP_ACP, _T("ANSI"), ID_ENCODING_ANSI },
    { ENC_UTF8, CP_UTF8, _T("UTF-8"), ID_ENCODING_UTF8 },
    { ENC_UTF16LE, 1200, _T("UTF-16 LE"), ID_ENCODING_UTF16LE },
    { ENC_UTF16BE, 1201, _T("UTF-16 BE"), ID_ENCODING_UTF16BE },
    { ENC_ASCII, 20127, _T("ASCII"), ID_ENCODING_ASCII },
    { ENC_ISO_8859_1, 28591, _T("ISO-8859-1"), ID_ENCODING_ISO_8859_1 },
    { ENC_WINDOWS_1252, 1252, _T("Windows-1252"), ID_ENCODING_WINDOWS_1252 },
    { ENC_UTF32, 12000, _T("UTF-32"), ID_ENCODING_UTF32 }, // Note: RichEdit might need special handling
    { ENC_ISO_8859_15, 28605, _T("ISO-8859-15"), ID_ENCODING_ISO_8859_15 },
    { ENC_SHIFT_JIS, 932, _T("Shift JIS"), ID_ENCODING_SHIFT_JIS },
    { ENC_EUC_JP, 51932, _T("EUC-JP"), ID_ENCODING_EUC_JP },
    { ENC_BIG5, 950, _T("Big5"), ID_ENCODING_BIG5 },
    { ENC_GB18030, 54936, _T("GB18030"), ID_ENCODING_GB18030 },
    { ENC_GBK, 936, _T("GBK"), ID_ENCODING_GBK },
    { ENC_GB2312, 936, _T("GB2312"), ID_ENCODING_GB2312 },
    { ENC_EUC_KR, 51949, _T("EUC-KR"), ID_ENCODING_EUC_KR },
    { ENC_ISO_8859_2, 28592, _T("ISO-8859-2"), ID_ENCODING_ISO_8859_2 },
    { ENC_WINDOWS_1250, 1250, _T("Windows-1250"), ID_ENCODING_WINDOWS_1250 },
    { ENC_KOI8_R, 20866, _T("KOI8-R"), ID_ENCODING_KOI8_R },
    { ENC_ISO_8859_5, 28595, _T("ISO-8859-5"), ID_ENCODING_ISO_8859_5 },
    { ENC_WINDOWS_1251, 1251, _T("Windows-1251"), ID_ENCODING_WINDOWS_1251 },
    { ENC_ISO_8859_6, 28596, _T("ISO-8859-6"), ID_ENCODING_ISO_8859_6 },
    { ENC_ISO_8859_7, 28597, _T("ISO-8859-7"), ID_ENCODING_ISO_8859_7 },
    { ENC_ISO_8859_8, 28598, _T("ISO-8859-8"), ID_ENCODING_ISO_8859_8 },
    { ENC_ISO_8859_9, 28599, _T("ISO-8859-9"), ID_ENCODING_ISO_8859_9 }
};

const EncodingInfo* GetEncodingInfo(Encoding enc) {
    for (const auto& info : g_Encodings) {
        if (info.id == enc) return &info;
    }
    return &g_Encodings[0]; // Default to ANSI
}

const EncodingInfo* GetEncodingInfoByCmd(int cmdId) {
    for (const auto& info : g_Encodings) {
        if (info.cmdId == cmdId) return &info;
    }
    return NULL;
}


struct Document {
    HWND hEditor;
    TCHAR szFileName[MAX_PATH];
    BOOL bIsDirty;
    BOOL bLoading;
    BOOL bPinned;
    FILETIME ftLastWriteTime;
    Encoding currentEncoding;
    long loadSequence;
    DWORD fileSize;
    BOOL isBinary;
};

HWND hMain;
HWND hEditor;
HWND hStatus;
HINSTANCE hInst;
HWND hFindReplaceDlg = NULL;
UINT uFindReplaceMsg = 0;
FINDREPLACE fr;
TCHAR szFindWhat[256];
TCHAR szReplaceWith[256];
BOOL bWordWrap = FALSE;
std::vector<CHARRANGE> vecSelectionStack;
std::vector<SelectionLevel> vecSelectionLevelStack;

Document g_Document;
PluginManager g_PluginManager;

// Config
std::deque<std::wstring> recentFiles;
TCHAR szConfigPath[MAX_PATH] = {0};
std::wstring g_StatusText[7]; // Cache for owner-draw status bar

// Encoding currentEncoding = ENC_UTF8; // Moved to Document

#define WM_UPDATE_PROGRESS (WM_USER + 200)
HWND hProgressBarStatus = NULL;

// Removed g_BgLoad and ProgressHelper

struct StreamContext {
    HANDLE hFile;
    std::vector<char> inBuffer;
    size_t inBufferPos;
    DWORD dwOffset;
    std::string outBuffer;
    HWND hProgressBar; // For background load
    DWORD dwTotalSize;
    DWORD dwReadSoFar;
    
    // Memory mapping
    HANDLE hMap;
    LPBYTE pData;
    DWORD dwMapPos;
};

// Forward declarations
void UpdateStatusBar();
void UpdateTitle();
void UpdatePluginsMenu(HWND hwnd);
void LoadConfig();
void SaveConfig();
void AddRecentFile(LPCTSTR pszPath);
void UpdateRecentFilesMenu();
void ApplyTheme();
BOOL LoadFromFile(LPCTSTR pszFile);
void DoReload();
void DoIndent();
void DoUnindent();

void CreateNewDocument(LPCTSTR pszFile);

#define WM_LOAD_COMPLETE (WM_USER + 102)
#define WM_LOAD_PROGRESS (WM_USER + 103)

struct LoadResult {
    std::vector<wchar_t> buffer;
    Encoding encoding;
    BOOL bSuccess;
    long sequence;
};

void ParallelConvert(const BYTE* pData, size_t size, Encoding encoding, std::vector<wchar_t>& outBuffer, HWND hMain)
{
    int numThreads = std::thread::hardware_concurrency();
    if (numThreads == 0) numThreads = 2;
    if (size < 1024 * 1024) numThreads = 1;

    std::vector<std::thread> threads;
    std::vector<std::vector<wchar_t>> chunks(numThreads);
    std::atomic<size_t> processedBytes(0);
    
    size_t chunkSize = size / numThreads;
    
    for(int i=0; i<numThreads; i++)
    {
        size_t start = i * chunkSize;
        size_t end = (i == numThreads - 1) ? size : (i + 1) * chunkSize;
        
        if (encoding == ENC_UTF8 && i > 0)
        {
            while(start < end && (pData[start] & 0xC0) == 0x80) start++;
        }
        
        threads.emplace_back([=, &chunks, &pData, &processedBytes]() {
            if (start >= end) return;
            
            size_t current = start;
            size_t step = 1024 * 1024; // 1MB chunks for progress reporting
            
            while(current < end)
            {
                size_t blockEnd = min(current + step, end);
                
                // Adjust blockEnd for UTF-8 boundary if needed
                if (encoding == ENC_UTF8 && blockEnd < end) {
                    while(blockEnd < end && (pData[blockEnd] & 0xC0) == 0x80) blockEnd++;
                }
                
                size_t len = blockEnd - current;
                const EncodingInfo* info = GetEncodingInfo(encoding);
                UINT cp = info->codePage;
                
                int needed = MultiByteToWideChar(cp, 0, (LPCSTR)(pData + current), (int)len, NULL, 0);
                if (needed > 0) {
                    size_t oldSize = chunks[i].size();
                    chunks[i].resize(oldSize + needed);
                    MultiByteToWideChar(cp, 0, (LPCSTR)(pData + current), (int)len, chunks[i].data() + oldSize, needed);
                }
                
                current = blockEnd;
                processedBytes += len;
                
                // Report progress
                PostMessage(hMain, WM_LOAD_PROGRESS, (WPARAM)processedBytes.load(), 0);
            }
        });
    }
    
    for(auto& t : threads) t.join();
    
    size_t total = 0;
    for(const auto& c : chunks) total += c.size();
    outBuffer.reserve(total + 1);
    
    for(const auto& c : chunks) outBuffer.insert(outBuffer.end(), c.begin(), c.end());
    outBuffer.push_back(0);
}

void LoadWorker(std::wstring filename, Encoding encoding, HWND hMain, long sequence, BOOL isBinary)
{
    HANDLE hFile = CreateFile(filename.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return;
    
    DWORD size = GetFileSize(hFile, NULL);
    
    HANDLE hMap = CreateFileMapping(hFile, NULL, PAGE_READONLY, 0, 0, NULL);
    if (!hMap) { CloseHandle(hFile); return; }
    
    const BYTE* pData = (const BYTE*)MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, 0);
    if (!pData) { CloseHandle(hMap); CloseHandle(hFile); return; }
    
    LoadResult* res = new LoadResult();
    res->encoding = encoding;
    res->bSuccess = TRUE;
    res->sequence = sequence;
    
    if (isBinary) {
        std::wstring msg = L"This file is binary and cannot be displayed in the text editor.\r\nUse the Hex Editor plugin to view the content.";
        res->buffer.assign(msg.begin(), msg.end());
        res->buffer.push_back(0);
    }
    else if (encoding == ENC_UTF16LE) {
        size_t chars = size / 2;
        res->buffer.resize(chars + 1);
        
        // Chunked copy for progress
        size_t copied = 0;
        size_t step = 1024 * 1024; // 1MB
        while(copied < size) {
            size_t chunk = min(step, size - copied);
            memcpy((BYTE*)res->buffer.data() + copied, pData + copied, chunk);
            copied += chunk;
            PostMessage(hMain, WM_LOAD_PROGRESS, (WPARAM)copied, 0);
        }
        res->buffer[chars] = 0;
    } else if (encoding == ENC_UTF16BE) {
        size_t chars = size / 2;
        res->buffer.resize(chars + 1);
        const WORD* src = (const WORD*)pData;
        
        size_t processed = 0;
        size_t step = 512 * 1024; // 512K chars
        
        while(processed < chars) {
            size_t chunk = min(step, chars - processed);
            for(size_t i=0; i<chunk; i++) {
                res->buffer[processed + i] = _byteswap_ushort(src[processed + i]);
            }
            processed += chunk;
            PostMessage(hMain, WM_LOAD_PROGRESS, (WPARAM)(processed * 2), 0);
        }
        res->buffer[chars] = 0;
    } else {
        ParallelConvert(pData, size, encoding, res->buffer, hMain);
    }
    
    UnmapViewOfFile(pData);
    CloseHandle(hMap);
    CloseHandle(hFile);
    
    PostMessage(hMain, WM_LOAD_COMPLETE, 0, (LPARAM)res);
}

BOOL GetFileLastWriteTime(LPCTSTR pszFile, FILETIME* pft)
{
    WIN32_FILE_ATTRIBUTE_DATA Data;
    if (GetFileAttributesEx(pszFile, GetFileExInfoStandard, &Data))
    {
        *pft = Data.ftLastWriteTime;
        return TRUE;
    }
    return FALSE;
}

void UpdateStreamProgress(StreamContext* ctx, DWORD bytesAdded) {
    ctx->dwReadSoFar += bytesAdded;
    if (ctx->hProgressBar) {
        // Update every 256KB or if finished
        if ((ctx->dwReadSoFar & 0x3FFFF) < bytesAdded || ctx->dwReadSoFar >= ctx->dwTotalSize) {
             SendMessage(ctx->hProgressBar, PBM_SETPOS, ctx->dwReadSoFar, 0);
             MSG msg;
             while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
                 TranslateMessage(&msg);
                 DispatchMessage(&msg);
             }
        }
    }
}

// Stream callback for reading file
DWORD CALLBACK StreamInCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
    StreamContext* ctx = (StreamContext*)dwCookie;
    *pcb = 0;

    if (ctx->pData) // Memory mapped
    {
        if (ctx->dwMapPos >= ctx->dwTotalSize) return 0; // EOF

        LONG toCopy = min(cb, (LONG)(ctx->dwTotalSize - ctx->dwMapPos));
        memcpy(pbBuff, ctx->pData + ctx->dwMapPos, toCopy);
        ctx->dwMapPos += toCopy;
        *pcb = toCopy;

        UpdateStreamProgress(ctx, toCopy);
        return 0;
    }

    if (ReadFile(ctx->hFile, pbBuff, cb, (LPDWORD)pcb, NULL))
    {
        UpdateStreamProgress(ctx, *pcb);
        return 0;
    }
    return 1; // Error
}

// Stream callback for writing file
DWORD CALLBACK StreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
    StreamContext* ctx = (StreamContext*)dwCookie;
    *pcb = cb; // Assume success for now

    if (WriteFile(ctx->hFile, pbBuff, cb, (LPDWORD)pcb, NULL))
    {
        return 0;
    }
    return 1; // Error
}

BOOL DoFileSave();
BOOL DoFileSaveAs();

BOOL CheckSaveChanges()
{
    Document& doc = g_Document;

    if (!doc.bIsDirty) return TRUE;

    // If untitled and empty, don't prompt
    if (doc.szFileName[0] == 0)
    {
        long nLen = SendMessage(doc.hEditor, WM_GETTEXTLENGTH, 0, 0);
        if (nLen == 0) return TRUE;
    }

    TCHAR szMsg[MAX_PATH + 64];
    if (doc.szFileName[0])
        StringCchPrintf(szMsg, MAX_PATH + 64, _T("Do you want to save changes to %s?"), doc.szFileName);
    else
        StringCchCopy(szMsg, MAX_PATH + 64, _T("Do you want to save changes to Untitled?"));

    int nResult = MessageBox(hMain, szMsg, _T("Just Notepad"), MB_YESNOCANCEL | MB_ICONWARNING);
    if (nResult == IDYES)
    {
        return DoFileSave();
    }
    else if (nResult == IDNO)
    {
        return TRUE;
    }
    else
    {
        return FALSE;
    }
}

BOOL CheckAllSaveChanges()
{
    return CheckSaveChanges();
}

void DoFileNew()
{
    if (CheckSaveChanges())
    {
        CreateNewDocument(NULL);
    }
}

BOOL LoadFromFile(LPCTSTR pszFile)
{
    Document& doc = g_Document;
    doc.loadSequence++;

    HANDLE hFile = CreateFile(pszFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile != INVALID_HANDLE_VALUE)
    {
        DWORD fileSize = GetFileSize(hFile, NULL);
        doc.fileSize = fileSize;
        
        HANDLE hMap = NULL;
        LPBYTE pData = NULL;
        
        if (fileSize > 0)
        {
            hMap = CreateFileMapping(hFile, NULL, PAGE_READONLY, 0, 0, NULL);
            if (hMap)
            {
                pData = (LPBYTE)MapViewOfFile(hMap, FILE_MAP_READ, 0, 0, 0);
            }
        }

        // Check for binary and encoding
        BOOL isBinary = FALSE;
        doc.currentEncoding = ENC_ANSI; // Default fallback

        if (pData && fileSize > 0)
        {
            // Check BOM
            if (fileSize >= 3 && pData[0] == 0xEF && pData[1] == 0xBB && pData[2] == 0xBF)
            {
                doc.currentEncoding = ENC_UTF8;
            }
            else if (fileSize >= 2 && pData[0] == 0xFF && pData[1] == 0xFE)
            {
                doc.currentEncoding = ENC_UTF16LE;
            }
            else if (fileSize >= 2 && pData[0] == 0xFE && pData[1] == 0xFF)
            {
                doc.currentEncoding = ENC_UTF16BE;
            }
            else
            {
                doc.currentEncoding = ENC_UTF8;
                
                // Simple binary check
                DWORD checkLen = min(fileSize, 1024);
                for (DWORD i = 0; i < checkLen; i++)
                {
                    BYTE b = pData[i];
                    if (b < 0x20 && b != 0x09 && b != 0x0A && b != 0x0D)
                    {
                        isBinary = TRUE;
                        break;
                    }
                }
                
                if (isBinary)
                {
                    doc.currentEncoding = ENC_ANSI;
                }
            }
        }
        else if (fileSize > 0 && !pData)
        {
            // Fallback to ReadFile if mapping failed but file is not empty
             BYTE buf[1024];
            DWORD read;
            if (ReadFile(hFile, buf, sizeof(buf), &read, NULL) && read > 0)
            {
                // Check BOM
                if (read >= 3 && buf[0] == 0xEF && buf[1] == 0xBB && buf[2] == 0xBF)
                {
                    doc.currentEncoding = ENC_UTF8;
                }
                else if (read >= 2 && buf[0] == 0xFF && buf[1] == 0xFE)
                {
                    doc.currentEncoding = ENC_UTF16LE;
                }
                else if (read >= 2 && buf[0] == 0xFE && buf[1] == 0xFF)
                {
                    doc.currentEncoding = ENC_UTF16BE;
                }
                else
                {
                    doc.currentEncoding = ENC_UTF8;
                    for (DWORD i = 0; i < read; i++)
                    {
                        BYTE b = buf[i];
                        if (b < 0x20 && b != 0x09 && b != 0x0A && b != 0x0D)
                        {
                            isBinary = TRUE;
                            break;
                        }
                    }
                    
                    if (isBinary)
                    {
                        doc.currentEncoding = ENC_ANSI;
                    }
                }
            }
            SetFilePointer(hFile, 0, NULL, FILE_BEGIN);
        }
        
        if (isBinary)
        {
            // Ensure word wrap is disabled for any future text
            if (bWordWrap)
            {
                bWordWrap = FALSE;
                SendMessage(doc.hEditor, EM_SETTARGETDEVICE, 0, 1);
                HMENU hMenu = GetMenu(hMain);
                CheckMenuItem(hMenu, ID_VIEW_WORDWRAP, MF_UNCHECKED);
            }

            // Do not load any text into the editor. Clear it.
            SendMessage(doc.hEditor, WM_SETTEXT, 0, (LPARAM)L"");

            // Clean up any mapping/handles before showing modal
            if (pData) UnmapViewOfFile(pData);
            if (hMap) CloseHandle(hMap);
            CloseHandle(hFile);

            // Update document state
            StringCchCopy(doc.szFileName, MAX_PATH, pszFile);
            GetFileLastWriteTime(doc.szFileName, &doc.ftLastWriteTime);
            doc.bIsDirty = FALSE;
            doc.bLoading = FALSE;
            doc.isBinary = TRUE;

            // Hide progress bar if visible
            ShowWindow(hProgressBarStatus, SW_HIDE);

            UpdateTitle();
            UpdateStatusBar();

            // Notify plugins – Hex Editor will open its modal window
            g_PluginManager.NotifyFileEvent(doc.szFileName, doc.hEditor, L"Loaded");

            return TRUE;
        }
        
        // Unified Loading Strategy (text files only)
        // Always use background loading path for files > 4KB (approx 1 screen)
        // For smaller files, load synchronously but use progress bar
        
        BOOL bBackgroundLoad = (fileSize > 4096);
        
        // Show Progress Bar
        SendMessage(hProgressBarStatus, PBM_SETRANGE32, 0, fileSize);
        SendMessage(hProgressBarStatus, PBM_SETPOS, 0, 0);
        ShowWindow(hProgressBarStatus, SW_SHOW);

        StreamContext ctx = {0};
        ctx.hFile = hFile;
        ctx.dwOffset = 0;
        ctx.hProgressBar = hProgressBarStatus;
        ctx.dwTotalSize = bBackgroundLoad ? min(4096, fileSize) : fileSize; // Load first 4KB or all
        ctx.dwReadSoFar = 0;
        ctx.hMap = hMap;
        ctx.pData = pData;
        ctx.dwMapPos = 0;
        
        doc.bLoading = TRUE;
        EDITSTREAM es = {0};
        es.dwCookie = (DWORD_PTR)&ctx;
        es.pfnCallback = StreamInCallback;

        DWORD dwFormat = SF_TEXT;
        const EncodingInfo* info = GetEncodingInfo(doc.currentEncoding);
        
        if (doc.currentEncoding == ENC_UTF16LE || doc.currentEncoding == ENC_UTF16BE)
        {
             dwFormat = SF_TEXT | SF_UNICODE;
        }
        else
        {
            dwFormat = SF_TEXT | (info->codePage << 16) | SF_USECODEPAGE;
        }

        SendMessage(doc.hEditor, EM_STREAMIN, dwFormat, (LPARAM)&es);
        
        if (bBackgroundLoad)
        {
            // Close handles used for preview
            if (pData) UnmapViewOfFile(pData);
            if (hMap) CloseHandle(hMap);
            CloseHandle(hFile);
            
            // Start Worker Thread
            std::wstring filename = pszFile;
            Encoding enc = doc.currentEncoding;
            long seq = doc.loadSequence;
            BOOL bin = isBinary;
            std::thread([filename, enc, seq, bin](){
                LoadWorker(filename, enc, hMain, seq, bin);
            }).detach();
            
            // Set ReadOnly while loading
            SendMessage(doc.hEditor, EM_SETREADONLY, TRUE, 0);
        }
        else
        {
            if (pData) UnmapViewOfFile(pData);
            if (hMap) CloseHandle(hMap);
            CloseHandle(hFile);
            doc.bLoading = FALSE;
            ShowWindow(hProgressBarStatus, SW_HIDE);
            
            // Ensure event mask is set so we get EN_CHANGE notifications
            SendMessage(doc.hEditor, EM_SETEVENTMASK, 0, ENM_SELCHANGE | ENM_CHANGE);
        }
        
        StringCchCopy(doc.szFileName, MAX_PATH, pszFile);
        GetFileLastWriteTime(doc.szFileName, &doc.ftLastWriteTime);

        doc.bIsDirty = FALSE;
        if (!bBackgroundLoad) doc.bLoading = FALSE;
        
        // Check ReadOnly
        DWORD dwAttrs = GetFileAttributes(doc.szFileName);
        BOOL bReadOnly = (dwAttrs != INVALID_FILE_ATTRIBUTES) && (dwAttrs & FILE_ATTRIBUTE_READONLY);
        if (!bBackgroundLoad) SendMessage(doc.hEditor, EM_SETREADONLY, bReadOnly, 0);
        HMENU hMenu = GetMenu(hMain);
        CheckMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, bReadOnly ? MF_CHECKED : MF_UNCHECKED);

        UpdateTitle();
        UpdateStatusBar();
        
        // Notify plugins
        g_PluginManager.NotifyFileEvent(doc.szFileName, doc.hEditor, L"Loaded");
        
        return TRUE;
    }
    return FALSE;
}

void OpenFile(LPCTSTR pszFile)
{
    if (CheckSaveChanges())
    {
        CreateNewDocument(pszFile);
        AddRecentFile(pszFile);
    }
}

void DoFileOpen()
{
    OPENFILENAME ofn = {0};
    TCHAR szFile[MAX_PATH] = {0};

    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hMain;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = _T("All Files (*.*)\0*.*\0Text Files (*.txt)\0*.txt\0");
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileName(&ofn))
    {
        OpenFile(ofn.lpstrFile);
    }
}

void DoReload()
{
    Document& doc = g_Document;

    if (!doc.szFileName[0]) return;

    if (doc.bIsDirty)
    {
        if (MessageBox(hMain, _T("You have unsaved changes. Reloading will lose them.\nAre you sure?"), _T("Just Notepad"), MB_YESNO | MB_ICONWARNING) != IDYES)
        {
            return;
        }
    }
    LoadFromFile(doc.szFileName);
}

BOOL DoFileSaveAs()
{
    Document& doc = g_Document;

    OPENFILENAME ofn = {0};
    TCHAR szFile[MAX_PATH] = {0};

    if (doc.szFileName[0])
        StringCchCopy(szFile, MAX_PATH, doc.szFileName);

    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hMain;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = _T("All Files (*.*)\0*.*\0Text Files (*.txt)\0*.txt\0");
    ofn.nFilterIndex = 1;
    ofn.lpstrDefExt = _T("txt");
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;

    if (GetSaveFileName(&ofn))
    {
        // Check if plugin handles save
        if (g_PluginManager.NotifySaveFile(ofn.lpstrFile, doc.hEditor))
        {
            StringCchCopy(doc.szFileName, MAX_PATH, ofn.lpstrFile);
            doc.bIsDirty = FALSE;
            UpdateTitle();
            g_PluginManager.NotifyFileEvent(doc.szFileName, doc.hEditor, L"Saved");
            return TRUE;
        }

        HANDLE hFile = CreateFile(ofn.lpstrFile, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE)
        {
            // Write BOM
            DWORD written;
            if (doc.currentEncoding == ENC_UTF8)
            {
                BYTE bom[] = {0xEF, 0xBB, 0xBF};
                WriteFile(hFile, bom, 3, &written, NULL);
            }
            else if (doc.currentEncoding == ENC_UTF16LE)
            {
                BYTE bom[] = {0xFF, 0xFE};
                WriteFile(hFile, bom, 2, &written, NULL);
            }
            else if (doc.currentEncoding == ENC_UTF16BE)
            {
                BYTE bom[] = {0xFE, 0xFF};
                WriteFile(hFile, bom, 2, &written, NULL);
            }

            StreamContext ctx = {0};
            ctx.hFile = hFile;
            
            EDITSTREAM es = {0};
            es.dwCookie = (DWORD_PTR)&ctx;
            es.pfnCallback = StreamOutCallback;
            
            DWORD dwFormat = SF_TEXT;
            const EncodingInfo* info = GetEncodingInfo(doc.currentEncoding);
            
            if (doc.currentEncoding == ENC_UTF16LE || doc.currentEncoding == ENC_UTF16BE)
            {
                dwFormat = SF_TEXT | SF_UNICODE;
            }
            else
            {
                dwFormat = SF_TEXT | (info->codePage << 16) | SF_USECODEPAGE;
            }

            SendMessage(doc.hEditor, EM_STREAMOUT, dwFormat, (LPARAM)&es);
            CloseHandle(hFile);

            StringCchCopy(doc.szFileName, MAX_PATH, szFile);
            GetFileLastWriteTime(doc.szFileName, &doc.ftLastWriteTime);
            doc.bIsDirty = FALSE;
            UpdateTitle();
            AddRecentFile(doc.szFileName);
            
            // Notify plugins
            g_PluginManager.NotifyFileEvent(doc.szFileName, doc.hEditor, L"Saved");
            
            return TRUE;
        }
    }
    return FALSE;
}

BOOL DoFileSave()
{
    Document& doc = g_Document;

    if (doc.szFileName[0])
    {
        // Check if plugin handles save
        if (g_PluginManager.NotifySaveFile(doc.szFileName, doc.hEditor))
        {
            doc.bIsDirty = FALSE;
            UpdateTitle();
            g_PluginManager.NotifyFileEvent(doc.szFileName, doc.hEditor, L"Saved");
            return TRUE;
        }

        if (doc.isBinary) {
            MessageBox(hMain, _T("Cannot save binary files in text mode."), _T("Just Notepad"), MB_ICONWARNING);
            return FALSE;
        }

        HANDLE hFile = CreateFile(doc.szFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE)
        {
            // Write BOM
            DWORD written;
            if (doc.currentEncoding == ENC_UTF8)
            {
                BYTE bom[] = {0xEF, 0xBB, 0xBF};
                WriteFile(hFile, bom, 3, &written, NULL);
            }
            else if (doc.currentEncoding == ENC_UTF16LE)
            {
                BYTE bom[] = {0xFF, 0xFE};
                WriteFile(hFile, bom, 2, &written, NULL);
            }
            else if (doc.currentEncoding == ENC_UTF16BE)
            {
                BYTE bom[] = {0xFE, 0xFF};
                WriteFile(hFile, bom, 2, &written, NULL);
            }

            StreamContext ctx = {0};
            ctx.hFile = hFile;
            
            EDITSTREAM es = {0};
            es.dwCookie = (DWORD_PTR)&ctx;
            es.pfnCallback = StreamOutCallback;
            
            DWORD dwFormat = SF_TEXT;
            const EncodingInfo* info = GetEncodingInfo(doc.currentEncoding);
            
            if (doc.currentEncoding == ENC_UTF16LE || doc.currentEncoding == ENC_UTF16BE)
            {
                dwFormat = SF_TEXT | SF_UNICODE;
            }
            else
            {
                dwFormat = SF_TEXT | (info->codePage << 16) | SF_USECODEPAGE;
            }

            SendMessage(doc.hEditor, EM_STREAMOUT, dwFormat, (LPARAM)&es);
            CloseHandle(hFile);
            
            GetFileLastWriteTime(doc.szFileName, &doc.ftLastWriteTime);
            doc.bIsDirty = FALSE;
            UpdateTitle();
            
            // Notify plugins
            g_PluginManager.NotifyFileEvent(doc.szFileName, doc.hEditor, L"Saved");
            
            return TRUE;
        }
        return FALSE;
    }
    else
    {
        return DoFileSaveAs();
    }
}

void DoPageSetup()
{
    PAGESETUPDLG psd = {0};
    psd.lStructSize = sizeof(psd);
    psd.hwndOwner = hMain;
    PageSetupDlg(&psd);
}

void DoPrint()
{
    PRINTDLG pd = {0};
    pd.lStructSize = sizeof(pd);
    pd.hwndOwner = hMain;
    pd.Flags = PD_RETURNDC | PD_NOPAGENUMS | PD_NOSELECTION;

    if (PrintDlg(&pd))
    {
        DOCINFO di = {0};
        di.cbSize = sizeof(di);
        di.lpszDocName = _T("Just Notepad Document");

        if (StartDoc(pd.hDC, &di) > 0)
        {
            StartPage(pd.hDC);

            FORMATRANGE fr = {0};
            fr.hdc = pd.hDC;
            fr.hdcTarget = pd.hDC;
            fr.rc.top = 0;
            fr.rc.left = 0;
            fr.rc.right = GetDeviceCaps(pd.hDC, HORZRES);
            fr.rc.bottom = GetDeviceCaps(pd.hDC, VERTRES);
            fr.rcPage = fr.rc;
            fr.chrg.cpMin = 0;
            fr.chrg.cpMax = -1;

            SendMessage(hEditor, EM_FORMATRANGE, TRUE, (LPARAM)&fr);
            SendMessage(hEditor, EM_DISPLAYBAND, 0, (LPARAM)&fr.rc);

            EndPage(pd.hDC);
            EndDoc(pd.hDC);
            
            SendMessage(hEditor, EM_FORMATRANGE, FALSE, 0); // Clear cache
        }
        DeleteDC(pd.hDC);
    }
}

void DoFind()
{
    if (hFindReplaceDlg)
    {
        SetFocus(hFindReplaceDlg);
        return;
    }

    ZeroMemory(&fr, sizeof(fr));
    fr.lStructSize = sizeof(fr);
    fr.hwndOwner = hMain;
    fr.lpstrFindWhat = szFindWhat;
    fr.wFindWhatLen = sizeof(szFindWhat);
    fr.Flags = FR_DOWN;

    hFindReplaceDlg = FindText(&fr);
}

// Search next occurrence based on flags
BOOL FindNextText(DWORD dwFlags)
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);

    FINDTEXTEX ft = {};
    ft.chrg.cpMin = (dwFlags & FR_DOWN) ? cr.cpMax : cr.cpMin;
    ft.chrg.cpMax = (dwFlags & FR_DOWN) ? -1 : 0;
    ft.lpstrText = szFindWhat;

    DWORD dwSearchFlags = 0;
    if (dwFlags & FR_DOWN) dwSearchFlags |= FR_DOWN;
    if (dwFlags & FR_MATCHCASE) dwSearchFlags |= FR_MATCHCASE;
    if (dwFlags & FR_WHOLEWORD) dwSearchFlags |= FR_WHOLEWORD;

    LRESULT lResult = SendMessage(hEditor, EM_FINDTEXTEX, dwSearchFlags, (LPARAM)&ft);
    if (lResult != -1)
    {
        SendMessage(hEditor, EM_SETSEL, ft.chrgText.cpMin, ft.chrgText.cpMin);
        SendMessage(hEditor, EM_SCROLLCARET, 0, 0);
        SendMessage(hEditor, EM_SETSEL, ft.chrgText.cpMin, ft.chrgText.cpMax);
        SendMessage(hEditor, EM_SCROLLCARET, 0, 0);
        return TRUE;
    }
    return FALSE;
}

void DoReplace()
{
    if (hFindReplaceDlg)
    {
        SetFocus(hFindReplaceDlg);
        return;
    }

    ZeroMemory(&fr, sizeof(fr));
    fr.lStructSize = sizeof(fr);
    fr.hwndOwner = hMain;
    fr.lpstrFindWhat = szFindWhat;
    fr.wFindWhatLen = sizeof(szFindWhat);
    fr.lpstrReplaceWith = szReplaceWith;
    fr.wReplaceWithLen = sizeof(szReplaceWith);
    fr.Flags = FR_DOWN | FR_REPLACE;

    hFindReplaceDlg = ReplaceText(&fr);
}

INT_PTR CALLBACK AboutDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;
    case WM_NOTIFY:
        switch (((LPNMHDR)lParam)->code)
        {
        case NM_CLICK:
        case NM_RETURN:
            {
                PNMLINK pNMLink = (PNMLINK)lParam;
                ShellExecute(NULL, _T("open"), pNMLink->item.szUrl, NULL, NULL, SW_SHOW);
                return (INT_PTR)TRUE;
            }
        }
        break;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// Map to track pending changes: filename -> enabled
static std::map<std::wstring, bool> g_PendingPluginStates;
static bool g_bUpdatingList = false;

void UpdateManagePluginsList(HWND hDlg) {
    g_bUpdatingList = true;
    HWND hList = GetDlgItem(hDlg, IDC_MANAGE_PLUGINS_LIST);
    ListView_DeleteAllItems(hList);
    
    TCHAR szSearch[256];
    GetDlgItemText(hDlg, IDC_MANAGE_PLUGINS_SEARCH, szSearch, 256);
    std::wstring search = szSearch;
    std::transform(search.begin(), search.end(), search.begin(), ::towlower);

    const auto& plugins = g_PluginManager.GetPlugins();
    int itemIndex = 0;
    for (size_t i = 0; i < plugins.size(); ++i) {
        const auto& plugin = plugins[i];
        
        std::wstring nameLower = plugin.name;
        std::transform(nameLower.begin(), nameLower.end(), nameLower.begin(), ::towlower);
        
        if (search.empty() || nameLower.find(search) != std::wstring::npos) {
            // Determine check state
            bool enabled = plugin.enabled;
            auto it = g_PendingPluginStates.find(plugin.filename);
            if (it != g_PendingPluginStates.end()) {
                enabled = it->second;
            } else {
                g_PendingPluginStates[plugin.filename] = enabled;
            }

            LVITEM lvi;
            lvi.mask = LVIF_TEXT | LVIF_PARAM;
            lvi.iItem = itemIndex;
            lvi.iSubItem = 0;
            lvi.pszText = (LPWSTR)plugin.name.c_str();
            lvi.lParam = (LPARAM)i; // Store original index in PluginManager
            
            ListView_InsertItem(hList, &lvi);
            ListView_SetCheckState(hList, itemIndex, enabled);
            itemIndex++;
        }
    }
    g_bUpdatingList = false;
}

INT_PTR CALLBACK ManagePluginsDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
        {
            g_PendingPluginStates.clear();
            
            HWND hList = GetDlgItem(hDlg, IDC_MANAGE_PLUGINS_LIST);
            ListView_SetExtendedListViewStyle(hList, LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT);
            
            LVCOLUMN lvc;
            lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
            lvc.iSubItem = 0;
            lvc.pszText = (LPWSTR)L"Plugin";
            lvc.cx = 200;
            lvc.fmt = LVCFMT_LEFT;
            ListView_InsertColumn(hList, 0, &lvc);

            UpdateManagePluginsList(hDlg);
        }
        return (INT_PTR)TRUE;

    case WM_NOTIFY:
        {
            LPNMHDR pnmh = (LPNMHDR)lParam;
            if (pnmh->idFrom == IDC_MANAGE_PLUGINS_LIST) {
                if (pnmh->code == LVN_ITEMCHANGED) {
                    LPNMLISTVIEW pnmv = (LPNMLISTVIEW)lParam;
                    if (pnmv->uChanged & LVIF_STATE) {
                        // Check if selection changed to update description
                        if ((pnmv->uNewState & LVIS_SELECTED) && !(pnmv->uOldState & LVIS_SELECTED)) {
                            int idx = pnmv->iItem;
                            int originalIdx = (int)pnmv->lParam;
                            const auto& plugins = g_PluginManager.GetPlugins();
                            if (originalIdx >= 0 && originalIdx < (int)plugins.size()) {
                                SetDlgItemText(hDlg, IDC_MANAGE_PLUGINS_DESC, plugins[originalIdx].description.c_str());
                            }
                        }
                        
                        // Check if check state changed
                        // LVIS_STATEIMAGEMASK covers the checkbox (indices 1 and 2)
                        if (!g_bUpdatingList && (pnmv->uChanged & LVIF_STATE) && 
                            ((pnmv->uNewState & LVIS_STATEIMAGEMASK) != (pnmv->uOldState & LVIS_STATEIMAGEMASK))) {
                             
                             int originalIdx = (int)pnmv->lParam;
                             const auto& plugins = g_PluginManager.GetPlugins();
                             if (originalIdx >= 0 && originalIdx < (int)plugins.size()) {
                                 // Get new state directly from uNewState
                                 int stateImageIndex = (pnmv->uNewState & LVIS_STATEIMAGEMASK) >> 12;
                                 if (stateImageIndex == 1 || stateImageIndex == 2) {
                                     BOOL bChecked = (stateImageIndex == 2);
                                     g_PendingPluginStates[plugins[originalIdx].filename] = bChecked;
                                 }
                             }
                        }
                    }
                }
            }
        }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_MANAGE_PLUGINS_SEARCH && HIWORD(wParam) == EN_CHANGE) {
            UpdateManagePluginsList(hDlg);
        }
        else if (LOWORD(wParam) == IDOK) {
            bool bChanged = false;
            const auto& plugins = g_PluginManager.GetPlugins();

            // Apply pending changes
            for (const auto& pair : g_PendingPluginStates) {
                for (const auto& plugin : plugins) {
                    if (plugin.filename == pair.first) {
                        if (plugin.enabled != pair.second) {
                            bChanged = true;
                        }
                        break;
                    }
                }
                g_PluginManager.SetPluginEnabled(pair.first, pair.second);
            }
            
            // Save settings
            TCHAR szPath[MAX_PATH];
            GetModuleFileName(NULL, szPath, MAX_PATH);
            PathRemoveFileSpec(szPath);
            PathAppend(szPath, _T("config.ini"));
            g_PluginManager.SaveSettings(szPath);
            
            EndDialog(hDlg, IDOK);

            if (bChanged) {
                DoReload();
            }

            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, IDCANCEL);
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void DoAbout()
{
    DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUT), hMain, AboutDlgProc);
}

void LoadConfig()
{
    // Get path to config file (same dir as exe)
    TCHAR szPath[MAX_PATH];
    GetModuleFileName(NULL, szPath, MAX_PATH);
    PathRemoveFileSpec(szPath);
    PathCombine(szConfigPath, szPath, _T("config.ini"));

    // Settings
    bWordWrap = GetPrivateProfileInt(_T("Settings"), _T("WordWrap"), 0, szConfigPath);
    int nZoom = GetPrivateProfileInt(_T("Settings"), _T("Zoom"), 100, szConfigPath);
    BOOL bStatusBar = GetPrivateProfileInt(_T("Settings"), _T("StatusBar"), 1, szConfigPath);

    // Apply Settings
    if (bWordWrap) SendMessage(hEditor, EM_SETTARGETDEVICE, 0, 0);
    else SendMessage(hEditor, EM_SETTARGETDEVICE, 0, 1);

    SendMessage(hEditor, EM_SETZOOM, nZoom, 100);
    
    if (!bStatusBar) ShowWindow(hStatus, SW_HIDE);

    // Update Menu
    HMENU hMenu = GetMenu(hMain);
    CheckMenuItem(hMenu, ID_VIEW_WORDWRAP, bWordWrap ? MF_CHECKED : MF_UNCHECKED);
    CheckMenuItem(hMenu, ID_VIEW_STATUSBAR, bStatusBar ? MF_CHECKED : MF_UNCHECKED);

    // Recent Files
    recentFiles.clear();
    for (int i = 0; i < 10; i++)
    {
        TCHAR key[16];
        StringCchPrintf(key, 16, _T("File%d"), i + 1);
        TCHAR val[MAX_PATH];
        GetPrivateProfileString(_T("RecentFiles"), key, _T(""), val, MAX_PATH, szConfigPath);
        if (val[0])
        {
            recentFiles.push_back(val);
        }
    }
    UpdateRecentFilesMenu();
}

void SaveConfig()
{
    // Settings
    WritePrivateProfileString(_T("Settings"), _T("WordWrap"), bWordWrap ? _T("1") : _T("0"), szConfigPath);
    
    int nNum, nDen;
    SendMessage(hEditor, EM_GETZOOM, (WPARAM)&nNum, (LPARAM)&nDen);
    int nZoom = (nDen == 0) ? 100 : (nNum * 100 / nDen);
    TCHAR szZoom[16];
    StringCchPrintf(szZoom, 16, _T("%d"), nZoom);
    WritePrivateProfileString(_T("Settings"), _T("Zoom"), szZoom, szConfigPath);

    BOOL bStatusBar = IsWindowVisible(hStatus);
    WritePrivateProfileString(_T("Settings"), _T("StatusBar"), bStatusBar ? _T("1") : _T("0"), szConfigPath);

    // Recent Files
    for (int i = 0; i < 10; i++)
    {
        TCHAR key[16];
        StringCchPrintf(key, 16, _T("File%d"), i + 1);
        if (i < recentFiles.size())
        {
            WritePrivateProfileString(_T("RecentFiles"), key, recentFiles[i].c_str(), szConfigPath);
        }
        else
        {
            WritePrivateProfileString(_T("RecentFiles"), key, NULL, szConfigPath);
        }
    }
}

void AddRecentFile(LPCTSTR pszPath)
{
    std::wstring path = pszPath;
    // Remove if exists
    auto it = std::remove(recentFiles.begin(), recentFiles.end(), path);
    recentFiles.erase(it, recentFiles.end());
    
    // Push front
    recentFiles.push_front(path);
    
    // Limit to 10
    if (recentFiles.size() > 10)
    {
        recentFiles.resize(10);
    }
    
    UpdateRecentFilesMenu();
}

void UpdateRecentFilesMenu()
{
    HMENU hMenu = GetMenu(hMain);
    HMENU hFileMenu = GetSubMenu(hMenu, 0);
    // Find "Open Recent" submenu. It's at index 2 (New, Open, Recent...)
    // Wait, let's check resource.rc again.
    // MENUITEM "&New\tCtrl+N", ID_FILE_NEW (0)
    // MENUITEM "&Open...\tCtrl+O", ID_FILE_OPEN (1)
    // POPUP "Open &Recent" (2)
    
    HMENU hRecentMenu = GetSubMenu(hFileMenu, 2);
    
    // Clear existing
    int nCount = GetMenuItemCount(hRecentMenu);
    for (int i = 0; i < nCount; i++)
    {
        DeleteMenu(hRecentMenu, 0, MF_BYPOSITION);
    }
    
    if (recentFiles.empty())
    {
        AppendMenu(hRecentMenu, MF_STRING | MF_GRAYED, 0, _T("Empty"));
    }
    else
    {
        int id = ID_FILE_RECENT_FIRST;
        for (const auto& file : recentFiles)
        {
            // Truncate if too long for display?
            AppendMenu(hRecentMenu, MF_STRING, id++, file.c_str());
        }
        
        AppendMenu(hRecentMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hRecentMenu, MF_STRING, ID_FILE_RECENT_LAST, _T("Clear Recent List"));
    }
}

INT_PTR CALLBACK GoToDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK)
        {
            BOOL bTranslated;
            int nLine = GetDlgItemInt(hDlg, IDC_GOTO_LINE, &bTranslated, FALSE);
            if (bTranslated && nLine > 0)
            {
                int nIndex = SendMessage(hEditor, EM_LINEINDEX, nLine - 1, 0);
                if (nIndex != -1)
                {
                    SendMessage(hEditor, EM_SETSEL, nIndex, nIndex);
                    SendMessage(hEditor, EM_SCROLLCARET, 0, 0);
                }
            }
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void DoGoTo()
{
    DialogBox(hInst, MAKEINTRESOURCE(IDD_GOTO), hMain, GoToDlgProc);
}

void ToggleWordWrap()
{
    bWordWrap = !bWordWrap;
    if (bWordWrap)
    {
        SendMessage(hEditor, EM_SETTARGETDEVICE, 0, 0);
    }
    else
    {
        SendMessage(hEditor, EM_SETTARGETDEVICE, 0, 1); // 1 means huge width (no wrap)
    }
    
    // Update menu check
    HMENU hMenu = GetMenu(hMain);
    CheckMenuItem(hMenu, ID_VIEW_WORDWRAP, bWordWrap ? MF_CHECKED : MF_UNCHECKED);
}

WNDPROC wpOrigEditProc;

LRESULT CALLBACK EditorWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_DROPFILES)
    {
        HDROP hDrop = (HDROP)wParam;
        TCHAR szFile[MAX_PATH];
        UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
        
        if (nFiles > 0)
        {
            if (DragQueryFile(hDrop, 0, szFile, MAX_PATH))
            {
                OpenFile(szFile);
            }
        }
        DragFinish(hDrop);
        return 0;
    }
    else if (msg == WM_KEYDOWN)
    {
        if (wParam == VK_TAB)
        {
            if (GetKeyState(VK_SHIFT) & 0x8000)
            {
                DoUnindent();
                return 0;
            }
            else
            {
                CHARRANGE cr;
                SendMessage(hwnd, EM_EXGETSEL, 0, (LPARAM)&cr);
                long lineStart = SendMessage(hwnd, EM_EXLINEFROMCHAR, 0, cr.cpMin);
                long lineEnd = SendMessage(hwnd, EM_EXLINEFROMCHAR, 0, cr.cpMax);
                
                if (lineStart != lineEnd || cr.cpMin != cr.cpMax)
                {
                    DoIndent();
                    return 0;
                }
            }
        }
    }
    else if (msg == WM_CHAR)
    {
        if (wParam == VK_TAB)
        {
            if (GetKeyState(VK_SHIFT) & 0x8000)
            {
                return 0;
            }
            
            CHARRANGE cr;
            SendMessage(hwnd, EM_EXGETSEL, 0, (LPARAM)&cr);
            if (cr.cpMin != cr.cpMax)
            {
                return 0;
            }
        }
    }
    else if (msg == WM_MOUSEWHEEL)
    {
        // Force snappy scrolling (3 lines per notch, no smooth animation)
        short zDelta = GET_WHEEL_DELTA_WPARAM(wParam);
        int lines = 3;
        SystemParametersInfo(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
        if (lines == WHEEL_PAGESCROLL) lines = 3; // Fallback
        
        int scrollLines = (zDelta / WHEEL_DELTA) * lines;
        
        // Use EM_LINESCROLL for snappy scrolling
        SendMessage(hwnd, EM_LINESCROLL, 0, -scrollLines);
        return 0; // Handled
    }
    return CallWindowProc(wpOrigEditProc, hwnd, msg, wParam, lParam);
}

void CreateNewDocument(LPCTSTR pszFile)
{
    if (g_Document.hEditor == NULL)
    {
        // Create Editor
        RECT rcClient;
        GetClientRect(hMain, &rcClient);
        RECT rcStatus;
        GetWindowRect(hStatus, &rcStatus);
        int nStatusHeight = rcStatus.bottom - rcStatus.top;
        
        g_Document.hEditor = CreateWindowEx(0, MSFTEDIT_CLASS, _T(""),
                                     WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_NOHIDESEL | ES_DISABLENOSCROLL,
                                     0, 0, rcClient.right, rcClient.bottom - nStatusHeight,
                                     hMain, (HMENU)IDC_EDITOR, hInst, NULL);
        hEditor = g_Document.hEditor;

        // Disable smooth scrolling
        SendMessage(g_Document.hEditor, EM_SETOPTIONS, ECOOP_OR, ECO_AUTOVSCROLL);
        // Actually, smooth scrolling is often a system setting or mouse driver feature.
        // But for RichEdit, we can try to ensure it scrolls by lines.
        // There isn't a direct "Disable Smooth Scroll" message for RichEdit 4.1 (MSFTEDIT_CLASS).
        // However, removing WS_EX_COMPOSITED might help if it was there (it's not).
        // Some sources suggest handling WM_MOUSEWHEEL manually, but that's complex.
        // Let's try setting the scroll speed or similar if possible.
        // Wait, the user said "scrolling should be very snappy".
        // Maybe they mean the "smooth scroll" animation that some modern apps do?
        // RichEdit usually doesn't do smooth animation unless configured.
        // But let's check if we can force it.
        
        // One trick is to handle WM_MOUSEWHEEL and do line scrolling manually.
        // But let's first check if there is a simpler way.
        // SystemParametersInfo(SPI_GETWHEELSCROLLLINES, ...)
        
        // If the user means the "smooth scrolling" where the content glides, that's usually not default in standard Win32 RichEdit.
        // Unless they are on a trackpad.
        
        // However, if they mean "ES_AUTOVSCROLL" causes it to scroll when typing? No.
        
        // Let's try to handle WM_MOUSEWHEEL in EditorWndProc to force line jumps.
        
        SendMessage(g_Document.hEditor, EM_EXLIMITTEXT, 0, -1);
        SendMessage(g_Document.hEditor, EM_SETEVENTMASK, 0, ENM_SELCHANGE | ENM_CHANGE);
        
        HFONT hFont = CreateFont(18, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, 
                                     OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                                     DEFAULT_PITCH | FF_SWISS, _T("Consolas"));
        SendMessage(g_Document.hEditor, WM_SETFONT, (WPARAM)hFont, TRUE);

        // Subclass Editor
        wpOrigEditProc = (WNDPROC)SetWindowLongPtr(g_Document.hEditor, GWLP_WNDPROC, (LONG_PTR)EditorWndProc);
        
        // Disable OLE Drag & Drop to allow WM_DROPFILES
        RevokeDragDrop(g_Document.hEditor);

        // Enable Drag & Drop for Editor
        DragAcceptFiles(g_Document.hEditor, TRUE);
        
        ApplyTheme();
    }
    else
    {
        SetWindowText(g_Document.hEditor, _T(""));
    }

    g_Document.bIsDirty = FALSE;
    g_Document.bLoading = FALSE;
    g_Document.bPinned = FALSE;
    g_Document.currentEncoding = ENC_UTF8;
    g_Document.szFileName[0] = 0;
    g_Document.ftLastWriteTime = {0};

    if (pszFile)
    {
        StringCchCopy(g_Document.szFileName, MAX_PATH, pszFile);
        LoadFromFile(pszFile);
    }
    
    UpdateTitle();
    UpdateStatusBar();
}

void UpdateTitle()
{
    Document& doc = g_Document;
    
    // Set the full path as a window property for plugins to access
    // We cast the pointer to HANDLE. Since plugins run in the same process, they can read this memory.
    SetProp(hMain, _T("FullPath"), (HANDLE)doc.szFileName);
    
    TCHAR szTitle[MAX_PATH + 32];
    if (doc.szFileName[0])
        StringCchPrintf(szTitle, MAX_PATH + 32, _T("%s%s - Just Notepad"), PathFindFileName(doc.szFileName), doc.bIsDirty ? _T("*") : _T(""));
    else
        StringCchPrintf(szTitle, MAX_PATH + 32, _T("Untitled%s - Just Notepad"), doc.bIsDirty ? _T("*") : _T(""));
    SetWindowText(hMain, szTitle);
}

void UpdatePluginsMenu(HWND hwnd) {
    HMENU hMenu = GetMenu(hwnd);
    if (!hMenu) return;
    
    // Check if Plugins menu already exists and remove it to rebuild
    int count = GetMenuItemCount(hMenu);
    for (int i = 0; i < count; ++i) {
        TCHAR buffer[256];
        GetMenuString(hMenu, i, buffer, 256, MF_BYPOSITION);
        if (_tcscmp(buffer, _T("Plugins")) == 0) {
            RemoveMenu(hMenu, i, MF_BYPOSITION);
            break;
        }
    }

    HMENU hPluginsMenu = CreatePopupMenu();
    
    // Add Manage Plugins item first
    AppendMenu(hPluginsMenu, MF_STRING, ID_MANAGE_PLUGINS, _T("Manage Plugins..."));
    AppendMenu(hPluginsMenu, MF_STRING, ID_PLUGIN_SEARCH, _T("Search Commands...\tShift+Alt+S"));
    AppendMenu(hPluginsMenu, MF_SEPARATOR, 0, NULL);
    
    const auto& plugins = g_PluginManager.GetPlugins();
    if (plugins.empty()) {
        AppendMenu(hPluginsMenu, MF_STRING | MF_GRAYED, 0, _T("(No plugins loaded)"));
    } else {
        for (const auto& plugin : plugins) {
            if (!plugin.enabled) continue; // Skip disabled plugins
            if (plugin.items.empty()) continue;

            if (plugin.items.size() == 1) {
                // Single item: Add directly to the menu using the plugin name
                std::wstring text = plugin.name;
                if (!plugin.items[0].shortcut.empty()) {
                    text += L"\t" + plugin.items[0].shortcut;
                }
                AppendMenu(hPluginsMenu, MF_STRING, plugin.items[0].commandId, text.c_str());
            } else {
                // Multiple items: Create a submenu
                HMENU hSubMenu = CreatePopupMenu();
                for (const auto& item : plugin.items) {
                    std::wstring text = item.name;
                    if (!item.shortcut.empty()) {
                        text += L"\t" + item.shortcut;
                    }
                    AppendMenu(hSubMenu, MF_STRING, item.commandId, text.c_str());
                }
                
                // Add the plugin submenu to the main Plugins menu
                AppendMenu(hPluginsMenu, MF_POPUP | MF_STRING, (UINT_PTR)hSubMenu, plugin.name.c_str());
            }
        }
    }


    // Insert "Plugins" menu before "Help" (usually the last one)
    count = GetMenuItemCount(hMenu);
    int helpIndex = -1;
    for (int i = 0; i < count; ++i) {
        TCHAR buffer[256];
        GetMenuString(hMenu, i, buffer, 256, MF_BYPOSITION);
        // Check for "Help" or "&Help"
        if (_tcscmp(buffer, _T("Help")) == 0 || _tcscmp(buffer, _T("&Help")) == 0) {
            helpIndex = i;
            break;
        }
    }
    
    if (helpIndex != -1) {
        InsertMenu(hMenu, helpIndex, MF_BYPOSITION | MF_POPUP | MF_STRING, (UINT_PTR)hPluginsMenu, _T("Plugins"));
    } else {
        AppendMenu(hMenu, MF_POPUP | MF_STRING, (UINT_PTR)hPluginsMenu, _T("Plugins"));
    }
    DrawMenuBar(hwnd);
}



void UpdateStatusBar()
{
    if (!hStatus) return;
    
    Document& doc = g_Document;

    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    long nLine = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nCol = cr.cpMin - SendMessage(hEditor, EM_LINEINDEX, nLine, 0);
    
    // Get text length
    GETTEXTLENGTHEX gtl = { GTL_DEFAULT, 1200 }; // 1200 is CP_UNICODE
    long nLen = SendMessage(hEditor, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 0);

    // Get file size
    const EncodingInfo* info = GetEncodingInfo(doc.currentEncoding);
    long nBytes = 0;

    if (doc.bLoading)
    {
        nBytes = doc.fileSize;
    }
    else
    {
        GETTEXTLENGTHEX gtlSize = { GTL_NUMBYTES | GTL_PRECISE, info->codePage };
        nBytes = SendMessage(hEditor, EM_GETTEXTLENGTHEX, (WPARAM)&gtlSize, 0);
    }
    
    TCHAR szSize[64];
    StrFormatByteSize(nBytes, szSize, 64);

    // Get Zoom
    int nNum, nDen;
    SendMessage(hEditor, EM_GETZOOM, (WPARAM)&nNum, (LPARAM)&nDen);
    int nZoom = (nDen == 0) ? 100 : (nNum * 100 / nDen);

    TCHAR szStatus[256];
    StringCchPrintf(szStatus, 256, _T("Ln %d, Col %d"), nLine + 1, nCol + 1);
    g_StatusText[0] = szStatus;
    SendMessage(hStatus, SB_SETTEXT, 0 | SBT_OWNERDRAW, 0);

    StringCchPrintf(szStatus, 256, _T("Chars: %d"), nLen);
    g_StatusText[1] = szStatus;
    SendMessage(hStatus, SB_SETTEXT, 1 | SBT_OWNERDRAW, 0);

    StringCchPrintf(szStatus, 256, _T("Size: %s"), szSize);
    g_StatusText[2] = szStatus;
    SendMessage(hStatus, SB_SETTEXT, 2 | SBT_OWNERDRAW, 0);

    StringCchPrintf(szStatus, 256, _T("Zoom: %d%%"), nZoom);
    g_StatusText[3] = szStatus;
    SendMessage(hStatus, SB_SETTEXT, 3 | SBT_OWNERDRAW, 0);

    g_StatusText[4] = _T("Windows (CRLF)");
    SendMessage(hStatus, SB_SETTEXT, 4 | SBT_OWNERDRAW, 0);
    
    g_StatusText[5] = info->name;
    SendMessage(hStatus, SB_SETTEXT, 5 | SBT_OWNERDRAW, 0);

    // Plugin Status
    std::wstring pluginStatus;
    const auto& plugins = g_PluginManager.GetPlugins();
    for (const auto& p : plugins) {
        if (p.enabled && p.GetPluginStatus) {
            const wchar_t* status = p.GetPluginStatus(g_Document.szFileName);
            if (status && status[0]) {
                if (!pluginStatus.empty()) pluginStatus += L" | ";
                pluginStatus += status;
            }
        }
    }
    
    static std::wstring s_lastPluginStatus;
    s_lastPluginStatus = pluginStatus;
    g_StatusText[6] = s_lastPluginStatus.c_str();
    SendMessage(hStatus, SB_SETTEXT, 6 | SBT_OWNERDRAW, 0);
}



SelectionLevel GetNextSelectionLevel(SelectionLevel current)
{
    switch (current)
    {
    case SEL_NONE: return SEL_CHARACTER;
    case SEL_CHARACTER: return SEL_SUBWORD;
    case SEL_SUBWORD: return SEL_WORD;
    case SEL_WORD: return SEL_PHRASE;
    case SEL_PHRASE: return SEL_SENTENCE;
    case SEL_SENTENCE: return SEL_LINE;
    case SEL_LINE: return SEL_PARAGRAPH;
    case SEL_PARAGRAPH: return SEL_SECTION;
    case SEL_SECTION: return SEL_DOCUMENT;
    case SEL_DOCUMENT: return SEL_DOCUMENT;
    }
    return SEL_NONE;
}

std::basic_string<TCHAR> GetEditorText(HWND hEditor)
{
    long nLen = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    GETTEXTEX gt = {0};
    gt.cb = (nLen + 1) * sizeof(TCHAR);
    gt.flags = GT_DEFAULT; 
    gt.codepage = 1200; 

    std::vector<TCHAR> buffer(nLen + 1);
    SendMessage(hEditor, EM_GETTEXTEX, (WPARAM)&gt, (LPARAM)buffer.data());
    return std::basic_string<TCHAR>(buffer.data());
}

void DoExpandSelection()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);

    if (vecSelectionStack.empty() || 
        vecSelectionStack.back().cpMin != cr.cpMin || 
        vecSelectionStack.back().cpMax != cr.cpMax)
    {
        vecSelectionStack.clear();
        vecSelectionLevelStack.clear();
        vecSelectionStack.push_back(cr);
        vecSelectionLevelStack.push_back(SEL_NONE);
    }

    SelectionLevel currentLevel = vecSelectionLevelStack.back();
    SelectionLevel nextLevel = GetNextSelectionLevel(currentLevel);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, nextLevel);
    
    // If range didn't change, try next level
    while (r.start == cr.cpMin && r.end == cr.cpMax && nextLevel != SEL_DOCUMENT)
    {
        nextLevel = GetNextSelectionLevel(nextLevel);
        r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, nextLevel);
    }
    
    if (r.start != cr.cpMin || r.end != cr.cpMax)
    {
        CHARRANGE crNew = { r.start, r.end };
        vecSelectionStack.push_back(crNew);
        vecSelectionLevelStack.push_back(nextLevel);
        SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
    }
}

void DoShrinkSelection()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);

    if (vecSelectionStack.empty()) return;

    if (vecSelectionStack.back().cpMin == cr.cpMin && 
        vecSelectionStack.back().cpMax == cr.cpMax)
    {
        if (vecSelectionStack.size() > 1)
        {
            vecSelectionStack.pop_back();
            vecSelectionLevelStack.pop_back();
            CHARRANGE crPrev = vecSelectionStack.back();
            SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crPrev);
        }
        else
        {
            // Collapse to start if at bottom of stack
            CHARRANGE crNew = { cr.cpMin, cr.cpMin };
            SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
            vecSelectionStack.clear();
            vecSelectionLevelStack.clear();
        }
    }
    else
    {
        vecSelectionStack.clear();
        vecSelectionLevelStack.clear();
    }
}

void DoSelectLine()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
    
    SendMessage(hEditor, EM_SETSEL, r.start, r.end);
}

void DoCancelSelection()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    // If there's a selection, move cursor to the start of it
    if (cr.cpMin != cr.cpMax)
    {
        SendMessage(hEditor, EM_SETSEL, cr.cpMin, cr.cpMin);
    }
}



void DoCopyLine()
{
    CHARRANGE crOld;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&crOld);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), crOld.cpMin, crOld.cpMax, SEL_LINE);
    
    CHARRANGE crLine = {r.start, r.end};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crLine);
    SendMessage(hEditor, WM_COPY, 0, 0);
    
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crOld);
}

void DoDeleteLine()
{
    CHARRANGE crOld;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&crOld);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), crOld.cpMin, crOld.cpMax, SEL_LINE);
    
    CHARRANGE crLine = {r.start, r.end};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crLine);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)_T(""));
}

void DoCutLine()
{
    DoCopyLine();
    DoDeleteLine();
}

void DoDuplicateLine()
{
    CHARRANGE crOld;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&crOld);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), crOld.cpMin, crOld.cpMax, SEL_LINE);
    
    // Get text
    std::basic_string<TCHAR> buffer = text.substr(r.start, r.end - r.start);
    
    // Insert
    SendMessage(hEditor, EM_SETSEL, r.end, r.end);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)buffer.c_str());
    
    // Restore selection (maybe move it down?)
    long nDiff = r.end - r.start;
    CHARRANGE crNew = {crOld.cpMin + nDiff, crOld.cpMax + nDiff};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoDuplicateSelection()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    if (cr.cpMin == cr.cpMax)
    {
        DoDuplicateLine();
        return;
    }
    
    long nLen = cr.cpMax - cr.cpMin;
    std::vector<TCHAR> buffer(nLen + 1);
    TEXTRANGE tr;
    tr.chrg = cr;
    tr.lpstrText = buffer.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    SendMessage(hEditor, EM_SETSEL, cr.cpMax, cr.cpMax);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)buffer.data());
    
    // Select the new duplicate? Or keep original?
    // VS Code selects the original.
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&cr);
}

void DoMoveLineUp()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
    
    if (r.start == 0) return; // Cannot move up
    
    // Get previous line
    Range rPrev = GetSelectionRange(text.c_str(), (long)text.length(), r.start - 1, r.start - 1, SEL_LINE);
    
    std::basic_string<TCHAR> sBlock = text.substr(r.start, r.end - r.start);
    std::basic_string<TCHAR> sPrev = text.substr(rPrev.start, rPrev.end - rPrev.start);
    
    // Simply swap the two line blocks without modifying newlines
    std::basic_string<TCHAR> sNew = sBlock + sPrev;
    
    SendMessage(hEditor, EM_SETSEL, rPrev.start, r.end);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)sNew.c_str());
    
    long nOffsetStart = cr.cpMin - r.start;
    long nOffsetEnd = cr.cpMax - r.start;
    
    CHARRANGE crNew;
    crNew.cpMin = rPrev.start + nOffsetStart;
    crNew.cpMax = rPrev.start + nOffsetEnd;
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoMoveLineDown()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
    
    if (r.end >= (long)text.length()) return; // Cannot move down
    
    // Get next line
    Range rNext = GetSelectionRange(text.c_str(), (long)text.length(), r.end, r.end, SEL_LINE);
    
    std::basic_string<TCHAR> sBlock = text.substr(r.start, r.end - r.start);
    std::basic_string<TCHAR> sNext = text.substr(rNext.start, rNext.end - rNext.start);
    
    // Simply swap the two line blocks without modifying newlines
    std::basic_string<TCHAR> sNew = sNext + sBlock;
    
    SendMessage(hEditor, EM_SETSEL, r.start, rNext.end);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)sNew.c_str());
    
    long nNewBlockStart = r.start + sNext.length();
    
    long nOffsetStart = cr.cpMin - r.start;
    long nOffsetEnd = cr.cpMax - r.start;
    
    CHARRANGE crNew;
    crNew.cpMin = nNewBlockStart + nOffsetStart;
    crNew.cpMax = nNewBlockStart + nOffsetEnd;
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}



void DoUpperCase()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    if (cr.cpMin == cr.cpMax)
    {
        std::basic_string<TCHAR> text = GetEditorText(hEditor);
        Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
        
        if (r.end > r.start)
        {
            SendMessage(hEditor, EM_SETSEL, r.start, r.end);
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
        }
        else
        {
            return;
        }
    }
    
    long nLen = cr.cpMax - cr.cpMin;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg = cr;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::basic_string<TCHAR> out = ToUpperCase(buf.data());
    
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&cr);
}

void DoLowerCase()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    if (cr.cpMin == cr.cpMax)
    {
        std::basic_string<TCHAR> text = GetEditorText(hEditor);
        Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
        
        if (r.end > r.start)
        {
            SendMessage(hEditor, EM_SETSEL, r.start, r.end);
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
        }
        else return;
    }
    
    long nLen = cr.cpMax - cr.cpMin;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg = cr;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::basic_string<TCHAR> out = ToLowerCase(buf.data());
    
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&cr);
}

void DoCapitalize()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    if (cr.cpMin == cr.cpMax)
    {
        std::basic_string<TCHAR> text = GetEditorText(hEditor);
        Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
        
        if (r.end > r.start)
        {
            SendMessage(hEditor, EM_SETSEL, r.start, r.end);
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
        }
        else return;
    }
    
    long nLen = cr.cpMax - cr.cpMin;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg = cr;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::basic_string<TCHAR> out = ToCapitalize(buf.data());
    
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&cr);
}

void DoSentenceCase()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    if (cr.cpMin == cr.cpMax)
    {
        std::basic_string<TCHAR> text = GetEditorText(hEditor);
        Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
        
        if (r.end > r.start)
        {
            SendMessage(hEditor, EM_SETSEL, r.start, r.end);
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
        }
        else return;
    }
    
    long nLen = cr.cpMax - cr.cpMin;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg = cr;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::basic_string<TCHAR> out = ToSentenceCase(buf.data());
    
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&cr);
}

void DoSortLines()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
    
    std::basic_string<TCHAR> sub = text.substr(r.start, r.end - r.start);
    std::basic_string<TCHAR> out = SortLines(sub.c_str());
    
    SendMessage(hEditor, EM_SETSEL, r.start, r.end);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    
    // Restore original selection offsets
    long offsetStart = cr.cpMin - r.start;
    long offsetEnd = cr.cpMax - r.start;
    CHARRANGE crNew = {r.start + offsetStart, r.start + offsetEnd};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoJoinLines()
{
    // Join Lines feature removed
}

void DoNextSubword()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    long nLen = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    if (cr.cpMax >= nLen) return;
    
    // Scan forward
    long pos = cr.cpMax;
    // Read chunk
    TCHAR buf[256];
    TEXTRANGE tr;
    tr.chrg.cpMin = pos;
    tr.chrg.cpMax = min(pos + 255, nLen);
    tr.lpstrText = buf;
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    int len = tr.chrg.cpMax - tr.chrg.cpMin;
    int offset = CalculateNextSubword(buf, len);
    
    SendMessage(hEditor, EM_SETSEL, pos + offset, pos + offset);
}

void DoPrevSubword()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    if (cr.cpMin <= 0) return;
    
    long pos = cr.cpMin;
    long startPos = max(0, pos - 255);
    TCHAR buf[256];
    TEXTRANGE tr;
    tr.chrg.cpMin = startPos;
    tr.chrg.cpMax = pos;
    tr.lpstrText = buf;
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    int len = tr.chrg.cpMax - tr.chrg.cpMin;
    int offset = CalculatePrevSubword(buf, len);
    
    SendMessage(hEditor, EM_SETSEL, startPos + offset, startPos + offset);
}

void DoIndent()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
    
    std::basic_string<TCHAR> sub = text.substr(r.start, r.end - r.start);
    std::basic_string<TCHAR> out = IndentLines(sub.c_str());

    SendMessage(hEditor, EM_SETSEL, r.start, r.end);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    
    // Restore original selection offsets
    long offsetStart = cr.cpMin - r.start;
    long offsetEnd = cr.cpMax - r.start;
    
    std::basic_string<TCHAR> subStart = sub.substr(0, offsetStart);
    std::basic_string<TCHAR> subEnd = sub.substr(0, offsetEnd);
    
    long newOffsetStart = (long)IndentLines(subStart).length();
    long newOffsetEnd = (long)IndentLines(subEnd).length();
    
    CHARRANGE crNew = {r.start + newOffsetStart, r.start + newOffsetEnd};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoUnindent()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    std::basic_string<TCHAR> text = GetEditorText(hEditor);
    Range r = GetSelectionRange(text.c_str(), (long)text.length(), cr.cpMin, cr.cpMax, SEL_LINE);
    
    std::basic_string<TCHAR> sub = text.substr(r.start, r.end - r.start);
    std::basic_string<TCHAR> out = UnindentLines(sub.c_str());

    SendMessage(hEditor, EM_SETSEL, r.start, r.end);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    
    // Restore original selection offsets
    long offsetStart = cr.cpMin - r.start;
    long offsetEnd = cr.cpMax - r.start;
    
    std::basic_string<TCHAR> subStart = sub.substr(0, offsetStart);
    std::basic_string<TCHAR> subEnd = sub.substr(0, offsetEnd);
    
    long newOffsetStart = (long)UnindentLines(subStart).length();
    long newOffsetEnd = (long)UnindentLines(subEnd).length();
    
    CHARRANGE crNew = {r.start + newOffsetStart, r.start + newOffsetEnd};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoToggleReadOnly()
{
    Document& doc = g_Document;

    if (!doc.szFileName[0]) return;
    
    DWORD dwAttrs = GetFileAttributes(doc.szFileName);
    if (dwAttrs == INVALID_FILE_ATTRIBUTES) return;
    
    if (dwAttrs & FILE_ATTRIBUTE_READONLY)
    {
        SetFileAttributes(doc.szFileName, dwAttrs & ~FILE_ATTRIBUTE_READONLY);
        SendMessage(doc.hEditor, EM_SETREADONLY, FALSE, 0);
    }
    else
    {
        SetFileAttributes(doc.szFileName, dwAttrs | FILE_ATTRIBUTE_READONLY);
        SendMessage(doc.hEditor, EM_SETREADONLY, TRUE, 0);
    }
    
    // Update menu check
    HMENU hMenu = GetMenu(hMain);
    CheckMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, (dwAttrs & FILE_ATTRIBUTE_READONLY) ? MF_UNCHECKED : MF_CHECKED); // Inverted logic because we just toggled
}



struct PluginSearchState {
    std::vector<PluginCommand> allCommands;
    std::vector<PluginCommand> filteredCommands;
} g_PluginSearch;

void UpdatePluginSearchList(HWND hDlg) {
    HWND hList = GetDlgItem(hDlg, IDC_PLUGIN_SEARCH_LIST);
    SendMessage(hList, LB_RESETCONTENT, 0, 0);
    
    TCHAR szSearch[256];
    GetDlgItemText(hDlg, IDC_PLUGIN_SEARCH_TEXT, szSearch, 256);
    std::wstring search = szSearch;
    std::transform(search.begin(), search.end(), search.begin(), ::towlower);
    
    g_PluginSearch.filteredCommands.clear();
    
    for (const auto& cmd : g_PluginSearch.allCommands) {
        std::wstring name = cmd.commandName;
        std::wstring plugin = cmd.pluginName;
        std::wstring full = plugin + L": " + name;
        if (!cmd.shortcut.empty()) {
            full += L" (" + cmd.shortcut + L")";
        }
        std::wstring fullLower = full;
        std::transform(fullLower.begin(), fullLower.end(), fullLower.begin(), ::towlower);
        
        if (search.empty() || fullLower.find(search) != std::wstring::npos) {
            g_PluginSearch.filteredCommands.push_back(cmd);
            SendMessage(hList, LB_ADDSTRING, 0, (LPARAM)full.c_str());
        }
    }
}

INT_PTR CALLBACK PluginSearchDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        UpdatePluginSearchList(hDlg);
        SetFocus(GetDlgItem(hDlg, IDC_PLUGIN_SEARCH_TEXT));
        return FALSE; // We set focus
        
    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_PLUGIN_SEARCH_TEXT && HIWORD(wParam) == EN_CHANGE) {
            UpdatePluginSearchList(hDlg);
        }
        else if (LOWORD(wParam) == IDC_PLUGIN_SEARCH_LIST && HIWORD(wParam) == LBN_DBLCLK) {
            SendMessage(hDlg, WM_COMMAND, IDOK, 0);
        }
        else if (LOWORD(wParam) == IDOK) {
            HWND hList = GetDlgItem(hDlg, IDC_PLUGIN_SEARCH_LIST);
            int sel = (int)SendMessage(hList, LB_GETCURSEL, 0, 0);
            if (sel != LB_ERR && sel < (int)g_PluginSearch.filteredCommands.size()) {
                int cmdId = g_PluginSearch.filteredCommands[sel].commandId;
                EndDialog(hDlg, cmdId);
            } else {
                EndDialog(hDlg, 0);
            }
            return TRUE;
        }
        else if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, 0);
            return TRUE;
        }
        break;
    }
    return FALSE;
}

void DoPluginSearch() {
    g_PluginSearch.allCommands = g_PluginManager.GetAllCommands();
    INT_PTR result = DialogBox(hInst, MAKEINTRESOURCE(IDD_PLUGIN_SEARCH), hMain, PluginSearchDlgProc);
    if (result > 0) {
        g_PluginManager.ExecutePluginCommand((int)result, g_Document.hEditor);
    }
}

void DoManagePlugins() {
    if (DialogBox(hInst, MAKEINTRESOURCE(IDD_MANAGE_PLUGINS), hMain, ManagePluginsDlgProc) == IDOK) {
        UpdatePluginsMenu(hMain);
    }
}

// Theme support
int g_CurrentTheme = 0; // 0: Light, 1: Dark, 2: Solarized Dark

void ApplyTheme()
{
    COLORREF bgColor, textColor;
    
    switch (g_CurrentTheme)
    {
    case 1: // Dark
        bgColor = RGB(30, 30, 30);
        textColor = RGB(220, 220, 220);
        break;
    case 2: // Solarized Dark
        bgColor = RGB(0, 43, 54);
        textColor = RGB(131, 148, 150);
        break;
    default: // Light
        bgColor = RGB(255, 255, 255);
        textColor = RGB(0, 0, 0);
        break;
    }

    if (hEditor)
    {
        SendMessage(hEditor, EM_SETBKGNDCOLOR, 0, bgColor);
        
        CHARFORMAT2 cf = {0};
        cf.cbSize = sizeof(cf);
        cf.dwMask = CFM_COLOR;
        cf.crTextColor = textColor;
        SendMessage(hEditor, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cf);
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == uFindReplaceMsg)
    {
        LPFINDREPLACE lpfr = (LPFINDREPLACE)lParam;
        if (lpfr->Flags & FR_DIALOGTERM)
        {
            hFindReplaceDlg = NULL;
            return 0;
        }

        if (lpfr->Flags & FR_FINDNEXT)
        {
            if (!FindNextText(lpfr->Flags))
                MessageBox(hwnd, _T("Text not found."), _T("Find"), MB_ICONINFORMATION);
        }
        else if (lpfr->Flags & FR_REPLACE)
        {
            CHARRANGE cr;
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
            // Check if current selection matches
            // (Simplified: just replace current selection if it matches, else find next)
             SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)szReplaceWith);
             FindNextText(lpfr->Flags);
        }
        else if (lpfr->Flags & FR_REPLACEALL)
        {
            // Move to start
            SendMessage(hEditor, EM_SETSEL, 0, 0);
            int nCount = 0;
            
            int textLen = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
            
            SendMessage(hProgressBarStatus, PBM_SETRANGE32, 0, textLen);
            SendMessage(hProgressBarStatus, PBM_SETPOS, 0, 0);
            ShowWindow(hProgressBarStatus, SW_SHOW);
            
            while (FindNextText(lpfr->Flags & ~FR_DOWN)) // Force search down
            {
                SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)szReplaceWith);
                nCount++;
                
                CHARRANGE cr;
                SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
                SendMessage(hProgressBarStatus, PBM_SETPOS, cr.cpMax, 0);
            }
            ShowWindow(hProgressBarStatus, SW_HIDE);
            TCHAR szMsg[64];
            StringCchPrintf(szMsg, 64, _T("Replaced %d occurrences."), nCount);
            MessageBox(hwnd, szMsg, _T("Replace All"), MB_OK);
        }
        return 0;
    }

    switch (msg)
    {
    case WM_APP + 200: // WM_APP_SET_THEME
        g_CurrentTheme = (int)wParam;
        ApplyTheme();
        return 0;
    case WM_CONTEXTMENU:
    {
        if ((HWND)wParam == hEditor)
        {
            HMENU hMenu = LoadMenu(hInst, MAKEINTRESOURCE(IDR_CONTEXTMENU));
            HMENU hSubMenu = GetSubMenu(hMenu, 0);
            
            // Enable/Disable items based on selection
            CHARRANGE cr;
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
            BOOL bHasSelection = (cr.cpMax - cr.cpMin) != 0;
            BOOL bCanPaste = SendMessage(hEditor, EM_CANPASTE, 0, 0);
            
            // Always enable Cut/Copy/Delete (they fallback to line operations)
            EnableMenuItem(hSubMenu, ID_EDIT_CUT, MF_ENABLED);
            EnableMenuItem(hSubMenu, ID_EDIT_COPY, MF_ENABLED);
            EnableMenuItem(hSubMenu, ID_EDIT_PASTE, bCanPaste ? MF_ENABLED : MF_GRAYED);
            EnableMenuItem(hSubMenu, ID_EDIT_DELETE, MF_ENABLED);

            if (!bHasSelection)
            {
                ModifyMenu(hSubMenu, ID_EDIT_CUT, MF_BYCOMMAND | MF_STRING, ID_EDIT_CUT, _T("Cut Line\tCtrl+X"));
                ModifyMenu(hSubMenu, ID_EDIT_COPY, MF_BYCOMMAND | MF_STRING, ID_EDIT_COPY, _T("Copy Line\tCtrl+C"));
                ModifyMenu(hSubMenu, ID_EDIT_DELETE, MF_BYCOMMAND | MF_STRING, ID_EDIT_DELETE, _T("Delete Line\tDel"));
                
                ModifyMenu(hSubMenu, ID_EDIT_DUPLICATESELECTION, MF_BYCOMMAND | MF_STRING, ID_EDIT_DUPLICATESELECTION, _T("Duplicate Line\tCtrl+D"));
                
                ModifyMenu(hSubMenu, ID_EDIT_UPPERCASE, MF_BYCOMMAND | MF_STRING, ID_EDIT_UPPERCASE, _T("Uppercase Line\tCtrl+Shift+U"));
                ModifyMenu(hSubMenu, ID_EDIT_LOWERCASE, MF_BYCOMMAND | MF_STRING, ID_EDIT_LOWERCASE, _T("Lowercase Line\tCtrl+U"));
                ModifyMenu(hSubMenu, ID_EDIT_CAPITALIZE, MF_BYCOMMAND | MF_STRING, ID_EDIT_CAPITALIZE, _T("Capitalize Line"));
                ModifyMenu(hSubMenu, ID_EDIT_SENTENCECASE, MF_BYCOMMAND | MF_STRING, ID_EDIT_SENTENCECASE, _T("Sentence Case Line"));
                
                ModifyMenu(hSubMenu, ID_EDIT_INDENT, MF_BYCOMMAND | MF_STRING, ID_EDIT_INDENT, _T("Indent Line"));
                ModifyMenu(hSubMenu, ID_EDIT_UNINDENT, MF_BYCOMMAND | MF_STRING, ID_EDIT_UNINDENT, _T("Unindent Line"));
                
                // ModifyMenu(hSubMenu, ID_EDIT_JOINLINES, MF_BYCOMMAND | MF_STRING, ID_EDIT_JOINLINES, _T("Join Lines")); // removed
                ModifyMenu(hSubMenu, ID_EDIT_SORTLINES, MF_BYCOMMAND | MF_STRING, ID_EDIT_SORTLINES, _T("Sort Lines"));
            }
            else
            {
                ModifyMenu(hSubMenu, ID_EDIT_CUT, MF_BYCOMMAND | MF_STRING, ID_EDIT_CUT, _T("Cut\tCtrl+X"));
                ModifyMenu(hSubMenu, ID_EDIT_COPY, MF_BYCOMMAND | MF_STRING, ID_EDIT_COPY, _T("Copy\tCtrl+C"));
                ModifyMenu(hSubMenu, ID_EDIT_DELETE, MF_BYCOMMAND | MF_STRING, ID_EDIT_DELETE, _T("Delete\tDel"));

                ModifyMenu(hSubMenu, ID_EDIT_DUPLICATESELECTION, MF_BYCOMMAND | MF_STRING, ID_EDIT_DUPLICATESELECTION, _T("Duplicate Selection\tCtrl+D"));
                
                ModifyMenu(hSubMenu, ID_EDIT_UPPERCASE, MF_BYCOMMAND | MF_STRING, ID_EDIT_UPPERCASE, _T("Uppercase Selection\tCtrl+Shift+U"));
                ModifyMenu(hSubMenu, ID_EDIT_LOWERCASE, MF_BYCOMMAND | MF_STRING, ID_EDIT_LOWERCASE, _T("Lowercase Selection\tCtrl+U"));
                ModifyMenu(hSubMenu, ID_EDIT_CAPITALIZE, MF_BYCOMMAND | MF_STRING, ID_EDIT_CAPITALIZE, _T("Capitalize Selection"));
                ModifyMenu(hSubMenu, ID_EDIT_SENTENCECASE, MF_BYCOMMAND | MF_STRING, ID_EDIT_SENTENCECASE, _T("Sentence Case Selection"));
                
                ModifyMenu(hSubMenu, ID_EDIT_INDENT, MF_BYCOMMAND | MF_STRING, ID_EDIT_INDENT, _T("Indent Selection"));
                ModifyMenu(hSubMenu, ID_EDIT_UNINDENT, MF_BYCOMMAND | MF_STRING, ID_EDIT_UNINDENT, _T("Unindent Selection"));
                
                // ModifyMenu(hSubMenu, ID_EDIT_JOINLINES, MF_BYCOMMAND | MF_STRING, ID_EDIT_JOINLINES, _T("Join Selected Lines")); // removed
                ModifyMenu(hSubMenu, ID_EDIT_SORTLINES, MF_BYCOMMAND | MF_STRING, ID_EDIT_SORTLINES, _T("Sort Selected Lines"));
            }
            
            // Remove redundant items from the bottom if they exist
            DeleteMenu(hSubMenu, ID_EDIT_COPYLINE, MF_BYCOMMAND);
            DeleteMenu(hSubMenu, ID_EDIT_DELETELINE, MF_BYCOMMAND);
            DeleteMenu(hSubMenu, ID_EDIT_DUPLICATELINE, MF_BYCOMMAND);
            
            // Check Unindent availability
            long nLine = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
            long nIndex = SendMessage(hEditor, EM_LINEINDEX, nLine, 0);
            long nLen = SendMessage(hEditor, EM_LINELENGTH, nIndex, 0);
            BOOL bCanUnindent = FALSE;
            if (nLen > 0)
            {
                TCHAR buf[16]; 
                TEXTRANGE tr;
                tr.chrg.cpMin = nIndex;
                tr.chrg.cpMax = min(nIndex + 10, nIndex + nLen);
                tr.lpstrText = buf;
                SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                if (buf[0] == '\t' || buf[0] == ' ') bCanUnindent = TRUE;
            }
            EnableMenuItem(hSubMenu, ID_EDIT_UNINDENT, bCanUnindent ? MF_ENABLED : MF_GRAYED);
            
            // Disable Sort/Join for single line selection
            long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
            long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMax);
            if (nLineStart == nLineEnd)
            {
                EnableMenuItem(hSubMenu, ID_EDIT_SORTLINES, MF_GRAYED);
                // EnableMenuItem(hSubMenu, ID_EDIT_JOINLINES, MF_GRAYED); // removed
            }
            else
            {
                EnableMenuItem(hSubMenu, ID_EDIT_SORTLINES, MF_ENABLED);
                // EnableMenuItem(hSubMenu, ID_EDIT_JOINLINES, MF_ENABLED); // removed
            }

            POINT pt;
            GetCursorPos(&pt);
            TrackPopupMenu(hSubMenu, TPM_LEFTALIGN | TPM_TOPALIGN, pt.x, pt.y, 0, hwnd, NULL);
            DestroyMenu(hMenu);
            return 0;
        }
    }
    break;
    case WM_DROPFILES:
    {
        HDROP hDrop = (HDROP)wParam;
        TCHAR szFile[MAX_PATH];
        UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
        
        if (nFiles > 0)
        {
            if (DragQueryFile(hDrop, 0, szFile, MAX_PATH))
            {
                OpenFile(szFile);
            }
        }
        DragFinish(hDrop);
    }
    break;
    case WM_CREATE:
    {
        hMain = hwnd;
        DragAcceptFiles(hwnd, TRUE);
        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        
        // Create Status Bar
        hStatus = CreateWindowEx(0, STATUSCLASSNAME, NULL,
            WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0,
            hwnd, (HMENU)IDC_STATUSBAR, hInst, NULL);
            
        int statwidths[] = {150, 300, 450, 550, 700, 850, -1};
        SendMessage(hStatus, SB_SETPARTS, 7, (LPARAM)statwidths);

        // Create Progress Bar in Status Bar
        hProgressBarStatus = CreateWindowEx(0, PROGRESS_CLASS, NULL, WS_CHILD,
            0, 0, 0, 0, hStatus, NULL, hInst, NULL);

        // Create Initial Document
        CreateNewDocument(NULL);

        // Register Find/Replace Message
        uFindReplaceMsg = RegisterWindowMessage(FINDMSGSTRING);
        
        // Initial menu state
        HMENU hMenu = GetMenu(hwnd);
        CheckMenuItem(hMenu, ID_VIEW_STATUSBAR, MF_CHECKED);

        // Load Plugin Settings
        TCHAR szPath[MAX_PATH];
        GetModuleFileName(NULL, szPath, MAX_PATH);
        PathRemoveFileSpec(szPath);
        PathAppend(szPath, _T("config.ini"));
        g_PluginManager.LoadSettings(szPath);

        UpdateTitle();
        UpdateStatusBar();
        
        SetTimer(hwnd, 1, 2000, NULL);
    }
    break;
    case WM_LOAD_PROGRESS:
        SendMessage(hProgressBarStatus, PBM_SETPOS, wParam, 0);
        return 0;
    case WM_LOAD_COMPLETE:
    {
        LoadResult* res = (LoadResult*)lParam;
        if (res && res->bSuccess && res->sequence == g_Document.loadSequence)
        {
            struct MemStreamContext {
                LoadResult* res;
                size_t offset;
            } ctx;
            ctx.res = res;
            ctx.offset = 0;
            
            EDITSTREAM es = {0};
            es.dwCookie = (DWORD_PTR)&ctx;
            es.pfnCallback = [](DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb) -> DWORD {
                MemStreamContext* c = (MemStreamContext*)dwCookie;
                size_t remaining = c->res->buffer.size() * sizeof(wchar_t) - c->offset;
                if (remaining == 0) {
                    *pcb = 0;
                    return 0;
                }
                
                LONG toCopy = min(cb, (LONG)remaining);
                memcpy(pbBuff, (BYTE*)c->res->buffer.data() + c->offset, toCopy);
                c->offset += toCopy;
                *pcb = toCopy;
                return 0;
            };
            
            SendMessage(g_Document.hEditor, WM_SETREDRAW, FALSE, 0);
            
            // Replace all content
            SendMessage(g_Document.hEditor, EM_SETSEL, 0, -1);
            SendMessage(g_Document.hEditor, EM_STREAMIN, SF_TEXT | SF_UNICODE, (LPARAM)&es);
            
            SendMessage(g_Document.hEditor, WM_SETREDRAW, TRUE, 0);
            InvalidateRect(g_Document.hEditor, NULL, TRUE);
            
            DWORD dwAttrs = GetFileAttributes(g_Document.szFileName);
            BOOL bReadOnly = (dwAttrs != INVALID_FILE_ATTRIBUTES) && (dwAttrs & FILE_ATTRIBUTE_READONLY);
            SendMessage(g_Document.hEditor, EM_SETREADONLY, bReadOnly, 0);
            
            // Notify plugins
            g_PluginManager.NotifyFileEvent(g_Document.szFileName, g_Document.hEditor, L"Loaded");

            ShowWindow(hProgressBarStatus, SW_HIDE);
            g_Document.bLoading = FALSE;
            
            // Ensure event mask is set
            SendMessage(g_Document.hEditor, EM_SETEVENTMASK, 0, ENM_SELCHANGE | ENM_CHANGE);
            
            UpdateStatusBar();
        }
        
        if (res) delete res;
        return 0;
    }
    case WM_TIMER:
        if (wParam == 1)
        {
            Document& doc = g_Document;
            if (doc.szFileName[0])
            {
                FILETIME ftCurrent;
                if (GetFileLastWriteTime(doc.szFileName, &ftCurrent))
                {
                    if (CompareFileTime(&doc.ftLastWriteTime, &ftCurrent) != 0)
                    {
                        KillTimer(hwnd, 1);
                        
                        TCHAR szMsg[MAX_PATH + 100];
                        StringCchPrintf(szMsg, MAX_PATH + 100, _T("The file '%s' has been modified by another program.\nDo you want to reload it?"), doc.szFileName);
                        
                        if (MessageBox(hwnd, szMsg, _T("Just Notepad"), MB_YESNO | MB_ICONQUESTION) == IDYES)
                        {
                            DoReload();
                        }
                        else
                        {
                            doc.ftLastWriteTime = ftCurrent;
                        }
                        
                        SetTimer(hwnd, 1, 2000, NULL);
                    }
                }

                // Check for attribute changes (ReadOnly)
                DWORD dwAttrs = GetFileAttributes(doc.szFileName);
                if (dwAttrs != INVALID_FILE_ATTRIBUTES)
                {
                    BOOL bFileReadOnly = (dwAttrs & FILE_ATTRIBUTE_READONLY) != 0;
                    
                    // Check current editor state
                    // Use EM_GETREADONLY for RichEdit
                    // But we can also check the menu state or just force sync
                    // Let's check the style
                    DWORD dwStyle = GetWindowLong(g_Document.hEditor, GWL_STYLE);
                    BOOL bEditorReadOnly = (dwStyle & ES_READONLY) != 0;
                    
                    if (bFileReadOnly != bEditorReadOnly)
                    {
                        SendMessage(g_Document.hEditor, EM_SETREADONLY, bFileReadOnly, 0);
                        HMENU hMenu = GetMenu(hwnd);
                        CheckMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, bFileReadOnly ? MF_CHECKED : MF_UNCHECKED);
                    }
                }
            }
            
            // Update status bar periodically to reflect plugin status changes
            UpdateStatusBar();
        }
        break;
    case WM_UPDATE_PROGRESS:
    {
        int percent = (int)wParam;
        if (hProgressBarStatus) {
            if (percent < 0) {
                ShowWindow(hProgressBarStatus, SW_HIDE);
            } else {
                // Ensure it's visible and positioned correctly (in case it was hidden during resize)
                if (!IsWindowVisible(hProgressBarStatus)) {
                    ShowWindow(hProgressBarStatus, SW_SHOW);
                    // Trigger resize to ensure position
                    SendMessage(hwnd, WM_SIZE, 0, 0);
                }
                SendMessage(hProgressBarStatus, PBM_SETPOS, percent, 0);
            }
        }
        return 0;
    }
    case WM_SIZE:
    {
        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        
        SendMessage(hStatus, WM_SIZE, 0, 0);
        
        int nStatusHeight = 0;
        if (IsWindowVisible(hStatus))
        {
            RECT rcStatus;
            GetWindowRect(hStatus, &rcStatus);
            nStatusHeight = rcStatus.bottom - rcStatus.top;

            // Position Progress Bar in the last part of status bar
            if (hProgressBarStatus) {
                RECT rcPart;
                SendMessage(hStatus, SB_GETRECT, 6, (LPARAM)&rcPart);
                // Add some padding
                InflateRect(&rcPart, -2, -2);
                SetWindowPos(hProgressBarStatus, NULL, rcPart.left, rcPart.top, rcPart.right - rcPart.left, rcPart.bottom - rcPart.top, SWP_NOZORDER);
            }
        }

        if (g_Document.hEditor)
        {
            MoveWindow(g_Document.hEditor, 0, 0, rcClient.right, rcClient.bottom - nStatusHeight, TRUE);
        }

        // Resize Progress Bar
        if (hProgressBarStatus)
        {
            RECT rcPart;
            SendMessage(hStatus, SB_GETRECT, 1, (LPARAM)&rcPart); // Part 1
            MoveWindow(hProgressBarStatus, rcPart.left, rcPart.top, rcPart.right - rcPart.left, rcPart.bottom - rcPart.top, TRUE);
        }
    }
    break;
    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT lpDrawItem = (LPDRAWITEMSTRUCT)lParam;
        if (lpDrawItem->hwndItem == hStatus)
        {
            int nPart = lpDrawItem->itemID;
            HDC hdc = lpDrawItem->hDC;
            RECT rc = lpDrawItem->rcItem;

            // Background
            COLORREF crBk = GetSysColor(COLOR_BTNFACE);
            HBRUSH hBr = CreateSolidBrush(crBk);
            FillRect(hdc, &rc, hBr);
            DeleteObject(hBr);

            // Text
            COLORREF crText = GetSysColor(COLOR_BTNTEXT);
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, crText);
            
            // Draw Text
            if (nPart >= 0 && nPart < 7)
            {
                // Add some padding
                rc.left += 4;
                DrawText(hdc, g_StatusText[nPart].c_str(), -1, &rc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            }
            return TRUE;
        }
    }
    break;
    case WM_NOTIFY:
    {
        NMHDR *pnm = (NMHDR *)lParam;
        
        if (pnm->hwndFrom == g_Document.hEditor && pnm->code == EN_SELCHANGE)
        {
            UpdateStatusBar();
        }
        else if (pnm->hwndFrom == hStatus && pnm->code == NM_CLICK)
        {
            NMMOUSE *pMouse = (NMMOUSE *)lParam;
            if (pMouse->dwItemSpec == 3) // Zoom
            {
                int nNum = 0, nDen = 0;
                SendMessage(hEditor, EM_GETZOOM, (WPARAM)&nNum, (LPARAM)&nDen);
                int currentZoom = (nDen == 0) ? 100 : (nNum * 100 / nDen);

                HMENU hMenu = CreatePopupMenu();
                AppendMenu(hMenu, MF_STRING | (currentZoom == 50 ? MF_CHECKED : 0), 10001, _T("50%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 60 ? MF_CHECKED : 0), 10002, _T("60%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 70 ? MF_CHECKED : 0), 10003, _T("70%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 80 ? MF_CHECKED : 0), 10004, _T("80%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 90 ? MF_CHECKED : 0), 10005, _T("90%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 100 ? MF_CHECKED : 0), 10006, _T("100%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 110 ? MF_CHECKED : 0), 10007, _T("110%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 120 ? MF_CHECKED : 0), 10008, _T("120%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 130 ? MF_CHECKED : 0), 10009, _T("130%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 140 ? MF_CHECKED : 0), 10010, _T("140%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 150 ? MF_CHECKED : 0), 10011, _T("150%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 175 ? MF_CHECKED : 0), 10012, _T("175%"));
                AppendMenu(hMenu, MF_STRING | (currentZoom == 200 ? MF_CHECKED : 0), 10013, _T("200%"));
                
                POINT pt;
                GetCursorPos(&pt);
                int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_NONOTIFY, pt.x, pt.y, 0, hwnd, NULL);
                DestroyMenu(hMenu);
                
                nNum = 0; nDen = 100;
                switch(cmd)
                {
                    case 10001: nNum = 50; break;
                    case 10002: nNum = 60; break;
                    case 10003: nNum = 70; break;
                    case 10004: nNum = 80; break;
                    case 10005: nNum = 90; break;
                    case 10006: nNum = 100; break;
                    case 10007: nNum = 110; break;
                    case 10008: nNum = 120; break;
                    case 10009: nNum = 130; break;
                    case 10010: nNum = 140; break;
                    case 10011: nNum = 150; break;
                    case 10012: nNum = 175; break;
                    case 10013: nNum = 200; break;
                }
                
                if (nNum > 0)
                {
                    SendMessage(hEditor, EM_SETZOOM, nNum, nDen);
                    UpdateStatusBar();
                }
            }
            else if (pMouse->dwItemSpec == 5) // Encoding
            {
                HMENU hMenu = CreatePopupMenu();
                for (const auto& info : g_Encodings)
                {
                    UINT uFlags = MF_STRING;
                    if (g_Document.currentEncoding == info.id)
                    {
                        uFlags |= MF_CHECKED;
                    }
                    AppendMenu(hMenu, uFlags, info.cmdId, info.name);
                }
                
                POINT pt;
                GetCursorPos(&pt);
                TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_TOPALIGN, pt.x, pt.y, 0, hwnd, NULL);
                DestroyMenu(hMenu);
            }
        }
    }
    break;
    case WM_INITMENUPOPUP:
    {
        HMENU hMenu = (HMENU)wParam;
        
        if (GetMenuState(hMenu, ID_EDIT_UNDO, MF_BYCOMMAND) != -1)
        {
            BOOL bCanUndo = SendMessage(hEditor, EM_CANUNDO, 0, 0);
            EnableMenuItem(hMenu, ID_EDIT_UNDO, bCanUndo ? MF_ENABLED : MF_GRAYED);
            
            BOOL bCanRedo = SendMessage(hEditor, EM_CANREDO, 0, 0);
            EnableMenuItem(hMenu, ID_EDIT_REDO, bCanRedo ? MF_ENABLED : MF_GRAYED);
            
            CHARRANGE cr;
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
            
            BOOL bCanPaste = SendMessage(hEditor, EM_CANPASTE, 0, 0);
            EnableMenuItem(hMenu, ID_EDIT_PASTE, bCanPaste ? MF_ENABLED : MF_GRAYED);
            
            // Check Unindent
            long nLine = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
            long nIndex = SendMessage(hEditor, EM_LINEINDEX, nLine, 0);
            long nLen = SendMessage(hEditor, EM_LINELENGTH, nIndex, 0);
            BOOL bCanUnindent = FALSE;
            if (nLen > 0)
            {
                TCHAR buf[16]; 
                TEXTRANGE tr;
                tr.chrg.cpMin = nIndex;
                tr.chrg.cpMax = min(nIndex + 10, nIndex + nLen);
                tr.lpstrText = buf;
                SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                if (buf[0] == '\t' || buf[0] == ' ') bCanUnindent = TRUE;
            }
            EnableMenuItem(hMenu, ID_EDIT_UNINDENT, bCanUnindent ? MF_ENABLED : MF_GRAYED);
        }


        
        Document& doc = g_Document;
        for (const auto& info : g_Encodings)
        {
             CheckMenuItem(hMenu, info.cmdId, (doc.currentEncoding == info.id) ? MF_CHECKED : MF_UNCHECKED);
        }

        // Read Only
        long style = GetWindowLong(g_Document.hEditor, GWL_STYLE);
        CheckMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, (style & ES_READONLY) ? MF_CHECKED : MF_UNCHECKED);
        EnableMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, g_Document.bLoading ? MF_GRAYED : MF_ENABLED);
        
        // View Menu Items
        CheckMenuItem(hMenu, ID_VIEW_STATUSBAR, IsWindowVisible(hStatus) ? MF_CHECKED : MF_UNCHECKED);
    }
    break;
    case WM_COMMAND:
    {
        if (LOWORD(wParam) == IDC_EDITOR && HIWORD(wParam) == EN_CHANGE)
        {
            Document& doc = g_Document;
            if (!doc.bLoading)
            {
                if (!doc.bIsDirty)
                {
                    doc.bIsDirty = TRUE;
                    UpdateTitle();
                    g_PluginManager.NotifyFileEvent(doc.szFileName, doc.hEditor, L"Modified");
                }
                // Notify plugins about text change (always, not just on first dirty)
                g_PluginManager.NotifyTextModified(doc.hEditor);
            }
        }

        switch (LOWORD(wParam))
        {
        case ID_FILE_NEW: DoFileNew(); break;
        case ID_FILE_OPEN: DoFileOpen(); break;
        case ID_FILE_SAVE: DoFileSave(); break;
        case ID_FILE_SAVEAS: DoFileSaveAs(); break;
        case ID_FILE_RELOAD: DoReload(); break;
        case ID_FILE_PAGESETUP: DoPageSetup(); break;
        case ID_FILE_PRINT: DoPrint(); break;
        case ID_PLUGIN_SEARCH: DoPluginSearch(); break;
        case ID_MANAGE_PLUGINS: DoManagePlugins(); break;
        case ID_FILE_EXIT: 
            if (CheckAllSaveChanges()) 
            {
                SaveConfig();
                PostQuitMessage(0); 
            }
            break;

        case ID_ENCODING_ANSI: 
        case ID_ENCODING_UTF8: 
        case ID_ENCODING_UTF16LE: 
        case ID_ENCODING_UTF16BE: 
        case ID_ENCODING_ASCII:
        case ID_ENCODING_ISO_8859_1:
        case ID_ENCODING_WINDOWS_1252:
        case ID_ENCODING_UTF32:
        case ID_ENCODING_ISO_8859_15:
        case ID_ENCODING_SHIFT_JIS:
        case ID_ENCODING_EUC_JP:
        case ID_ENCODING_BIG5:
        case ID_ENCODING_GB18030:
        case ID_ENCODING_GBK:
        case ID_ENCODING_GB2312:
        case ID_ENCODING_EUC_KR:
        case ID_ENCODING_ISO_8859_2:
        case ID_ENCODING_WINDOWS_1250:
        case ID_ENCODING_KOI8_R:
        case ID_ENCODING_ISO_8859_5:
        case ID_ENCODING_WINDOWS_1251:
        case ID_ENCODING_ISO_8859_6:
        case ID_ENCODING_ISO_8859_7:
        case ID_ENCODING_ISO_8859_8:
        case ID_ENCODING_ISO_8859_9:
            {
                const EncodingInfo* info = GetEncodingInfoByCmd(LOWORD(wParam));
                if (info)
                {
                    g_Document.currentEncoding = info->id; 
                    g_Document.bIsDirty = TRUE;
                    UpdateStatusBar(); 
                    UpdateTitle(); 
                }
            }
            break;


        
        case ID_HELP_ABOUT:
            DoAbout();
            break;

        case ID_EDIT_UNDO: SendMessage(hEditor, EM_UNDO, 0, 0); break;
        case ID_EDIT_REDO: SendMessage(hEditor, EM_REDO, 0, 0); break;
        case ID_EDIT_CUT: 
        {
            CHARRANGE cr;
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
            if (cr.cpMin == cr.cpMax) DoCutLine();
            else SendMessage(hEditor, WM_CUT, 0, 0);
        }
        break;
        case ID_EDIT_COPY: 
        {
            CHARRANGE cr;
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
            if (cr.cpMin == cr.cpMax) DoCopyLine();
            else SendMessage(hEditor, WM_COPY, 0, 0);
        }
        break;
        case ID_EDIT_PASTE: SendMessage(hEditor, WM_PASTE, 0, 0); break;
        case ID_EDIT_DELETE: 
        {
            CHARRANGE cr;
            SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
            if (cr.cpMin == cr.cpMax) DoDeleteLine();
            else SendMessage(hEditor, WM_CLEAR, 0, 0);
        }
        break;
        case ID_EDIT_SELECTALL: SendMessage(hEditor, EM_SETSEL, 0, -1); break;
        
        case ID_EDIT_FIND: DoFind(); break;

        case ID_EDIT_REPLACE: DoReplace(); break;
        case ID_EDIT_FINDNEXT: 
            if (szFindWhat[0]) FindNextText(FR_DOWN); 
            else DoFind();
            break;
        case ID_EDIT_FINDPREV:
            if (szFindWhat[0]) FindNextText(0); // 0 means up
            else DoFind();
            break;
        case ID_EDIT_GOTO: DoGoTo(); break;
        case ID_EDIT_EXPAND_SELECTION: DoExpandSelection(); break;
        case ID_EDIT_SHRINK_SELECTION: DoShrinkSelection(); break;
        case ID_EDIT_COPYLINE: DoCopyLine(); break;
        case ID_EDIT_DELETELINE: DoDeleteLine(); break;
        case ID_EDIT_DUPLICATELINE: DoDuplicateLine(); break;
        case ID_EDIT_DUPLICATESELECTION: DoDuplicateSelection(); break;
        case ID_EDIT_MOVELINEUP: DoMoveLineUp(); break;
        case ID_EDIT_MOVELINEDOWN: DoMoveLineDown(); break;
        case ID_FILE_TOGGLEREADONLY: DoToggleReadOnly(); break;
        case ID_EDIT_SELECTLINE: DoSelectLine(); break;
        case ID_EDIT_CANCEL_SELECTION: DoCancelSelection(); break;
        case ID_EDIT_UPPERCASE: DoUpperCase(); break;
        case ID_EDIT_LOWERCASE: DoLowerCase(); break;
        case ID_EDIT_CAPITALIZE: DoCapitalize(); break;
        case ID_EDIT_SENTENCECASE: DoSentenceCase(); break;
        case ID_EDIT_SORTLINES: DoSortLines(); break;
        // case ID_EDIT_JOINLINES: DoJoinLines(); break; // removed
        case ID_EDIT_NEXTSUBWORD: DoNextSubword(); break;
        case ID_EDIT_PREVSUBWORD: DoPrevSubword(); break;
        case ID_EDIT_INDENT: DoIndent(); break;
        case ID_EDIT_UNINDENT: DoUnindent(); break;

        case ID_VIEW_ZOOMIN: 
        {
            int nNum, nDen;
            SendMessage(hEditor, EM_GETZOOM, (WPARAM)&nNum, (LPARAM)&nDen);
            if (nDen == 0) { nNum = 100; nDen = 100; }
            nNum += 10;
            SendMessage(hEditor, EM_SETZOOM, nNum, nDen);
            UpdateStatusBar();
        }
        break;
        case ID_VIEW_ZOOMOUT:
        {
            int nNum, nDen;
            SendMessage(hEditor, EM_GETZOOM, (WPARAM)&nNum, (LPARAM)&nDen);
            if (nDen == 0) { nNum = 100; nDen = 100; }
            nNum -= 10;
            if (nNum < 10) nNum = 10;
            SendMessage(hEditor, EM_SETZOOM, nNum, nDen);
            UpdateStatusBar();
        }
        break;
        case ID_VIEW_WORDWRAP: ToggleWordWrap(); break;
        case ID_VIEW_STATUSBAR:
        {
            BOOL bVisible = IsWindowVisible(hStatus);
            ShowWindow(hStatus, bVisible ? SW_HIDE : SW_SHOW);
            
            HMENU hMenu = GetMenu(hMain);
            CheckMenuItem(hMenu, ID_VIEW_STATUSBAR, bVisible ? MF_UNCHECKED : MF_CHECKED);
            
            SendMessage(hwnd, WM_SIZE, 0, 0); // Recalc layout
        }
        break;

        case ID_TAB_COPY_PATH:
            if (g_Document.szFileName[0])
            {
                if (OpenClipboard(hwnd))
                {
                    EmptyClipboard();
                    size_t len = (_tcslen(g_Document.szFileName) + 1) * sizeof(TCHAR);
                    HGLOBAL hGlob = GlobalAlloc(GMEM_MOVEABLE, len);
                    if (hGlob)
                    {
                        memcpy(GlobalLock(hGlob), g_Document.szFileName, len);
                        GlobalUnlock(hGlob);
                        SetClipboardData(CF_UNICODETEXT, hGlob);
                    }
                    CloseClipboard();
                }
            }
            break;
        case ID_TAB_COPY_RELATIVE_PATH:
            if (g_Document.szFileName[0])
            {
                TCHAR szPath[MAX_PATH];
                TCHAR szCurrentDir[MAX_PATH];
                GetCurrentDirectory(MAX_PATH, szCurrentDir);
                
                if (PathRelativePathTo(szPath, szCurrentDir, FILE_ATTRIBUTE_DIRECTORY, g_Document.szFileName, FILE_ATTRIBUTE_NORMAL))
                {
                    if (OpenClipboard(hwnd))
                    {
                        EmptyClipboard();
                        size_t len = (_tcslen(szPath) + 1) * sizeof(TCHAR);
                        HGLOBAL hGlob = GlobalAlloc(GMEM_MOVEABLE, len);
                        if (hGlob)
                        {
                            memcpy(GlobalLock(hGlob), szPath, len);
                            GlobalUnlock(hGlob);
                            SetClipboardData(CF_UNICODETEXT, hGlob);
                        }
                        CloseClipboard();
                    }
                }
            }
            break;
        case ID_TAB_REVEAL_IN_EXPLORER:
            if (g_Document.szFileName[0])
            {
                TCHAR szParam[MAX_PATH + 10];
                StringCchPrintf(szParam, MAX_PATH + 10, _T("/select,\"%s\""), g_Document.szFileName);
                ShellExecute(NULL, _T("open"), _T("explorer.exe"), szParam, NULL, SW_SHOW);
            }
            break;
        case ID_TAB_OPEN_NEW_WINDOW:
            if (g_Document.szFileName[0])
            {
                TCHAR szExe[MAX_PATH];
                GetModuleFileName(NULL, szExe, MAX_PATH);
                TCHAR szParam[MAX_PATH + 2];
                StringCchPrintf(szParam, MAX_PATH + 2, _T("\"%s\""), g_Document.szFileName);
                ShellExecute(NULL, _T("open"), szExe, szParam, NULL, SW_SHOW);
            }
            break;
        }
        
        // Handle Plugins
        if (LOWORD(wParam) >= ID_PLUGIN_FIRST && LOWORD(wParam) <= ID_PLUGIN_LAST) {
            g_PluginManager.ExecutePluginCommand(LOWORD(wParam), g_Document.hEditor);
        }

        // Handle Recent Files
        if (LOWORD(wParam) >= ID_FILE_RECENT_FIRST && LOWORD(wParam) < ID_FILE_RECENT_LAST)
        {
            int index = LOWORD(wParam) - ID_FILE_RECENT_FIRST;
            if (index >= 0 && index < recentFiles.size())
            {
                if (CheckSaveChanges())
                {
                    LoadFromFile(recentFiles[index].c_str());
                }
            }
        }
        else if (LOWORD(wParam) == ID_FILE_RECENT_LAST) // Clear
        {
            recentFiles.clear();
            UpdateRecentFilesMenu();
        }
    }
    break;
    case WM_CLOSE:
        if (CheckAllSaveChanges()) 
        {
            SaveConfig();
            DestroyWindow(hwnd);
        }
        return 0;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void SetStatusProgress(int percent) {
    if (hMain) {
        SendMessage(hMain, WM_UPDATE_PROGRESS, percent, 0);
    }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    hInst = hInstance;
    LoadLibrary(_T("Msftedit.dll"));
    
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_BAR_CLASSES | ICC_LINK_CLASS;
    InitCommonControlsEx(&icex);

    WNDCLASSEX wc = {0};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = _T("JustNotepadClass");
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON));

    if (!RegisterClassEx(&wc))
        return 1;

    HWND hwnd = CreateWindowEx(0, wc.lpszClassName, _T("Just Notepad"),
                               WS_OVERLAPPEDWINDOW,
                               CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
                               NULL, NULL, hInstance, NULL);

    if (!hwnd)
        return 1;

    // Load Menu and Accelerators from Resource
    HMENU hMenu = LoadMenu(hInstance, MAKEINTRESOURCE(IDR_MAINMENU));
    
    // Populate Encoding Menu
    HMENU hFileMenu = GetSubMenu(hMenu, 0);
    HMENU hEncodingMenu = GetSubMenu(hFileMenu, 10); // Index 10 is Encoding
    if (hEncodingMenu)
    {
        while (GetMenuItemCount(hEncodingMenu) > 0) DeleteMenu(hEncodingMenu, 0, MF_BYPOSITION);
        for (const auto& info : g_Encodings) AppendMenu(hEncodingMenu, MF_STRING, info.cmdId, info.name);
    }

    SetMenu(hwnd, hMenu);
    
    // Load Plugins
    TCHAR szExePath[MAX_PATH];
    GetModuleFileName(NULL, szExePath, MAX_PATH);
    PathRemoveFileSpec(szExePath);
    PathAppend(szExePath, _T("plugins"));
    
    HostFunctions hostFuncs;
    hostFuncs.SetProgress = SetStatusProgress;
    g_PluginManager.SetHostFunctions(hostFuncs);
    
    g_PluginManager.LoadPlugins(szExePath);
    UpdatePluginsMenu(hwnd);

    LoadConfig();

    HACCEL hAccel = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR));

    ShowWindow(hwnd, SW_SHOWMAXIMIZED);
    UpdateWindow(hwnd);

    int argc;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv && argc > 1)
    {
        LoadFromFile(argv[1]);
        AddRecentFile(argv[1]);
    }
    if (argv) LocalFree(argv);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        if (g_PluginManager.TranslateAccelerator(&msg)) continue;

        if (hFindReplaceDlg == NULL || !IsDialogMessage(hFindReplaceDlg, &msg))
        {
            if (!TranslateAccelerator(hwnd, hAccel, &msg))
            {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }

    return (int)msg.wParam;
}
