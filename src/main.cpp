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
#include <regex>
#include <shlwapi.h>
#include <shellapi.h>
#include <shlobj.h>
#include <deque>
#include <dwmapi.h>
#include <uxtheme.h>
#include "resource.h"
#include "TextHelpers.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")

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
    BOOL bHexMode;
    BOOL bPinned;
    FILETIME ftLastWriteTime;
    Encoding currentEncoding;
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

// Config
std::deque<std::wstring> recentFiles;
TCHAR szConfigPath[MAX_PATH] = {0};
std::wstring g_StatusText[6]; // Cache for owner-draw status bar

// Encoding currentEncoding = ENC_UTF8; // Moved to Document

#define IDT_BACKGROUND_LOAD 999
HWND hProgressBarStatus = NULL;

struct BackgroundLoadState {
    BOOL bActive;
    HANDLE hFile;
    HANDLE hMap;
    LPBYTE pData;
    DWORD dwTotalSize;
    DWORD dwCurrentPos;
    BOOL bHex;
    Encoding encoding;
    std::vector<char> inBuffer;
    DWORD dwOffset;
    DWORD dwOriginalEventMask;
} g_BgLoad;

class ProgressHelper {
    HWND hProgressDlg;
    HWND hProgressBar;
    HWND hParent;
    
public:
    ProgressHelper(HWND parent, const TCHAR* title, int maxRange) : hParent(parent) {
        if (maxRange >= 5 * 1024 * 1024 && maxRange < 10 * 1024 * 1024) // Only show modal for 5MB <= size < 10MB
        {
            hProgressDlg = CreateWindowEx(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, _T("#32770"), title,
                WS_POPUP | WS_CAPTION | WS_VISIBLE | DS_MODALFRAME, 
                0, 0, 300, 80, parent, NULL, hInst, NULL);
                
            RECT rcParent, rcDlg;
            GetWindowRect(parent, &rcParent);
            GetWindowRect(hProgressDlg, &rcDlg);
            int x = rcParent.left + (rcParent.right - rcParent.left - (rcDlg.right - rcDlg.left)) / 2;
            int y = rcParent.top + (rcParent.bottom - rcParent.top - (rcDlg.bottom - rcDlg.top)) / 2;
            SetWindowPos(hProgressDlg, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
            
            hProgressBar = CreateWindowEx(0, PROGRESS_CLASS, NULL, WS_CHILD | WS_VISIBLE,
                10, 10, 265, 20, hProgressDlg, NULL, hInst, NULL);
                
            SendMessage(hProgressBar, PBM_SETRANGE32, 0, maxRange);
            SendMessage(hProgressBar, PBM_SETSTEP, 1, 0);
            
            EnableWindow(parent, FALSE);
            UpdateWindow(hProgressDlg);
            
            MSG msg;
            while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
        else
        {
            hProgressDlg = NULL;
            hProgressBar = NULL;
        }
    }
    
    ~ProgressHelper() {
        if (hProgressDlg)
        {
            EnableWindow(hParent, TRUE);
            DestroyWindow(hProgressDlg);
            SetFocus(hParent);
        }
    }
    
    void Update(int pos) {
        if (hProgressBar)
        {
            SendMessage(hProgressBar, PBM_SETPOS, pos, 0);
            MSG msg;
            while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }
};

struct StreamContext {
    HANDLE hFile;
    BOOL bHex;
    std::vector<char> inBuffer;
    size_t inBufferPos;
    DWORD dwOffset;
    std::string outBuffer;
    ProgressHelper* pProgress;
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
    if (ctx->pProgress) {
        ctx->pProgress->Update(ctx->dwReadSoFar);
    }
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
        if (ctx->bHex)
        {
            LONG bytesCopied = 0;
            while (bytesCopied < cb)
            {
                if (ctx->inBufferPos >= ctx->inBuffer.size())
                {
                    ctx->inBuffer.clear();
                    ctx->inBufferPos = 0;

                    if (ctx->dwMapPos >= ctx->dwTotalSize) break; // EOF

                    DWORD bytesToProcess = min(4096, ctx->dwTotalSize - ctx->dwMapPos);
                    std::string batchOutput;
                    batchOutput.reserve(bytesToProcess * 5);

                    DWORD processed = 0;
                    while (processed < bytesToProcess)
                    {
                        DWORD chunk = min(16, bytesToProcess - processed);
                        BYTE* buf = ctx->pData + ctx->dwMapPos + processed;
                        
                        FormatHexLine(ctx->dwOffset, buf, chunk, batchOutput);
                        
                        ctx->dwOffset += chunk;
                        processed += chunk;
                    }
                    
                    ctx->dwMapPos += processed;
                    UpdateStreamProgress(ctx, processed);
                    
                    ctx->inBuffer.insert(ctx->inBuffer.end(), batchOutput.begin(), batchOutput.end());
                }
                
                size_t available = ctx->inBuffer.size() - ctx->inBufferPos;
                LONG toCopy = min((LONG)available, cb - bytesCopied);
                memcpy(pbBuff + bytesCopied, ctx->inBuffer.data() + ctx->inBufferPos, toCopy);
                
                ctx->inBufferPos += toCopy;
                bytesCopied += toCopy;
            }
            *pcb = bytesCopied;
            return 0;
        }
        else
        {
            if (ctx->dwMapPos >= ctx->dwTotalSize) return 0; // EOF

            LONG toCopy = min(cb, (LONG)(ctx->dwTotalSize - ctx->dwMapPos));
            memcpy(pbBuff, ctx->pData + ctx->dwMapPos, toCopy);
            ctx->dwMapPos += toCopy;
            *pcb = toCopy;

            UpdateStreamProgress(ctx, toCopy);
            return 0;
        }
    }

    if (!ctx->bHex)
    {
        if (ReadFile(ctx->hFile, pbBuff, cb, (LPDWORD)pcb, NULL))
        {
            UpdateStreamProgress(ctx, *pcb);
            return 0;
        }
        return 1; // Error
    }
    else
    {
        LONG bytesCopied = 0;
        while (bytesCopied < cb)
        {
            if (ctx->inBufferPos >= ctx->inBuffer.size())
            {
                ctx->inBuffer.clear();
                ctx->inBufferPos = 0;

                BYTE buf[4096];
                DWORD bytesRead = 0;
                if (!ReadFile(ctx->hFile, buf, 4096, &bytesRead, NULL) || bytesRead == 0)
                {
                    break; // EOF or Error
                }
                
                UpdateStreamProgress(ctx, bytesRead);

                std::string batchOutput;
                batchOutput.reserve(bytesRead * 5);

                DWORD processed = 0;
                while (processed < bytesRead)
                {
                    DWORD chunk = min(16, bytesRead - processed);
                    FormatHexLine(ctx->dwOffset, buf + processed, chunk, batchOutput);
                    ctx->dwOffset += chunk;
                    processed += chunk;
                }
                
                ctx->inBuffer.insert(ctx->inBuffer.end(), batchOutput.begin(), batchOutput.end());
            }
            
            size_t available = ctx->inBuffer.size() - ctx->inBufferPos;
            LONG toCopy = min((LONG)available, cb - bytesCopied);
            memcpy(pbBuff + bytesCopied, ctx->inBuffer.data() + ctx->inBufferPos, toCopy);
            
            ctx->inBufferPos += toCopy;
            bytesCopied += toCopy;
        }
        *pcb = bytesCopied;
        return 0;
    }
}

// Stream callback for writing file
DWORD CALLBACK StreamOutCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
    StreamContext* ctx = (StreamContext*)dwCookie;
    *pcb = cb; // Assume success for now

    if (!ctx->bHex)
    {
        if (WriteFile(ctx->hFile, pbBuff, cb, (LPDWORD)pcb, NULL))
        {
            return 0;
        }
        return 1; // Error
    }
    else
    {
        ctx->outBuffer.append((char*)pbBuff, cb);
        
        size_t pos = 0;
        size_t nextPos;
        while ((nextPos = ctx->outBuffer.find('\n', pos)) != std::string::npos)
        {
            std::string line = ctx->outBuffer.substr(pos, nextPos - pos + 1);
            pos = nextPos + 1;
            
            // Parse line
            // Format: XXXXXXXX  HH HH ...
            std::vector<BYTE> bytes;
            if (ParseHexLine(line, bytes))
            {
                DWORD written;
                WriteFile(ctx->hFile, bytes.data(), (DWORD)bytes.size(), &written, NULL);
            }
        }
        
        ctx->outBuffer.erase(0, pos);
        return 0;
    }
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

    // Stop any active background load
    if (g_BgLoad.bActive)
    {
        KillTimer(hMain, IDT_BACKGROUND_LOAD);
        g_BgLoad.bActive = FALSE;
        if (g_BgLoad.pData) UnmapViewOfFile(g_BgLoad.pData);
        if (g_BgLoad.hMap) CloseHandle(g_BgLoad.hMap);
        CloseHandle(g_BgLoad.hFile);
        ShowWindow(hProgressBarStatus, SW_HIDE);
        doc.bLoading = FALSE;
        
        // Restore Event Mask
        SendMessage(doc.hEditor, EM_SETEVENTMASK, 0, g_BgLoad.dwOriginalEventMask);
        
        // Restore ReadOnly based on current file attribute
        DWORD dwAttrs = GetFileAttributes(doc.szFileName);
        BOOL bReadOnly = (dwAttrs != INVALID_FILE_ATTRIBUTES) && (dwAttrs & FILE_ATTRIBUTE_READONLY);
        SendMessage(doc.hEditor, EM_SETREADONLY, bReadOnly, 0);
        
        // Re-enable ReadOnly menu item
        HMENU hMenu = GetMenu(hMain);
        EnableMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, MF_ENABLED);
    }

    HANDLE hFile = CreateFile(pszFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile != INVALID_HANDLE_VALUE)
    {
        DWORD fileSize = GetFileSize(hFile, NULL);
        
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
                    if (pData[i] == 0)
                    {
                        isBinary = TRUE;
                        break;
                    }
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
                        if (buf[i] == 0)
                        {
                            isBinary = TRUE;
                            break;
                        }
                    }
                }
            }
            SetFilePointer(hFile, 0, NULL, FILE_BEGIN);
        }
        
        doc.bHexMode = isBinary;
        
        // Background Loading for Large Files (> 10MB)
        BOOL bBackgroundLoad = (fileSize > 10 * 1024 * 1024);
        
        ProgressHelper progress(hMain, _T("Loading File..."), fileSize);

        StreamContext ctx = {0};
        ctx.hFile = hFile;
        ctx.bHex = doc.bHexMode;
        ctx.dwOffset = 0;
        ctx.pProgress = bBackgroundLoad ? NULL : &progress; // Don't use modal progress for bg load
        ctx.dwTotalSize = bBackgroundLoad ? min(64 * 1024, fileSize) : fileSize; // Load first 64KB
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
            // Setup background loading
            g_BgLoad.bActive = TRUE;
            g_BgLoad.hFile = hFile;
            g_BgLoad.hMap = hMap;
            g_BgLoad.pData = pData;
            g_BgLoad.dwTotalSize = fileSize;
            g_BgLoad.dwCurrentPos = ctx.dwMapPos;
            g_BgLoad.bHex = doc.bHexMode;
            g_BgLoad.encoding = doc.currentEncoding;
            g_BgLoad.inBuffer = ctx.inBuffer;
            g_BgLoad.dwOffset = ctx.dwOffset;
            
            // Disable events and editing for speed
            g_BgLoad.dwOriginalEventMask = SendMessage(doc.hEditor, EM_GETEVENTMASK, 0, 0);
            SendMessage(doc.hEditor, EM_SETEVENTMASK, 0, 0);
            SendMessage(doc.hEditor, EM_SETREADONLY, TRUE, 0);
            
            // Show Progress Bar
            SendMessage(hProgressBarStatus, PBM_SETRANGE32, 0, fileSize);
            SendMessage(hProgressBarStatus, PBM_SETPOS, ctx.dwMapPos, 0);
            ShowWindow(hProgressBarStatus, SW_SHOW);
            
            // Start Timer - Give it 100ms to render the first screen before blocking
            SetTimer(hMain, IDT_BACKGROUND_LOAD, 100, NULL);
        }
        else
        {
            if (pData) UnmapViewOfFile(pData);
            if (hMap) CloseHandle(hMap);
            CloseHandle(hFile);
            doc.bLoading = FALSE;
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
        return TRUE;
    }
    return FALSE;
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
        if (CheckSaveChanges())
        {
            CreateNewDocument(ofn.lpstrFile);
            AddRecentFile(ofn.lpstrFile);
        }
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
            ctx.bHex = doc.bHexMode;
            
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
            ctx.bHex = doc.bHexMode;
            
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
    if (msg == WM_KEYDOWN)
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
                
                if (lineStart != lineEnd)
                {
                    DoIndent();
                    return 0;
                }
            }
        }
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
                                     WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_NOHIDESEL,
                                     0, 0, rcClient.right, rcClient.bottom - nStatusHeight,
                                     hMain, (HMENU)IDC_EDITOR, hInst, NULL);
        hEditor = g_Document.hEditor;
        
        SendMessage(g_Document.hEditor, EM_EXLIMITTEXT, 0, -1);
        SendMessage(g_Document.hEditor, EM_SETEVENTMASK, 0, ENM_SELCHANGE | ENM_CHANGE);
        
        HFONT hFont = CreateFont(18, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, 
                                     OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                                     DEFAULT_PITCH | FF_SWISS, _T("Consolas"));
        SendMessage(g_Document.hEditor, WM_SETFONT, (WPARAM)hFont, TRUE);

        // Subclass Editor
        wpOrigEditProc = (WNDPROC)SetWindowLongPtr(g_Document.hEditor, GWLP_WNDPROC, (LONG_PTR)EditorWndProc);
    }
    else
    {
        SetWindowText(g_Document.hEditor, _T(""));
    }

    g_Document.bIsDirty = FALSE;
    g_Document.bLoading = FALSE;
    g_Document.bHexMode = FALSE;
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
    
    TCHAR szTitle[MAX_PATH + 32];
    if (doc.szFileName[0])
        StringCchPrintf(szTitle, MAX_PATH + 32, _T("%s%s - Just Notepad"), PathFindFileName(doc.szFileName), doc.bIsDirty ? _T("*") : _T(""));
    else
        StringCchPrintf(szTitle, MAX_PATH + 32, _T("Untitled%s - Just Notepad"), doc.bIsDirty ? _T("*") : _T(""));
    SetWindowText(hMain, szTitle);
}

BOOL IsAppInShellMenu()
{
    HKEY hKey;
    if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\*\\shell\\JustNotepad"), 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        RegCloseKey(hKey);
        return TRUE;
    }
    return FALSE;
}

void AddAppToShellMenu()
{
    TCHAR szPath[MAX_PATH];
    GetModuleFileName(NULL, szPath, MAX_PATH);
    
    HKEY hKey;
    if (RegCreateKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\*\\shell\\JustNotepad"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("Edit in Just Notepad"), (DWORD)(_tcslen(_T("Edit in Just Notepad")) + 1) * sizeof(TCHAR));
        RegSetValueEx(hKey, _T("Icon"), 0, REG_SZ, (BYTE*)szPath, (DWORD)(_tcslen(szPath) + 1) * sizeof(TCHAR));
        
        HKEY hKeyCmd;
        if (RegCreateKeyEx(hKey, _T("command"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyCmd, NULL) == ERROR_SUCCESS)
        {
            TCHAR szCmd[MAX_PATH + 10];
            StringCchPrintf(szCmd, MAX_PATH + 10, _T("\"%s\" \"%%1\""), szPath);
            RegSetValueEx(hKeyCmd, NULL, 0, REG_SZ, (BYTE*)szCmd, (DWORD)(_tcslen(szCmd) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyCmd);
        }
        RegCloseKey(hKey);
    }
    
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
}

void RemoveAppFromShellMenu()
{
    // RegDeleteTree is available on Vista+
    // For older systems, we might need a recursive delete function, but let's assume Vista+
    HMODULE hAdvApi32 = LoadLibrary(_T("Advapi32.dll"));
    if (hAdvApi32)
    {
        typedef LSTATUS (APIENTRY *PFN_RegDeleteTree)(HKEY, LPCTSTR);
        PFN_RegDeleteTree pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeW");
        #ifdef UNICODE
        if (!pfnRegDeleteTree) pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeW");
        #else
        if (!pfnRegDeleteTree) pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeA");
        #endif
        
        if (pfnRegDeleteTree)
        {
            pfnRegDeleteTree(HKEY_CURRENT_USER, _T("Software\\Classes\\*\\shell\\JustNotepad"));
        }
        FreeLibrary(hAdvApi32);
    }
}

BOOL IsTxtAssociated();

void DisassociateTxtFiles()
{
    HMODULE hAdvApi32 = LoadLibrary(_T("Advapi32.dll"));
    if (hAdvApi32)
    {
        typedef LSTATUS (APIENTRY *PFN_RegDeleteTree)(HKEY, LPCTSTR);
        PFN_RegDeleteTree pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeW");
        #ifdef UNICODE
        if (!pfnRegDeleteTree) pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeW");
        #else
        if (!pfnRegDeleteTree) pfnRegDeleteTree = (PFN_RegDeleteTree)GetProcAddress(hAdvApi32, "RegDeleteTreeA");
        #endif
        
        if (pfnRegDeleteTree)
        {
            // Check if .txt is associated with JustNotepad.txt
            if (IsTxtAssociated())
            {
                // Delete .txt key (or just the default value? Deleting the key might be too aggressive if there are other subkeys)
                // Actually, we should just delete the default value or set it to empty.
                // But to be clean, let's try to delete the key if it only contains our association.
                // For safety, let's just delete the default value.
                HKEY hKey;
                if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\.txt"), 0, KEY_WRITE, &hKey) == ERROR_SUCCESS)
                {
                    RegDeleteValue(hKey, NULL);
                    RegCloseKey(hKey);
                }

                // Delete JustNotepad.txt ProgID
                pfnRegDeleteTree(HKEY_CURRENT_USER, _T("Software\\Classes\\JustNotepad.txt"));
            }
        }
        FreeLibrary(hAdvApi32);
    }
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
}

BOOL IsTxtAssociated()
{
    HKEY hKey;
    TCHAR szValue[256];
    DWORD dwSize = sizeof(szValue);
    if (RegOpenKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\.txt"), 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        if (RegQueryValueEx(hKey, NULL, NULL, NULL, (LPBYTE)szValue, &dwSize) == ERROR_SUCCESS)
        {
            RegCloseKey(hKey);
            return _tcscmp(szValue, _T("JustNotepad.txt")) == 0;
        }
        RegCloseKey(hKey);
    }
    return FALSE;
}

BOOL IsUserAdmin()
{
    BOOL b;
    SID_IDENTIFIER_AUTHORITY NtAuthority = SECURITY_NT_AUTHORITY;
    PSID AdministratorsGroup; 
    b = AllocateAndInitializeSid(&NtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &AdministratorsGroup); 
    if(b) 
    {
        if (!CheckTokenMembership( NULL, AdministratorsGroup, &b)) 
        {
             b = FALSE;
        } 
        FreeSid(AdministratorsGroup); 
    }

    return b;
}

#ifndef MB_ICONSHIELD
#define MB_ICONSHIELD 0x0000000CL // Vista+
#endif

void AssociateTxtFiles()
{
    if (!IsUserAdmin())
    {
        int result = MessageBox(hMain, _T("This feature requires Administrator privileges to set system-wide associations.\nDo you want to restart Just Notepad as Administrator?"), _T("Administrator Required"), MB_YESNO | MB_ICONINFORMATION);
        if (result == IDYES)
        {
            TCHAR szPath[MAX_PATH];
            GetModuleFileName(NULL, szPath, MAX_PATH);
            ShellExecute(NULL, _T("runas"), szPath, NULL, NULL, SW_SHOWNORMAL);
            SendMessage(hMain, WM_CLOSE, 0, 0);
        }
        return;
    }

    TCHAR szPath[MAX_PATH];
    GetModuleFileName(NULL, szPath, MAX_PATH);

    HKEY hKey;
    // .txt -> JustNotepad.txt (HKCR)
    if (RegCreateKeyEx(HKEY_CLASSES_ROOT, _T(".txt"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("JustNotepad.txt"), (DWORD)(_tcslen(_T("JustNotepad.txt")) + 1) * sizeof(TCHAR));
        RegCloseKey(hKey);
    }

    // JustNotepad.txt -> Text Document (HKCR)
    if (RegCreateKeyEx(HKEY_CLASSES_ROOT, _T("JustNotepad.txt"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("Text Document"), (DWORD)(_tcslen(_T("Text Document")) + 1) * sizeof(TCHAR));
        
        // DefaultIcon
        HKEY hKeyIcon;
        if (RegCreateKeyEx(hKey, _T("DefaultIcon"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyIcon, NULL) == ERROR_SUCCESS)
        {
            TCHAR szIcon[MAX_PATH + 5];
            StringCchPrintf(szIcon, MAX_PATH + 5, _T("%s,0"), szPath);
            RegSetValueEx(hKeyIcon, NULL, 0, REG_SZ, (BYTE*)szIcon, (DWORD)(_tcslen(szIcon) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyIcon);
        }

        // shell\open\command
        HKEY hKeyShell;
        if (RegCreateKeyEx(hKey, _T("shell\\open\\command"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyShell, NULL) == ERROR_SUCCESS)
        {
            TCHAR szCmd[MAX_PATH + 10];
            StringCchPrintf(szCmd, MAX_PATH + 10, _T("\"%s\" \"%%1\""), szPath);
            RegSetValueEx(hKeyShell, NULL, 0, REG_SZ, (BYTE*)szCmd, (DWORD)(_tcslen(szCmd) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyShell);
        }
        RegCloseKey(hKey);
    }

    // Also do HKCU for good measure (per-user override)
    if (RegCreateKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\.txt"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("JustNotepad.txt"), (DWORD)(_tcslen(_T("JustNotepad.txt")) + 1) * sizeof(TCHAR));
        RegCloseKey(hKey);
    }

    if (RegCreateKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\JustNotepad.txt"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE*)_T("Text Document"), (DWORD)(_tcslen(_T("Text Document")) + 1) * sizeof(TCHAR));
        
        HKEY hKeyIcon;
        if (RegCreateKeyEx(hKey, _T("DefaultIcon"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyIcon, NULL) == ERROR_SUCCESS)
        {
            TCHAR szIcon[MAX_PATH + 5];
            StringCchPrintf(szIcon, MAX_PATH + 5, _T("%s,0"), szPath);
            RegSetValueEx(hKeyIcon, NULL, 0, REG_SZ, (BYTE*)szIcon, (DWORD)(_tcslen(szIcon) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyIcon);
        }

        HKEY hKeyShell;
        if (RegCreateKeyEx(hKey, _T("shell\\open\\command"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyShell, NULL) == ERROR_SUCCESS)
        {
            TCHAR szCmd[MAX_PATH + 10];
            StringCchPrintf(szCmd, MAX_PATH + 10, _T("\"%s\" \"%%1\""), szPath);
            RegSetValueEx(hKeyShell, NULL, 0, REG_SZ, (BYTE*)szCmd, (DWORD)(_tcslen(szCmd) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyShell);
        }
        RegCloseKey(hKey);
    }

    // Register in RegisteredApplications
    if (RegCreateKeyEx(HKEY_CURRENT_USER, _T("Software\\JustNotepad\\Capabilities"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, _T("ApplicationDescription"), 0, REG_SZ, (BYTE*)_T("Just Notepad"), (DWORD)(_tcslen(_T("Just Notepad")) + 1) * sizeof(TCHAR));
        RegSetValueEx(hKey, _T("ApplicationName"), 0, REG_SZ, (BYTE*)_T("Just Notepad"), (DWORD)(_tcslen(_T("Just Notepad")) + 1) * sizeof(TCHAR));
        
        HKEY hKeyExt;
        if (RegCreateKeyEx(hKey, _T("FileAssociations"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKeyExt, NULL) == ERROR_SUCCESS)
        {
            RegSetValueEx(hKeyExt, _T(".txt"), 0, REG_SZ, (BYTE*)_T("JustNotepad.txt"), (DWORD)(_tcslen(_T("JustNotepad.txt")) + 1) * sizeof(TCHAR));
            RegCloseKey(hKeyExt);
        }
        RegCloseKey(hKey);
    }
    
    if (RegCreateKeyEx(HKEY_CURRENT_USER, _T("Software\\RegisteredApplications"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, _T("JustNotepad"), 0, REG_SZ, (BYTE*)_T("Software\\JustNotepad\\Capabilities"), (DWORD)(_tcslen(_T("Software\\JustNotepad\\Capabilities")) + 1) * sizeof(TCHAR));
        RegCloseKey(hKey);
    }
    
    // Add to OpenWithProgids
    if (RegCreateKeyEx(HKEY_CURRENT_USER, _T("Software\\Classes\\.txt\\OpenWithProgids"), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueEx(hKey, _T("JustNotepad.txt"), 0, REG_NONE, NULL, 0);
        RegCloseKey(hKey);
    }
    
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
    
    // Try to launch Open With dialog to let user pick default
    // Create a dummy file
    TCHAR szDummy[MAX_PATH];
    GetTempPath(MAX_PATH, szDummy);
    StringCchCat(szDummy, MAX_PATH, _T("JustNotepad_Assoc.txt"));
    HANDLE hFile = CreateFile(szDummy, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile != INVALID_HANDLE_VALUE) CloseHandle(hFile);
    
    // Launch Open With
    SHELLEXECUTEINFO sei = { sizeof(sei) };
    sei.fMask = SEE_MASK_INVOKEIDLIST;
    sei.lpVerb = _T("openas");
    sei.lpFile = szDummy;
    sei.nShow = SW_SHOWNORMAL;
    ShellExecuteEx(&sei);
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
    GETTEXTLENGTHEX gtlSize = { GTL_NUMBYTES | GTL_PRECISE, info->codePage };
    long nBytes = SendMessage(hEditor, EM_GETTEXTLENGTHEX, (WPARAM)&gtlSize, 0);
    
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

void DoOpenFolder()
{
    Document& doc = g_Document;

    if (doc.szFileName[0])
    {
        TCHAR szDir[MAX_PATH];
        StringCchCopy(szDir, MAX_PATH, doc.szFileName);
        PathRemoveFileSpec(szDir);
        ShellExecute(NULL, _T("explore"), szDir, NULL, NULL, SW_SHOW);
    }
}

void DoOpenCmd()
{
    Document& doc = g_Document;

    if (doc.szFileName[0])
    {
        TCHAR szDir[MAX_PATH];
        StringCchCopy(szDir, MAX_PATH, doc.szFileName);
        PathRemoveFileSpec(szDir);
        
        TCHAR szCmdLine[MAX_PATH];
        if (GetEnvironmentVariable(_T("ComSpec"), szCmdLine, MAX_PATH) == 0)
        {
            StringCchCopy(szCmdLine, MAX_PATH, _T("cmd.exe"));
        }
        
        STARTUPINFO si = { sizeof(si) };
        PROCESS_INFORMATION pi = { 0 };
        
        if (CreateProcess(NULL, szCmdLine, NULL, NULL, FALSE, CREATE_NEW_CONSOLE, NULL, szDir, &si, &pi))
        {
            WaitForInputIdle(pi.hProcess, 2000);
            Sleep(200); // Wait for window to be ready
            
            LPCTSTR pszFile = PathFindFileName(doc.szFileName);
            size_t len = _tcslen(pszFile);
            std::vector<INPUT> inputs;
            inputs.reserve(len * 2);
            
            for (size_t i = 0; i < len; i++)
            {
                INPUT input = { 0 };
                input.type = INPUT_KEYBOARD;
                input.ki.wScan = (WORD)pszFile[i];
                input.ki.dwFlags = KEYEVENTF_UNICODE;
                inputs.push_back(input);
                
                input.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
                inputs.push_back(input);
            }
            
            SendInput((UINT)inputs.size(), inputs.data(), sizeof(INPUT));
            
            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);
        }
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

void DoOpenDefault()
{
    Document& doc = g_Document;

    if (doc.szFileName[0])
    {
        ShellExecute(NULL, _T("open"), doc.szFileName, NULL, NULL, SW_SHOW);
    }
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
    
    // Restore original selection offsets (accounting for added indentation)
    long offsetStart = cr.cpMin - r.start;
    long offsetEnd = cr.cpMax - r.start;
    CHARRANGE crNew = {r.start + offsetStart, r.start + offsetEnd};
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
    
    // Restore original selection offsets (accounting for removed indentation)
    long offsetStart = cr.cpMin - r.start;
    long offsetEnd = cr.cpMax - r.start;
    CHARRANGE crNew = {r.start + offsetStart, r.start + offsetEnd};
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

// Advanced Search Logic
struct SearchMatch {
    int lineNum; // 0-based
    int matchStart; // Index in line
    int matchLen;
    int score;
    long absolutePos; // Absolute character position
    std::basic_string<TCHAR> lineText; // Cache the line text
};

bool FuzzyMatch(const std::string& text, const std::string& query, int& score) {
    if (query.empty()) return true;
    
    int tIdx = 0;
    int qIdx = 0;
    int tLen = (int)text.length();
    int qLen = (int)query.length();
    
    score = 0;
    int consecutive = 0;
    
    while (tIdx < tLen && qIdx < qLen) {
        if (tolower(text[tIdx]) == tolower(query[qIdx])) {
            score += 10;
            score += consecutive * 5;
            consecutive++;
            qIdx++;
        } else {
            consecutive = 0;
            score -= 1;
        }
        tIdx++;
    }
    
    return qIdx == qLen;
}

HFONT hDialogListFont = NULL;

struct SearchState {
    BOOL bActive;
    std::vector<TCHAR> buffer;
    std::basic_string<TCHAR> query;
    BOOL bRegex;
    BOOL bFuzzy;
    long currentOffset;
    int lineNum;
    long totalLen;
    
    // For incremental processing
    std::basic_stringstream<TCHAR> ss;
    
    // Results
    std::vector<SearchMatch>* pMatches;
    HWND hDlg;
    int itemsPerPage;
    int* pCurrentPage;
};

SearchState g_SearchState = {0};
#define IDT_SEARCH_BG 1001

void UpdateSearchList(HWND hDlg, const std::vector<SearchMatch>& matches, int page, int itemsPerPage);

void StartSearch(HWND hDlg, std::vector<SearchMatch>& matches, int* pCurrentPage)
{
    TCHAR szQuery[256];
    GetDlgItemText(hDlg, IDC_SEARCH_TEXT, szQuery, 256);
    
    // Stop existing search
    if (g_SearchState.bActive) {
        KillTimer(hDlg, IDT_SEARCH_BG);
        g_SearchState.bActive = FALSE;
    }
    
    g_SearchState.bRegex = IsDlgButtonChecked(hDlg, IDC_SEARCH_REGEX);
    g_SearchState.bFuzzy = IsDlgButtonChecked(hDlg, IDC_SEARCH_FUZZY);
    g_SearchState.query = szQuery;
    g_SearchState.pMatches = &matches;
    g_SearchState.hDlg = hDlg;
    g_SearchState.itemsPerPage = 30; // Max 30 results per page
    g_SearchState.pCurrentPage = pCurrentPage;
    
    int nLen = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    g_SearchState.buffer.resize(nLen + 1);
    SendMessage(hEditor, WM_GETTEXT, nLen + 1, (LPARAM)g_SearchState.buffer.data());
    g_SearchState.totalLen = nLen;
    
    matches.clear();
    *pCurrentPage = 1;

    // Initialize stream
    // Note: stringstream copies the buffer, which is not ideal for huge files but okay for now.
    // For huge files we should use a custom stream or just pointers.
    // But since we already have the buffer in memory...
    g_SearchState.ss.str(std::basic_string<TCHAR>(g_SearchState.buffer.data()));
    g_SearchState.ss.clear();
    g_SearchState.currentOffset = 0;
    g_SearchState.lineNum = 0;
    g_SearchState.bActive = TRUE;
    
    // Reset Progress
    SendDlgItemMessage(hDlg, IDC_SEARCH_PROGRESS, PBM_SETRANGE32, 0, nLen);
    SendDlgItemMessage(hDlg, IDC_SEARCH_PROGRESS, PBM_SETPOS, 0, 0);
    ShowWindow(GetDlgItem(hDlg, IDC_SEARCH_PROGRESS), SW_SHOW);
    
    // Start Timer for background processing
    SetTimer(hDlg, IDT_SEARCH_BG, 10, NULL);
}

void ProcessSearchChunk()
{
    if (!g_SearchState.bActive) return;
    
    std::basic_string<TCHAR> segment;
    
    #ifdef UNICODE
    std::wstring wQuery = g_SearchState.query;
    #else
    std::string sQuery = g_SearchState.query;
    #endif

    // Process a chunk (e.g. 100 lines or 5ms)
    int linesProcessed = 0;
    const int LINES_PER_CHUNK = 100;
    
    while(linesProcessed < LINES_PER_CHUNK && std::getline(g_SearchState.ss, segment))
    {
        size_t lineLenInStream = segment.length() + 1; // +1 for \n
        if (!segment.empty() && segment.back() == '\r') 
        {
            segment.pop_back();
            lineLenInStream = segment.length() + 2;
        }
        else
        {
            lineLenInStream = segment.length() + 1;
        }
        
        if (g_SearchState.ss.eof())
        {
            lineLenInStream--; 
        }

        long internalOffset = g_SearchState.currentOffset - g_SearchState.lineNum;
        BOOL bMatchFound = FALSE;
        
        if (g_SearchState.bRegex)
        {
            try {
                #ifdef UNICODE
                std::wregex re(wQuery, std::regex_constants::icase);
                #else
                std::regex re(sQuery, std::regex_constants::icase);
                #endif
                
                #ifdef UNICODE
                std::wsmatch m;
                auto searchStart = segment.cbegin();
                while (std::regex_search(searchStart, segment.cend(), m, re))
                #else
                std::smatch m;
                auto searchStart = segment.cbegin();
                while (std::regex_search(searchStart, segment.cend(), m, re))
                #endif
                {
                    SearchMatch sm;
                    sm.lineNum = g_SearchState.lineNum;
                    sm.matchStart = (int)(m.position() + (searchStart - segment.cbegin()));
                    sm.matchLen = (int)m.length();
                    sm.score = 0;
                    sm.absolutePos = internalOffset + sm.matchStart;
                    sm.lineText = segment;
                    g_SearchState.pMatches->push_back(sm);
                    bMatchFound = TRUE;
                    
                    searchStart += m.position() + m.length();
                    if (m.length() == 0) searchStart++;
                }
            }
            catch (...) {}
        }
        else if (g_SearchState.bFuzzy)
        {
            int score = 0;
            std::string lineStr;
            std::string queryStr;
            
            #ifdef UNICODE
            for(auto c : segment) lineStr += (char)c;
            for(auto c : wQuery) queryStr += (char)c;
            #else
            lineStr = segment;
            queryStr = sQuery;
            #endif
            
            if (FuzzyMatch(lineStr, queryStr, score))
            {
                SearchMatch m;
                m.lineNum = g_SearchState.lineNum;
                m.matchStart = -1; 
                m.matchLen = 0;
                m.score = score;
                m.absolutePos = internalOffset; // Start of line
                m.lineText = segment;
                g_SearchState.pMatches->push_back(m);
                bMatchFound = TRUE;
            }
        }
        else
        {
            size_t pos = 0;
            std::basic_string<TCHAR> textLower = ToLowerCase(segment);
            std::basic_string<TCHAR> queryLower = ToLowerCase(g_SearchState.query);
            
            while ((pos = textLower.find(queryLower, pos)) != std::basic_string<TCHAR>::npos)
            {
                SearchMatch sm;
                sm.lineNum = g_SearchState.lineNum;
                sm.matchStart = (int)pos;
                sm.matchLen = (int)queryLower.length();
                sm.score = 0;
                sm.absolutePos = internalOffset + sm.matchStart;
                sm.lineText = segment;
                g_SearchState.pMatches->push_back(sm);
                bMatchFound = TRUE;
                pos += queryLower.length();
            }
        }
        
        g_SearchState.currentOffset += lineLenInStream;
        g_SearchState.lineNum++;
        linesProcessed++;
    }
    
    // Update Progress
    SendDlgItemMessage(g_SearchState.hDlg, IDC_SEARCH_PROGRESS, PBM_SETPOS, g_SearchState.currentOffset, 0);
    
    // Update List ONLY if we haven't filled the first page yet
    static BOOL bFirstPageFilled = FALSE;
    if (*g_SearchState.pCurrentPage == 1 && !bFirstPageFilled)
    {
        if (g_SearchState.pMatches->size() >= (size_t)g_SearchState.itemsPerPage)
        {
            UpdateSearchList(g_SearchState.hDlg, *g_SearchState.pMatches, *g_SearchState.pCurrentPage, g_SearchState.itemsPerPage);
            bFirstPageFilled = TRUE;
        }
        else
        {
             // Still filling first page, update it so user sees something
             UpdateSearchList(g_SearchState.hDlg, *g_SearchState.pMatches, *g_SearchState.pCurrentPage, g_SearchState.itemsPerPage);
        }
    }
    
    // Always update page info
    int totalPages = (int)((g_SearchState.pMatches->size() + g_SearchState.itemsPerPage - 1) / g_SearchState.itemsPerPage);
    if (totalPages < 1) totalPages = 1;
    TCHAR pageInfo[64];
    StringCchPrintf(pageInfo, 64, _T("Page %d of %d"), *g_SearchState.pCurrentPage, totalPages);
    SetDlgItemText(g_SearchState.hDlg, IDC_PAGE_INFO, pageInfo);
    EnableWindow(GetDlgItem(g_SearchState.hDlg, IDC_NEXT_PAGE), *g_SearchState.pCurrentPage < totalPages);

    if (g_SearchState.ss.eof())
    {
        // Done
        KillTimer(g_SearchState.hDlg, IDT_SEARCH_BG);
        g_SearchState.bActive = FALSE;
        bFirstPageFilled = FALSE; // Reset for next search
        
        if (g_SearchState.bFuzzy)
        {
            std::sort(g_SearchState.pMatches->begin(), g_SearchState.pMatches->end(), [](const SearchMatch& a, const SearchMatch& b) {
                return a.score > b.score;
            });
            // Refresh list after sort
            UpdateSearchList(g_SearchState.hDlg, *g_SearchState.pMatches, *g_SearchState.pCurrentPage, g_SearchState.itemsPerPage);
        }
        
        // Hide progress bar or show 100%
        // ShowWindow(GetDlgItem(g_SearchState.hDlg, IDC_SEARCH_PROGRESS), SW_HIDE);
    }
}

void UpdateSearchList(HWND hDlg, const std::vector<SearchMatch>& matches, int page, int itemsPerPage)
{
    SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_RESETCONTENT, 0, 0);
    
    if (matches.empty())
    {
        SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_ADDSTRING, 0, (LPARAM)_T("No matches found."));
        SetDlgItemText(hDlg, IDC_PAGE_INFO, _T("Page 0 of 0"));
        EnableWindow(GetDlgItem(hDlg, IDC_PREV_PAGE), FALSE);
        EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), FALSE);
        return;
    }

    int totalPages = (int)((matches.size() + itemsPerPage - 1) / itemsPerPage);
    if (page < 1) page = 1;
    if (page > totalPages) page = totalPages;
    
    int startIdx = (page - 1) * itemsPerPage;
    int endIdx = min((int)matches.size(), startIdx + itemsPerPage);
    
    SendMessage(GetDlgItem(hDlg, IDC_SEARCH_LIST), WM_SETREDRAW, FALSE, 0);

    for (int i = startIdx; i < endIdx; i++)
    {
        const auto& m = matches[i];
        int line = m.lineNum;
        
        {
            TCHAR buf[512];
            StringCchPrintf(buf, 512, _T("%4d: %s"), line + 1, m.lineText.c_str());
            int idx = SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_ADDSTRING, 0, (LPARAM)buf);
            SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_SETITEMDATA, idx, line);
        }
    }

    SendMessage(GetDlgItem(hDlg, IDC_SEARCH_LIST), WM_SETREDRAW, TRUE, 0);
    InvalidateRect(GetDlgItem(hDlg, IDC_SEARCH_LIST), NULL, TRUE);
    
    TCHAR pageInfo[64];
    StringCchPrintf(pageInfo, 64, _T("Page %d of %d"), page, totalPages);
    SetDlgItemText(hDlg, IDC_PAGE_INFO, pageInfo);
    
    EnableWindow(GetDlgItem(hDlg, IDC_PREV_PAGE), page > 1);
    EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), page < totalPages);
}

INT_PTR CALLBACK AdvancedSearchDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    static std::vector<SearchMatch> matches;
    static int currentPage = 1;
    static const int itemsPerPage = 100;

    switch (message)
    {
    case WM_INITDIALOG:
    {
        hDialogListFont = CreateFont(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, 
                                 OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                                 DEFAULT_PITCH | FF_SWISS, _T("Consolas"));
        SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, WM_SETFONT, (WPARAM)hDialogListFont, TRUE);
        
        EnableWindow(GetDlgItem(hDlg, IDC_PREV_PAGE), FALSE);
        EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), FALSE);
        SetDlgItemText(hDlg, IDC_PAGE_INFO, _T(""));
        
        // Resize dialog to 50% of screen and center it
        RECT rcScreen;
        GetWindowRect(GetDesktopWindow(), &rcScreen);
        int screenW = rcScreen.right - rcScreen.left;
        int screenH = rcScreen.bottom - rcScreen.top;
        int width = screenW * 50 / 100;
        int height = screenH * 50 / 100;
        int x = rcScreen.left + (screenW - width) / 2;
        int y = rcScreen.top + (screenH - height) / 2;
        
        SetWindowPos(hDlg, NULL, x, y, width, height, SWP_NOZORDER);
        
        return (INT_PTR)TRUE;
    }
    case WM_SIZE:
    {
        int width = LOWORD(lParam);
        int height = HIWORD(lParam);
        
        int margin = 10;
        int btnWidth = 90;
        int btnHeight = 24;
        int editHeight = 24;
        int labelHeight = 16;
        int labelWidth = 70;
        
        int editX = margin + labelWidth + 5;
        int editW = width - editX - margin - btnWidth - 10;
        int btnX = width - margin - btnWidth;
        
        int row1Y = 10;
        int row2Y = row1Y + editHeight + 10;
        int groupY = row2Y + editHeight + 15;
        int groupH = 45;
        int resultsY = groupY + groupH + 10;
        int progressY = resultsY;
        int listY = resultsY + 20;
        int bottomY = height - btnHeight - 10;
        
        // Top Row: Find
        SetWindowPos(GetDlgItem(hDlg, IDC_STATIC_FIND_LABEL), NULL, margin, row1Y + 4, labelWidth, labelHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_TEXT), NULL, editX, row1Y, editW, editHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_BTN), NULL, btnX, row1Y, btnWidth, btnHeight, SWP_NOZORDER);
        
        // Second Row: Replace
        SetWindowPos(GetDlgItem(hDlg, IDC_STATIC_REPLACE_LABEL), NULL, margin, row2Y + 4, labelWidth, labelHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_REPLACE_TEXT), NULL, editX, row2Y, editW, editHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_REPLACE_ALL_BTN), NULL, btnX, row2Y, btnWidth, btnHeight, SWP_NOZORDER);
        
        // GroupBox
        SetWindowPos(GetDlgItem(hDlg, IDC_STATIC_OPTIONS_GROUP), NULL, margin, groupY, width - margin*2, groupH, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_REGEX), NULL, margin + 10, groupY + 20, 120, 20, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_FUZZY), NULL, margin + 140, groupY + 20, 100, 20, SWP_NOZORDER);
        
        // Progress Bar
        SetWindowPos(GetDlgItem(hDlg, IDC_STATIC_RESULTS_LABEL), NULL, margin, resultsY, 50, labelHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_PROGRESS), NULL, margin + 55, resultsY, width - margin*2 - 55, 15, SWP_NOZORDER);
        
        // ListBox
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_LIST), NULL, margin, listY, width - margin*2, bottomY - listY - 10, SWP_NOZORDER);
        
        // Bottom Controls
        SetWindowPos(GetDlgItem(hDlg, IDC_PREV_PAGE), NULL, margin, bottomY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_PAGE_INFO), NULL, margin + btnWidth + 10, bottomY + 5, 100, labelHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_NEXT_PAGE), NULL, margin + btnWidth + 120, bottomY, btnWidth, btnHeight, SWP_NOZORDER);
        
        SetWindowPos(GetDlgItem(hDlg, IDCANCEL), NULL, width - margin - btnWidth, bottomY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_REPLACE_BTN), NULL, width - margin - btnWidth - 10 - 110, bottomY, 110, btnHeight, SWP_NOZORDER);
        
        return (INT_PTR)TRUE;
    }
    case WM_TIMER:
        if (wParam == IDT_SEARCH_BG)
        {
            ProcessSearchChunk();
        }
        break;
    case WM_DESTROY:
        if (g_SearchState.bActive) {
            KillTimer(hDlg, IDT_SEARCH_BG);
            g_SearchState.bActive = FALSE;
        }
        if (hDialogListFont) DeleteObject(hDialogListFont);
        matches.clear();
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_SEARCH_BTN)
        {
            StartSearch(hDlg, matches, &currentPage);
        }
        else if (LOWORD(wParam) == IDC_PREV_PAGE)
        {
            if (currentPage > 1)
            {
                currentPage--;
                UpdateSearchList(hDlg, matches, currentPage, itemsPerPage);
            }
        }
        else if (LOWORD(wParam) == IDC_NEXT_PAGE)
        {
            int totalPages = (int)((matches.size() + itemsPerPage - 1) / itemsPerPage);
            if (currentPage < totalPages)
            {
                currentPage++;
                UpdateSearchList(hDlg, matches, currentPage, itemsPerPage);
            }
        }
        else if (LOWORD(wParam) == IDC_REPLACE_BTN)
        {
            int idx = SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETCURSEL, 0, 0);
            if (idx != LB_ERR)
            {
                int line = SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETITEMDATA, idx, 0);
                if (line != -1)
                {
                    // Check if fuzzy
                    BOOL bFuzzy = IsDlgButtonChecked(hDlg, IDC_SEARCH_FUZZY);
                    if (bFuzzy) {
                        MessageBox(hDlg, _T("Replace not supported in Fuzzy mode."), _T("Info"), MB_ICONINFORMATION);
                        return TRUE;
                    }

                    for (const auto& m : matches)
                    {
                        if (m.lineNum == line)
                        {
                            // Found match
                            TCHAR szReplace[256];
                            GetDlgItemText(hDlg, IDC_REPLACE_TEXT, szReplace, 256);
                            
                            if (m.matchStart != -1)
                            {
                                SendMessage(hEditor, EM_SETSEL, m.absolutePos, m.absolutePos + m.matchLen);
                                SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)szReplace);
                                
                                // Refresh search
                                StartSearch(hDlg, matches, &currentPage);
                            }
                            break;
                        }
                    }
                }
            }
        }
        else if (LOWORD(wParam) == IDC_REPLACE_ALL_BTN)
        {
            BOOL bFuzzy = IsDlgButtonChecked(hDlg, IDC_SEARCH_FUZZY);
            if (bFuzzy) {
                MessageBox(hDlg, _T("Replace not supported in Fuzzy mode."), _T("Info"), MB_ICONINFORMATION);
                return TRUE;
            }

            // Wait for search to finish if active?
            if (g_SearchState.bActive)
            {
                 MessageBox(hDlg, _T("Please wait for search to complete."), _T("Info"), MB_ICONINFORMATION);
                 return TRUE;
            }
            
            if (matches.empty()) 
            {
                 StartSearch(hDlg, matches, &currentPage);
                 // We can't replace immediately if search is async.
                 // For now, let's just say "Search started, click Replace All again when done"
                 // Or we could make Replace All synchronous.
                 return TRUE;
            }
            
            TCHAR szReplace[256];
            GetDlgItemText(hDlg, IDC_REPLACE_TEXT, szReplace, 256);
            
            // Sort matches by line/position descending to avoid index shifts
            std::sort(matches.begin(), matches.end(), [](const SearchMatch& a, const SearchMatch& b) {
                if (a.lineNum != b.lineNum) return a.lineNum > b.lineNum;
                return a.matchStart > b.matchStart;
            });
            
            int count = 0;
            for (const auto& m : matches)
            {
                if (m.matchStart != -1)
                {
                    SendMessage(hEditor, EM_SETSEL, m.absolutePos, m.absolutePos + m.matchLen);
                    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)szReplace);
                    count++;
                }
            }
            
            TCHAR msg[64];
            StringCchPrintf(msg, 64, _T("Replaced %d occurrences."), count);
            MessageBox(hDlg, msg, _T("Replace All"), MB_OK);
            
            StartSearch(hDlg, matches, &currentPage);
        }
        else if (LOWORD(wParam) == IDC_SEARCH_LIST)
        {
            if (HIWORD(wParam) == LBN_SELCHANGE || HIWORD(wParam) == LBN_DBLCLK)
            {
                int idx = SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETCURSEL, 0, 0);
                if (idx != LB_ERR)
                {
                    int line = SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETITEMDATA, idx, 0);
                    if (line != -1)
                    {
                        int selStart = -1;
                        int selEnd = -1;
                        BOOL bFound = FALSE;

                        // Find match info
                        for(const auto& m : matches) {
                            if (m.lineNum == line) {
                                selStart = m.absolutePos;
                                selEnd = selStart + m.matchLen;
                                bFound = TRUE;
                                break;
                            }
                        }
                        
                        if (bFound)
                        {
                            // Scroll to start
                            SendMessage(hEditor, EM_SETSEL, selStart, selStart);
                            SendMessage(hEditor, EM_SCROLLCARET, 0, 0);
                            
                            // Select full
                            SendMessage(hEditor, EM_SETSEL, selStart, selEnd);
                            SendMessage(hEditor, EM_SCROLLCARET, 0, 0);
                        }
                        else
                        {
                            int nIndex = SendMessage(hEditor, EM_LINEINDEX, line, 0);
                            if (nIndex != -1)
                            {
                                long nLineLen = SendMessage(hEditor, EM_LINELENGTH, nIndex, 0);
                                SendMessage(hEditor, EM_SETSEL, nIndex, nIndex + nLineLen);
                                SendMessage(hEditor, EM_SCROLLCARET, 0, 0);
                            }
                        }
                        
                        if (HIWORD(wParam) == LBN_DBLCLK)
                        {
                            EndDialog(hDlg, IDOK);
                        }
                    }
                }
            }
        }
        else if (LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, IDCANCEL);
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void DoAdvancedSearch()
{
    DialogBox(hInst, MAKEINTRESOURCE(IDD_ADVANCED_SEARCH), hMain, AdvancedSearchDlgProc);
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
            ProgressHelper progress(hwnd, _T("Replacing..."), textLen);
            
            while (FindNextText(lpfr->Flags & ~FR_DOWN)) // Force search down
            {
                SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)szReplaceWith);
                nCount++;
                
                CHARRANGE cr;
                SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
                progress.Update(cr.cpMax);
            }
            TCHAR szMsg[64];
            StringCchPrintf(szMsg, 64, _T("Replaced %d occurrences."), nCount);
            MessageBox(hwnd, szMsg, _T("Replace All"), MB_OK);
        }
        return 0;
    }

    switch (msg)
    {
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
                if (CheckSaveChanges())
                {
                    CreateNewDocument(szFile);
                    AddRecentFile(szFile);
                }
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
            
        int statwidths[] = {150, 300, 450, 550, 700, -1};
        SendMessage(hStatus, SB_SETPARTS, 6, (LPARAM)statwidths);

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

        UpdateTitle();
        UpdateStatusBar();
        
        SetTimer(hwnd, 1, 2000, NULL);
    }
    break;
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
            }
        }
        else if (wParam == IDT_BACKGROUND_LOAD)
        {
            if (g_BgLoad.bActive)
            {
                KillTimer(hwnd, IDT_BACKGROUND_LOAD); // One shot - load the rest in one go

                // Load the rest of the file in one go
                StreamContext ctx = {0};
                ctx.hFile = g_BgLoad.hFile;
                ctx.bHex = g_BgLoad.bHex;
                ctx.dwOffset = g_BgLoad.dwOffset;
                ctx.pProgress = NULL; 
                ctx.hProgressBar = hProgressBarStatus; // Use status bar for progress
                ctx.dwTotalSize = g_BgLoad.dwTotalSize;
                ctx.dwReadSoFar = g_BgLoad.dwCurrentPos;
                ctx.hMap = g_BgLoad.hMap;
                ctx.pData = g_BgLoad.pData;
                ctx.dwMapPos = g_BgLoad.dwCurrentPos;
                ctx.inBuffer = g_BgLoad.inBuffer;
                
                EDITSTREAM es = {0};
                es.dwCookie = (DWORD_PTR)&ctx;
                es.pfnCallback = StreamInCallback;
                
                DWORD dwFormat = SF_TEXT;
                if (g_BgLoad.encoding == ENC_UTF16LE || g_BgLoad.encoding == ENC_UTF16BE)
                     dwFormat = SF_TEXT | SF_UNICODE;
                else
                {
                    const EncodingInfo* info = GetEncodingInfo(g_BgLoad.encoding);
                    dwFormat = SF_TEXT | (info->codePage << 16) | SF_USECODEPAGE;
                }
                
                // Save current state
                POINT ptScroll;
                SendMessage(g_Document.hEditor, EM_GETSCROLLPOS, 0, (LPARAM)&ptScroll);
                CHARRANGE cr;
                SendMessage(g_Document.hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
                
                // Disable updates to prevent flickering and scrolling
                SendMessage(g_Document.hEditor, WM_SETREDRAW, FALSE, 0);
                
                // Append to end
                SendMessage(g_Document.hEditor, EM_SETSEL, -1, -1);
                SendMessage(g_Document.hEditor, EM_STREAMIN, dwFormat | SFF_SELECTION, (LPARAM)&es);
                
                // Restore state
                SendMessage(g_Document.hEditor, EM_EXSETSEL, 0, (LPARAM)&cr);
                SendMessage(g_Document.hEditor, EM_SETSCROLLPOS, 0, (LPARAM)&ptScroll);
                
                SendMessage(g_Document.hEditor, WM_SETREDRAW, TRUE, 0);
                RedrawWindow(g_Document.hEditor, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW); 
                
                // Done
                g_BgLoad.bActive = FALSE;
                
                if (g_BgLoad.pData) UnmapViewOfFile(g_BgLoad.pData);
                if (g_BgLoad.hMap) CloseHandle(g_BgLoad.hMap);
                CloseHandle(g_BgLoad.hFile);
                
                ShowWindow(hProgressBarStatus, SW_HIDE);
                g_Document.bLoading = FALSE;
                
                // Restore Event Mask
                SendMessage(g_Document.hEditor, EM_SETEVENTMASK, 0, g_BgLoad.dwOriginalEventMask);
                
                // Restore ReadOnly based on file attribute
                DWORD dwAttrs = GetFileAttributes(g_Document.szFileName);
                BOOL bReadOnly = (dwAttrs != INVALID_FILE_ATTRIBUTES) && (dwAttrs & FILE_ATTRIBUTE_READONLY);
                SendMessage(g_Document.hEditor, EM_SETREADONLY, bReadOnly, 0);
                
                // Re-enable ReadOnly menu item
                HMENU hMenu = GetMenu(hwnd);
                EnableMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, MF_ENABLED);
                
                // Restore timer for file changes
                SetTimer(hwnd, 1, 2000, NULL);
            }
        }
        break;
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
            if (nPart >= 0 && nPart < 6)
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

        CheckMenuItem(hMenu, ID_OPTIONS_ADD_TO_SHELL, IsAppInShellMenu() ? MF_CHECKED : MF_UNCHECKED);
        // CheckMenuItem(hMenu, ID_OPTIONS_ASSOCIATE_TXT, IsTxtAssociated() ? MF_CHECKED : MF_UNCHECKED);
        
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
            if (!doc.bLoading && !doc.bIsDirty)
            {
                doc.bIsDirty = TRUE;
                UpdateTitle();
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

        case ID_OPTIONS_ADD_TO_SHELL:
            if (IsAppInShellMenu())
                RemoveAppFromShellMenu();
            else
                AddAppToShellMenu();
            break;
        case ID_OPTIONS_ASSOCIATE_TXT:
            AssociateTxtFiles();
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
        case ID_EDIT_ADVANCED_SEARCH: DoAdvancedSearch(); break;
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
        case ID_FILE_OPENFOLDER: DoOpenFolder(); break;
        case ID_FILE_OPENCMD: DoOpenCmd(); break;
        case ID_FILE_OPENDEFAULT: DoOpenDefault(); break;
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
