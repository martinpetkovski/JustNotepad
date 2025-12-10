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
#include <shlwapi.h>
#include <shellapi.h>
#include "resource.h"
#include "TextHelpers.h"

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")

#define IDC_EDITOR 1001
#define IDC_STATUSBAR 1002

HWND hMain;
HWND hEditor;
HWND hStatus;
HINSTANCE hInst;
HWND hFindReplaceDlg = NULL;
UINT uFindReplaceMsg = 0;
FINDREPLACE fr;
TCHAR szFindWhat[256];
TCHAR szReplaceWith[256];
TCHAR szFileName[MAX_PATH] = {0};
BOOL bWordWrap = FALSE;
BOOL bHexMode = FALSE;
std::vector<CHARRANGE> vecSelectionStack;

struct StreamContext {
    HANDLE hFile;
    BOOL bHex;
    std::vector<char> inBuffer;
    DWORD dwOffset;
    std::string outBuffer;
};

// Forward declarations
void UpdateStatusBar();
void UpdateTitle();

// Stream callback for reading file
DWORD CALLBACK StreamInCallback(DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG *pcb)
{
    StreamContext* ctx = (StreamContext*)dwCookie;
    *pcb = 0;

    if (!ctx->bHex)
    {
        if (ReadFile(ctx->hFile, pbBuff, cb, (LPDWORD)pcb, NULL))
        {
            return 0;
        }
        return 1; // Error
    }
    else
    {
        LONG bytesCopied = 0;
        while (bytesCopied < cb)
        {
            if (ctx->inBuffer.empty())
            {
                BYTE buf[16];
                DWORD bytesRead = 0;
                if (!ReadFile(ctx->hFile, buf, 16, &bytesRead, NULL) || bytesRead == 0)
                {
                    break; // EOF or Error
                }
                
                std::string line;
                FormatHexLine(ctx->dwOffset, buf, bytesRead, line);
                ctx->dwOffset += bytesRead;
                
                ctx->inBuffer.insert(ctx->inBuffer.end(), line.begin(), line.end());
            }
            
            LONG toCopy = min((LONG)ctx->inBuffer.size(), cb - bytesCopied);
            memcpy(pbBuff + bytesCopied, ctx->inBuffer.data(), toCopy);
            
            ctx->inBuffer.erase(ctx->inBuffer.begin(), ctx->inBuffer.begin() + toCopy);
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

void DoFileNew()
{
    SetWindowText(hEditor, _T(""));
    szFileName[0] = '\0';
    bHexMode = FALSE;
    
    SendMessage(hEditor, EM_SETREADONLY, FALSE, 0);
    HMENU hMenu = GetMenu(hMain);
    CheckMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, MF_UNCHECKED);

    UpdateTitle();
    UpdateStatusBar();
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
        HANDLE hFile = CreateFile(ofn.lpstrFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE)
        {
            // Check for binary
            BYTE buf[1024];
            DWORD read;
            BOOL isBinary = FALSE;
            if (ReadFile(hFile, buf, sizeof(buf), &read, NULL) && read > 0)
            {
                for (DWORD i = 0; i < read; i++)
                {
                    if (buf[i] == 0 && read > 2) // Simple check for null byte
                    {
                        // Check for UTF-16 BOM
                        if (i == 1 && buf[0] == 0xFF && buf[1] == 0xFE) continue;
                        if (i == 0 && buf[0] == 0xFF && buf[1] == 0xFE) continue;
                        isBinary = TRUE;
                        break;
                    }
                }
            }
            SetFilePointer(hFile, 0, NULL, FILE_BEGIN); // Reset
            
            bHexMode = isBinary;
            
            StreamContext ctx = {0};
            ctx.hFile = hFile;
            ctx.bHex = bHexMode;
            ctx.dwOffset = 0;
            
            EDITSTREAM es = {0};
            es.dwCookie = (DWORD_PTR)&ctx;
            es.pfnCallback = StreamInCallback;
            SendMessage(hEditor, EM_STREAMIN, SF_TEXT, (LPARAM)&es);
            CloseHandle(hFile);
            
            StringCchCopy(szFileName, MAX_PATH, szFile);
            
            // Check ReadOnly
            DWORD dwAttrs = GetFileAttributes(szFileName);
            BOOL bReadOnly = (dwAttrs != INVALID_FILE_ATTRIBUTES) && (dwAttrs & FILE_ATTRIBUTE_READONLY);
            SendMessage(hEditor, EM_SETREADONLY, bReadOnly, 0);
            HMENU hMenu = GetMenu(hMain);
            CheckMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, bReadOnly ? MF_CHECKED : MF_UNCHECKED);

            UpdateTitle();
            UpdateStatusBar();
        }
    }
}

BOOL DoFileSaveAs()
{
    OPENFILENAME ofn = {0};
    TCHAR szFile[MAX_PATH] = {0};

    if (szFileName[0])
        StringCchCopy(szFile, MAX_PATH, szFileName);

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
            StreamContext ctx = {0};
            ctx.hFile = hFile;
            ctx.bHex = bHexMode;
            
            EDITSTREAM es = {0};
            es.dwCookie = (DWORD_PTR)&ctx;
            es.pfnCallback = StreamOutCallback;
            SendMessage(hEditor, EM_STREAMOUT, SF_TEXT, (LPARAM)&es);
            CloseHandle(hFile);

            StringCchCopy(szFileName, MAX_PATH, szFile);
            UpdateTitle();
            return TRUE;
        }
    }
    return FALSE;
}

void DoFileSave()
{
    if (szFileName[0])
    {
        HANDLE hFile = CreateFile(szFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE)
        {
            StreamContext ctx = {0};
            ctx.hFile = hFile;
            ctx.bHex = bHexMode;
            
            EDITSTREAM es = {0};
            es.dwCookie = (DWORD_PTR)&ctx;
            es.pfnCallback = StreamOutCallback;
            SendMessage(hEditor, EM_STREAMOUT, SF_TEXT, (LPARAM)&es);
            CloseHandle(hFile);
        }
    }
    else
    {
        DoFileSaveAs();
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
        di.lpszDocName = _T("JustNotepad Document");

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
    fr.lpstrReplaceWith = szReplaceWith;
    fr.wFindWhatLen = sizeof(szFindWhat);
    fr.wReplaceWithLen = sizeof(szReplaceWith);
    fr.Flags = FR_DOWN;

    hFindReplaceDlg = ReplaceText(&fr);
}

BOOL FindNextText(DWORD dwFlags)
{
    FINDTEXTEX ft;
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);

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
        SendMessage(hEditor, EM_SETSEL, ft.chrgText.cpMin, ft.chrgText.cpMax);
        SendMessage(hEditor, EM_SCROLLCARET, 0, 0);
        return TRUE;
    }
    return FALSE;
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

void UpdateTitle()
{
    TCHAR szTitle[MAX_PATH + 32];
    if (szFileName[0])
        StringCchPrintf(szTitle, MAX_PATH + 32, _T("JustNotepad - %s"), szFileName);
    else
        StringCchCopy(szTitle, MAX_PATH + 32, _T("JustNotepad - Untitled"));
    SetWindowText(hMain, szTitle);
}

void UpdateStatusBar()
{
    if (!hStatus) return;

    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    long nLine = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nCol = cr.cpMin - SendMessage(hEditor, EM_LINEINDEX, nLine, 0);
    
    // Get text length
    GETTEXTLENGTHEX gtl = { GTL_DEFAULT, 1200 }; // 1200 is CP_UNICODE
    long nLen = SendMessage(hEditor, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 0);

    // Get Zoom
    int nNum, nDen;
    SendMessage(hEditor, EM_GETZOOM, (WPARAM)&nNum, (LPARAM)&nDen);
    int nZoom = (nDen == 0) ? 100 : (nNum * 100 / nDen);

    TCHAR szStatus[256];
    StringCchPrintf(szStatus, 256, _T("Ln %d, Col %d"), nLine + 1, nCol + 1);
    SendMessage(hStatus, SB_SETTEXT, 0, (LPARAM)szStatus);

    StringCchPrintf(szStatus, 256, _T("Chars: %d"), nLen);
    SendMessage(hStatus, SB_SETTEXT, 1, (LPARAM)szStatus);

    StringCchPrintf(szStatus, 256, _T("Zoom: %d%%"), nZoom);
    SendMessage(hStatus, SB_SETTEXT, 2, (LPARAM)szStatus);

    SendMessage(hStatus, SB_SETTEXT, 3, (LPARAM)_T("Windows (CRLF)")); // Simplified
    SendMessage(hStatus, SB_SETTEXT, 4, (LPARAM)_T("UTF-8")); // Simplified
}

CHARRANGE GetExpandedRange(CHARRANGE cr)
{
    CHARRANGE crNew = cr;
    long nLen = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);

    // 1. If nothing or just caret selected, select word
    if (cr.cpMin == cr.cpMax)
    {
        // Use EM_FINDWORDBREAK to find word boundaries
        long nStart = SendMessage(hEditor, EM_FINDWORDBREAK, WB_MOVEWORDLEFT, cr.cpMin);
        long nEnd = SendMessage(hEditor, EM_FINDWORDBREAK, WB_MOVEWORDRIGHT, cr.cpMin);
        
        // If we are at the end of a word, WB_MOVEWORDLEFT might go back to the start of the *current* word
        // But if we are in whitespace, it might behave differently.
        // Let's try a simpler approach: WB_ISDELIMITER check?
        // Actually, let's just trust EM_FINDWORDBREAK for now.
        
        if (nStart < nEnd)
        {
            crNew.cpMin = nStart;
            crNew.cpMax = nEnd;
            return crNew;
        }
    }

    // 2. If word selected (or similar), try to select enclosing brackets/quotes
    // This is a simplified check. We scan outwards.
    // Limit scan to 2000 chars to avoid performance hit on huge files
    long nScanLimit = 2000;
    long nStartScan = max(0, cr.cpMin - nScanLimit);
    long nEndScan = min(nLen, cr.cpMax + nScanLimit);
    
    // Get text range
    TEXTRANGE tr;
    tr.chrg.cpMin = nStartScan;
    tr.chrg.cpMax = nEndScan;
    std::vector<TCHAR> buffer(nEndScan - nStartScan + 1);
    tr.lpstrText = buffer.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    // Map selection to buffer indices
    long nBufMin = cr.cpMin - nStartScan;
    long nBufMax = cr.cpMax - nStartScan;
    
    Range r = FindEnclosingBrackets(buffer.data(), (int)buffer.size() - 1, nBufMin, nBufMax);
    if (r.start != -1)
    {
        crNew.cpMin = nStartScan + r.start;
        crNew.cpMax = nStartScan + r.end;
        return crNew;
    }

    // 3. Select Line
    long nLineStart = SendMessage(hEditor, EM_LINEFROMCHAR, cr.cpMin, 0);
    long nLineEnd = SendMessage(hEditor, EM_LINEFROMCHAR, cr.cpMax, 0);
    
    long nStartChar = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEndChar = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEndChar == -1) nEndChar = nLen; // Last line
    
    if (nStartChar < cr.cpMin || nEndChar > cr.cpMax)
    {
        crNew.cpMin = nStartChar;
        crNew.cpMax = nEndChar;
        return crNew;
    }

    // 4. Select Paragraph (delimited by empty lines)
    // Scan up
    long nCurrLine = nLineStart;
    while (nCurrLine > 0)
    {
        long nIdx = SendMessage(hEditor, EM_LINEINDEX, nCurrLine - 1, 0);
        long nLenLine = SendMessage(hEditor, EM_LINELENGTH, nIdx, 0);
        if (nLenLine == 0) break;
        nCurrLine--;
    }
    long nParaStart = SendMessage(hEditor, EM_LINEINDEX, nCurrLine, 0);
    
    // Scan down
    nCurrLine = nLineEnd;
    long nLineCount = SendMessage(hEditor, EM_GETLINECOUNT, 0, 0);
    while (nCurrLine < nLineCount - 1)
    {
        long nIdx = SendMessage(hEditor, EM_LINEINDEX, nCurrLine + 1, 0);
        long nLenLine = SendMessage(hEditor, EM_LINELENGTH, nIdx, 0);
        if (nLenLine == 0) break;
        nCurrLine++;
    }
    long nParaEnd = SendMessage(hEditor, EM_LINEINDEX, nCurrLine + 1, 0);
    if (nParaEnd == -1) nParaEnd = nLen;

    if (nParaStart < cr.cpMin || nParaEnd > cr.cpMax)
    {
        crNew.cpMin = nParaStart;
        crNew.cpMax = nParaEnd;
        return crNew;
    }

    // 5. Select All
    crNew.cpMin = 0;
    crNew.cpMax = -1;
    return crNew;
}

void DoExpandSelection()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);

    // If stack is empty or top is not current selection, reset stack
    if (vecSelectionStack.empty() || 
        vecSelectionStack.back().cpMin != cr.cpMin || 
        vecSelectionStack.back().cpMax != cr.cpMax)
    {
        vecSelectionStack.clear();
        vecSelectionStack.push_back(cr);
    }

    CHARRANGE crNew = GetExpandedRange(cr);
    
    // Only update if changed
    if (crNew.cpMin != cr.cpMin || crNew.cpMax != cr.cpMax)
    {
        vecSelectionStack.push_back(crNew);
        SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
        // Scroll into view if needed?
        // SendMessage(hEditor, EM_SCROLLCARET, 0, 0);
    }
}

void DoShrinkSelection()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);

    // If stack is empty or top is not current selection, we can't shrink reliably based on history
    // But if the user just expanded, the stack should be valid.
    if (vecSelectionStack.empty()) return;

    // Check if current matches top
    if (vecSelectionStack.back().cpMin == cr.cpMin && 
        vecSelectionStack.back().cpMax == cr.cpMax)
    {
        if (vecSelectionStack.size() > 1)
        {
            vecSelectionStack.pop_back();
            CHARRANGE crPrev = vecSelectionStack.back();
            SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crPrev);
        }
    }
    else
    {
        // User moved selection manually, clear stack
        vecSelectionStack.clear();
    }
}

void DoOpenFolder()
{
    if (szFileName[0])
    {
        TCHAR szDir[MAX_PATH];
        StringCchCopy(szDir, MAX_PATH, szFileName);
        PathRemoveFileSpec(szDir);
        ShellExecute(NULL, _T("explore"), szDir, NULL, NULL, SW_SHOW);
    }
}

void DoOpenCmd()
{
    if (szFileName[0])
    {
        TCHAR szDir[MAX_PATH];
        StringCchCopy(szDir, MAX_PATH, szFileName);
        PathRemoveFileSpec(szDir);
        ShellExecute(NULL, NULL, _T("cmd.exe"), NULL, szDir, SW_SHOW);
    }
}

void DoCopyLine()
{
    CHARRANGE crOld;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&crOld);
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, crOld.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, crOld.cpMax);
    
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    CHARRANGE crLine = {nStart, nEnd};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crLine);
    SendMessage(hEditor, WM_COPY, 0, 0);
    
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crOld);
}

void DoDeleteLine()
{
    CHARRANGE crOld;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&crOld);
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, crOld.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, crOld.cpMax);
    
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    CHARRANGE crLine = {nStart, nEnd};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crLine);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)_T(""));
}

void DoDuplicateLine()
{
    CHARRANGE crOld;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&crOld);
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, crOld.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, crOld.cpMax);
    
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    // Get text
    long nLen = nEnd - nStart;
    std::vector<TCHAR> buffer(nLen + 1);
    TEXTRANGE tr;
    tr.chrg.cpMin = nStart;
    tr.chrg.cpMax = nEnd;
    tr.lpstrText = buffer.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    // Insert
    SendMessage(hEditor, EM_SETSEL, nEnd, nEnd);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)buffer.data());
    
    // Restore selection (maybe move it down?)
    // VS Code moves selection down to the duplicated line
    long nDiff = nEnd - nStart;
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
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMax);
    
    // If selection ends at the start of a line, don't include that line
    if (cr.cpMax > cr.cpMin && cr.cpMax == SendMessage(hEditor, EM_LINEINDEX, nLineEnd, 0))
    {
        nLineEnd--;
    }

    if (nLineStart == 0) return; // Cannot move up

    // Range of lines to move (Block)
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    // Range of previous line (Prev)
    long nPrevStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart - 1, 0);
    long nPrevEnd = nStart;
    
    // Get text of Block
    long nLen = nEnd - nStart;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg.cpMin = nStart;
    tr.chrg.cpMax = nEnd;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    std::basic_string<TCHAR> sBlock = buf.data();
    
    // Get text of Prev
    long nPrevLen = nPrevEnd - nPrevStart;
    std::vector<TCHAR> prevBuf(nPrevLen + 1);
    tr.chrg.cpMin = nPrevStart;
    tr.chrg.cpMax = nPrevEnd;
    tr.lpstrText = prevBuf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    std::basic_string<TCHAR> sPrev = prevBuf.data();
    
    // Ensure newlines
    if (!sBlock.empty() && sBlock.back() != '\n') sBlock += _T("\r\n");
    if (!sPrev.empty() && sPrev.back() != '\n') sPrev += _T("\r\n");
    
    // Construct new text: Block + Prev
    std::basic_string<TCHAR> sNew = sBlock + sPrev;
    
    // Replace everything from nPrevStart to nEnd
    SendMessage(hEditor, EM_SETSEL, nPrevStart, nEnd);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)sNew.c_str());
    
    // Restore selection
    long nOffsetStart = cr.cpMin - nStart;
    long nOffsetEnd = cr.cpMax - nStart;
    
    CHARRANGE crNew;
    crNew.cpMin = nPrevStart + nOffsetStart;
    crNew.cpMax = nPrevStart + nOffsetEnd;
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoMoveLineDown()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMax);
    
    if (cr.cpMax > cr.cpMin && cr.cpMax == SendMessage(hEditor, EM_LINEINDEX, nLineEnd, 0))
    {
        nLineEnd--;
    }
    
    long nCount = SendMessage(hEditor, EM_GETLINECOUNT, 0, 0);
    if (nLineEnd >= nCount - 1) return; // Already at bottom
    
    // Range of lines to move (Block)
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    // Range of next line (Next)
    long nNextStart = nEnd;
    long nNextEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 2, 0);
    if (nNextEnd == -1) nNextEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    // Get text of Block
    long nLen = nEnd - nStart;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg.cpMin = nStart;
    tr.chrg.cpMax = nEnd;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    std::basic_string<TCHAR> sBlock = buf.data();
    
    // Get text of Next
    long nNextLen = nNextEnd - nNextStart;
    std::vector<TCHAR> nextBuf(nNextLen + 1);
    tr.chrg.cpMin = nNextStart;
    tr.chrg.cpMax = nNextEnd;
    tr.lpstrText = nextBuf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    std::basic_string<TCHAR> sNext = nextBuf.data();
    
    // Ensure newlines
    if (!sBlock.empty() && sBlock.back() != '\n') sBlock += _T("\r\n");
    if (!sNext.empty() && sNext.back() != '\n') sNext += _T("\r\n");
    
    // Construct new text: Next + Block
    std::basic_string<TCHAR> sNew = sNext + sBlock;
    
    // Replace everything from nStart to nNextEnd
    SendMessage(hEditor, EM_SETSEL, nStart, nNextEnd);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)sNew.c_str());
    
    // Restore selection
    long nNewBlockStart = nStart + sNext.length();
    
    long nOffsetStart = cr.cpMin - nStart;
    long nOffsetEnd = cr.cpMax - nStart;
    
    CHARRANGE crNew;
    crNew.cpMin = nNewBlockStart + nOffsetStart;
    crNew.cpMax = nNewBlockStart + nOffsetEnd;
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoOpenDefault()
{
    if (szFileName[0])
    {
        ShellExecute(NULL, _T("open"), szFileName, NULL, NULL, SW_SHOW);
    }
}

void DoSelectLine()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMax);
    
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    SendMessage(hEditor, EM_SETSEL, nStart, nEnd);
}

void DoUpperCase()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    if (cr.cpMin == cr.cpMax) return;
    
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
    if (cr.cpMin == cr.cpMax) return;
    
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
    if (cr.cpMin == cr.cpMax) return;
    
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
    if (cr.cpMin == cr.cpMax) return;
    
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
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMax);
    
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    long nLen = nEnd - nStart;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg.cpMin = nStart;
    tr.chrg.cpMax = nEnd;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::basic_string<TCHAR> out = SortLines(buf.data());
    
    SendMessage(hEditor, EM_SETSEL, nStart, nEnd);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    
    // Restore selection
    CHARRANGE crNew = {nStart, nStart + (long)out.length()};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoJoinLines()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMax);
    
    if (cr.cpMin == cr.cpMax)
    {
        long nCount = SendMessage(hEditor, EM_GETLINECOUNT, 0, 0);
        if (nLineStart >= nCount - 1) return;
        nLineEnd = nLineStart + 1;
    }
    else if (cr.cpMax > cr.cpMin && cr.cpMax == SendMessage(hEditor, EM_LINEINDEX, nLineEnd, 0))
    {
        nLineEnd--;
    }
    
    if (nLineStart == nLineEnd) return;

    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    long nLen = nEnd - nStart;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg.cpMin = nStart;
    tr.chrg.cpMax = nEnd;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::basic_string<TCHAR> out = JoinLines(buf.data());
    
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
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
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMax);
    
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    long nLen = nEnd - nStart;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg.cpMin = nStart;
    tr.chrg.cpMax = nEnd;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::basic_string<TCHAR> out = IndentLines(buf.data());

    SendMessage(hEditor, EM_SETSEL, nStart, nEnd);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    
    CHARRANGE crNew = {nStart, nStart + (long)out.length()};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoUnindent()
{
    CHARRANGE cr;
    SendMessage(hEditor, EM_EXGETSEL, 0, (LPARAM)&cr);
    
    long nLineStart = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMin);
    long nLineEnd = SendMessage(hEditor, EM_EXLINEFROMCHAR, 0, cr.cpMax);
    
    long nStart = SendMessage(hEditor, EM_LINEINDEX, nLineStart, 0);
    long nEnd = SendMessage(hEditor, EM_LINEINDEX, nLineEnd + 1, 0);
    if (nEnd == -1) nEnd = SendMessage(hEditor, WM_GETTEXTLENGTH, 0, 0);
    
    long nLen = nEnd - nStart;
    std::vector<TCHAR> buf(nLen + 1);
    TEXTRANGE tr;
    tr.chrg.cpMin = nStart;
    tr.chrg.cpMax = nEnd;
    tr.lpstrText = buf.data();
    SendMessage(hEditor, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    
    std::basic_string<TCHAR> out = UnindentLines(buf.data());

    SendMessage(hEditor, EM_SETSEL, nStart, nEnd);
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)out.c_str());
    
    CHARRANGE crNew = {nStart, nStart + (long)out.length()};
    SendMessage(hEditor, EM_EXSETSEL, 0, (LPARAM)&crNew);
}

void DoToggleReadOnly()
{
    if (!szFileName[0]) return;
    
    DWORD dwAttrs = GetFileAttributes(szFileName);
    if (dwAttrs == INVALID_FILE_ATTRIBUTES) return;
    
    if (dwAttrs & FILE_ATTRIBUTE_READONLY)
    {
        SetFileAttributes(szFileName, dwAttrs & ~FILE_ATTRIBUTE_READONLY);
        SendMessage(hEditor, EM_SETREADONLY, FALSE, 0);
    }
    else
    {
        SetFileAttributes(szFileName, dwAttrs | FILE_ATTRIBUTE_READONLY);
        SendMessage(hEditor, EM_SETREADONLY, TRUE, 0);
    }
    
    // Update menu check
    HMENU hMenu = GetMenu(hMain);
    CheckMenuItem(hMenu, ID_FILE_TOGGLEREADONLY, (dwAttrs & FILE_ATTRIBUTE_READONLY) ? MF_UNCHECKED : MF_CHECKED); // Inverted logic because we just toggled
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
            while (FindNextText(lpfr->Flags & ~FR_DOWN)) // Force search down
            {
                SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)szReplaceWith);
                nCount++;
            }
            TCHAR szMsg[64];
            StringCchPrintf(szMsg, 64, _T("Replaced %d occurrences."), nCount);
            MessageBox(hwnd, szMsg, _T("Replace All"), MB_OK);
        }
        return 0;
    }

    switch (msg)
    {
    case WM_CREATE:
    {
        hMain = hwnd;
        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        
        // Create Status Bar
        hStatus = CreateWindowEx(0, STATUSCLASSNAME, NULL,
            WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0,
            hwnd, (HMENU)IDC_STATUSBAR, hInst, NULL);
            
        int statwidths[] = {150, 300, 400, 550, -1};
        SendMessage(hStatus, SB_SETPARTS, 5, (LPARAM)statwidths);

        // Calculate Editor Height
        RECT rcStatus;
        GetWindowRect(hStatus, &rcStatus);
        int nStatusHeight = rcStatus.bottom - rcStatus.top;

        hEditor = CreateWindowEx(0, MSFTEDIT_CLASS, _T(""),
                                 WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_NOHIDESEL,
                                 0, 0, rcClient.right, rcClient.bottom - nStatusHeight,
                                 hwnd, (HMENU)IDC_EDITOR, hInst, NULL);
        
        SendMessage(hEditor, EM_EXLIMITTEXT, 0, -1);
        SendMessage(hEditor, EM_SETEVENTMASK, 0, ENM_SELCHANGE | ENM_CHANGE);
        
        HFONT hFont = CreateFont(18, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, ANSI_CHARSET, 
                                 OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, 
                                 DEFAULT_PITCH | FF_SWISS, _T("Consolas"));
        SendMessage(hEditor, WM_SETFONT, (WPARAM)hFont, TRUE);

        // Register Find/Replace Message
        uFindReplaceMsg = RegisterWindowMessage(FINDMSGSTRING);
        
        UpdateTitle();
        UpdateStatusBar();
    }
    break;
    case WM_SIZE:
    {
        RECT rcClient;
        GetClientRect(hwnd, &rcClient);
        
        SendMessage(hStatus, WM_SIZE, 0, 0);
        
        RECT rcStatus;
        GetWindowRect(hStatus, &rcStatus);
        int nStatusHeight = rcStatus.bottom - rcStatus.top;

        if (hEditor)
        {
            MoveWindow(hEditor, 0, 0, rcClient.right, rcClient.bottom - nStatusHeight, TRUE);
        }
    }
    break;
    case WM_NOTIFY:
    {
        NMHDR *pnm = (NMHDR *)lParam;
        if (pnm->hwndFrom == hEditor && pnm->code == EN_SELCHANGE)
        {
            UpdateStatusBar();
        }
    }
    break;
    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
        case ID_FILE_NEW: DoFileNew(); break;
        case ID_FILE_OPEN: DoFileOpen(); break;
        case ID_FILE_SAVE: DoFileSave(); break;
        case ID_FILE_SAVEAS: DoFileSaveAs(); break;
        case ID_FILE_PAGESETUP: DoPageSetup(); break;
        case ID_FILE_PRINT: DoPrint(); break;
        case ID_FILE_EXIT: PostQuitMessage(0); break;

        case ID_EDIT_UNDO: SendMessage(hEditor, EM_UNDO, 0, 0); break;
        case ID_EDIT_REDO: SendMessage(hEditor, EM_REDO, 0, 0); break;
        case ID_EDIT_CUT: SendMessage(hEditor, WM_CUT, 0, 0); break;
        case ID_EDIT_COPY: SendMessage(hEditor, WM_COPY, 0, 0); break;
        case ID_EDIT_PASTE: SendMessage(hEditor, WM_PASTE, 0, 0); break;
        case ID_EDIT_DELETE: SendMessage(hEditor, WM_CLEAR, 0, 0); break;
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
        case ID_FILE_OPENFOLDER: DoOpenFolder(); break;
        case ID_FILE_OPENCMD: DoOpenCmd(); break;
        case ID_FILE_OPENDEFAULT: DoOpenDefault(); break;
        case ID_FILE_TOGGLEREADONLY: DoToggleReadOnly(); break;
        case ID_EDIT_SELECTLINE: DoSelectLine(); break;
        case ID_EDIT_UPPERCASE: DoUpperCase(); break;
        case ID_EDIT_LOWERCASE: DoLowerCase(); break;
        case ID_EDIT_CAPITALIZE: DoCapitalize(); break;
        case ID_EDIT_SENTENCECASE: DoSentenceCase(); break;
        case ID_EDIT_SORTLINES: DoSortLines(); break;
        case ID_EDIT_JOINLINES: DoJoinLines(); break;
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
            if (IsWindowVisible(hStatus)) ShowWindow(hStatus, SW_HIDE);
            else ShowWindow(hStatus, SW_SHOW);
            SendMessage(hwnd, WM_SIZE, 0, 0); // Recalc layout
            break;
        }
    }
    break;
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
    hInst = hInstance;
    LoadLibrary(_T("Msftedit.dll"));
    
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_BAR_CLASSES;
    InitCommonControlsEx(&icex);

    WNDCLASSEX wc = {0};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = _T("JustNotepadClass");
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);

    if (!RegisterClassEx(&wc))
        return 1;

    HWND hwnd = CreateWindowEx(0, wc.lpszClassName, _T("JustNotepad"),
                               WS_OVERLAPPEDWINDOW,
                               CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
                               NULL, NULL, hInstance, NULL);

    if (!hwnd)
        return 1;

    // Load Menu and Accelerators from Resource
    HMENU hMenu = LoadMenu(hInstance, MAKEINTRESOURCE(IDR_MAINMENU));
    SetMenu(hwnd, hMenu);
    
    HACCEL hAccel = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR));

    ShowWindow(hwnd, SW_SHOWMAXIMIZED);
    UpdateWindow(hwnd);

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
