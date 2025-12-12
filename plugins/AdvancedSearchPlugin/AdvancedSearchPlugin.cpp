#include "../../src/PluginInterface.h"
#include "resource.h"
#include <windows.h>
#include <commctrl.h>
#include <richedit.h>
#include <tchar.h>
#include <strsafe.h>
#include <vector>
#include <string>
#include <algorithm>
#include <regex>
#include <thread>
#include <mutex>
#include <atomic>
#include <string_view>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "comctl32.lib")

HINSTANCE g_hInst = NULL;
HWND g_hMain = NULL;
HWND g_hEditor = NULL;

// Settings
BOOL g_bRegex = FALSE;
BOOL g_bFuzzy = FALSE;
std::wstring g_SettingsPath;

extern "C" {
    PLUGIN_API void Initialize(const wchar_t* settingsPath) {
        g_SettingsPath = settingsPath;
        g_bRegex = GetPrivateProfileInt(L"Settings", L"Regex", 0, g_SettingsPath.c_str());
        g_bFuzzy = GetPrivateProfileInt(L"Settings", L"Fuzzy", 0, g_SettingsPath.c_str());
    }

    PLUGIN_API void Shutdown() {
        if (!g_SettingsPath.empty()) {
            WritePrivateProfileString(L"Settings", L"Regex", g_bRegex ? L"1" : L"0", g_SettingsPath.c_str());
            WritePrivateProfileString(L"Settings", L"Fuzzy", g_bFuzzy ? L"1" : L"0", g_SettingsPath.c_str());
        }
    }
}

#define WM_SEARCH_UPDATE (WM_USER + 100)
#define WM_SEARCH_COMPLETE (WM_USER + 101)

struct SearchMatch {
    int lineNum; // 0-based
    int matchStart; // Index in line
    int matchLen;
    int score;
    long absolutePos; // Absolute character position
    std::basic_string<TCHAR> lineText; // Cache the line text
};

struct SearchState {
    std::atomic<bool> bActive;
    std::vector<TCHAR> buffer;
    std::basic_string<TCHAR> query;
    BOOL bRegex;
    BOOL bFuzzy;
    
    // Optimization buffers
    std::basic_string<TCHAR> queryLower;

    // Results
    std::vector<SearchMatch>* pMatches;
    HWND hDlg;
    int itemsPerPage;
    int* pCurrentPage;
    
    // UI Update Throttling
    DWORD lastUpdateTime;
    BOOL bFirstPageFilled;
    
    std::thread workerThread;
};

SearchState g_SearchState = {0};
std::mutex g_SearchMutex;
HFONT hDialogListFont = NULL;

std::basic_string<TCHAR> ToLowerCase(const std::basic_string<TCHAR>& input) {
    std::basic_string<TCHAR> output = input;
    std::transform(output.begin(), output.end(), output.begin(), ::towlower);
    return output;
}

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

void UpdateSearchList(HWND hDlg, const std::vector<SearchMatch>& matches, int page, int itemsPerPage, BOOL bAppendIfPossible = FALSE);

void SearchWorker()
{
    const TCHAR* pBuffer = g_SearchState.buffer.data();
    size_t maxLen = g_SearchState.buffer.size();
    if (maxLen > 0 && pBuffer[maxLen-1] == 0) maxLen--; // Exclude null terminator if present

    size_t bufferPos = 0;
    int lineNum = 0;
    long correction = 0;
    
    std::vector<TCHAR> lowerBuffer;
    std::vector<SearchMatch> pendingMatches;
    
    #ifdef UNICODE
    std::wstring wQuery = g_SearchState.query;
    #else
    std::string sQuery = g_SearchState.query;
    #endif

    DWORD lastReportTime = GetTickCount();

    while(g_SearchState.bActive && bufferPos < maxLen)
    {
        size_t lineStart = bufferPos;
        size_t lineEnd = lineStart;
        
        // Scan for newline (handle \r, \n, \r\n)
        while(lineEnd < maxLen) {
            if (pBuffer[lineEnd] == _T('\n')) break;
            if (pBuffer[lineEnd] == _T('\r')) {
                if (lineEnd + 1 < maxLen && pBuffer[lineEnd+1] == _T('\n')) {
                    // Found \r\n, treat as end of line
                }
                break;
            }
            lineEnd++;
        }
        
        size_t contentLen = lineEnd - lineStart;
        
        std::basic_string_view<TCHAR> lineView(pBuffer + lineStart, contentLen);
        long internalOffset = (long)lineStart; 
        
        if (g_SearchState.bRegex)
        {
            try {
                #ifdef UNICODE
                std::wregex re(wQuery, std::regex_constants::icase);
                std::wcmatch m;
                const wchar_t* searchStart = lineView.data();
                const wchar_t* searchEnd = lineView.data() + lineView.length();
                auto current = searchStart;
                while (std::regex_search(current, searchEnd, m, re))
                #else
                std::regex re(sQuery, std::regex_constants::icase);
                std::cmatch m;
                const char* searchStart = lineView.data();
                const char* searchEnd = lineView.data() + lineView.length();
                auto current = searchStart;
                while (std::regex_search(current, searchEnd, m, re))
                #endif
                {
                    SearchMatch sm;
                    sm.lineNum = lineNum;
                    sm.matchStart = (int)(m.position() + (current - searchStart));
                    sm.matchLen = (int)m.length();
                    sm.score = 0;
                    sm.absolutePos = (internalOffset + sm.matchStart) - correction;
                    sm.lineText = std::basic_string<TCHAR>(lineView);
                    pendingMatches.push_back(sm);
                    
                    current += m.position() + m.length();
                    if (m.length() == 0) current++;
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
            lineStr.reserve(lineView.length());
            for(auto c : lineView) lineStr += (char)c;
            for(auto c : wQuery) queryStr += (char)c;
            #else
            lineStr = std::string(lineView);
            queryStr = sQuery;
            #endif
            
            if (FuzzyMatch(lineStr, queryStr, score))
            {
                SearchMatch m;
                m.lineNum = lineNum;
                m.matchStart = -1; 
                m.matchLen = 0;
                m.score = score;
                m.absolutePos = internalOffset - correction; 
                m.lineText = std::basic_string<TCHAR>(lineView);
                pendingMatches.push_back(m);
            }
        }
        else
        {
            // Normal search
            size_t len = lineView.length();
            if (lowerBuffer.size() < len + 1) {
                lowerBuffer.resize(len + 256);
            }
            
            TCHAR* pLower = lowerBuffer.data();
            if (len > 0) {
                memcpy(pLower, lineView.data(), len * sizeof(TCHAR));
            }
            pLower[len] = 0;
            
            CharLowerBuff(pLower, (DWORD)len);
            
            std::basic_string_view<TCHAR> textLower(pLower, len);
            const std::basic_string<TCHAR>& queryLower = g_SearchState.queryLower;
            
            size_t pos = 0;
            while ((pos = textLower.find(queryLower, pos)) != std::basic_string_view<TCHAR>::npos)
            {
                SearchMatch sm;
                sm.lineNum = lineNum;
                sm.matchStart = (int)pos;
                sm.matchLen = (int)queryLower.length();
                sm.score = 0;
                sm.absolutePos = (internalOffset + sm.matchStart) - correction;
                sm.lineText = std::basic_string<TCHAR>(lineView);
                pendingMatches.push_back(sm);
                pos += queryLower.length();
            }
        }
        
        // Advance bufferPos and update correction
        if (lineEnd < maxLen) {
            if (pBuffer[lineEnd] == _T('\r')) {
                if (lineEnd + 1 < maxLen && pBuffer[lineEnd+1] == _T('\n')) {
                    bufferPos = lineEnd + 2;
                    correction++; // Buffer has \r\n (2 chars), Internal has \r (1 char)
                } else {
                    bufferPos = lineEnd + 1;
                }
            } else if (pBuffer[lineEnd] == _T('\n')) {
                bufferPos = lineEnd + 1;
            } else {
                bufferPos = lineEnd + 1;
            }
        } else {
            bufferPos = maxLen;
        }
        
        lineNum++;
        
        // Batch update
        if (pendingMatches.size() >= 100 || (GetTickCount() - lastReportTime > 100 && !pendingMatches.empty()))
        {
            {
                std::lock_guard<std::mutex> lock(g_SearchMutex);
                g_SearchState.pMatches->insert(g_SearchState.pMatches->end(), pendingMatches.begin(), pendingMatches.end());
            }
            pendingMatches.clear();
            PostMessage(g_SearchState.hDlg, WM_SEARCH_UPDATE, (WPARAM)bufferPos, 0);
            lastReportTime = GetTickCount();
        }
        else if (GetTickCount() - lastReportTime > 100)
        {
             PostMessage(g_SearchState.hDlg, WM_SEARCH_UPDATE, (WPARAM)bufferPos, 0);
             lastReportTime = GetTickCount();
        }
    }
    
    if (!pendingMatches.empty())
    {
        std::lock_guard<std::mutex> lock(g_SearchMutex);
        g_SearchState.pMatches->insert(g_SearchState.pMatches->end(), pendingMatches.begin(), pendingMatches.end());
    }
    
    PostMessage(g_SearchState.hDlg, WM_SEARCH_COMPLETE, 0, 0);
}

void StartSearch(HWND hDlg, std::vector<SearchMatch>& matches, int* pCurrentPage)
{
    TCHAR szQuery[256];
    GetDlgItemText(hDlg, IDC_SEARCH_TEXT, szQuery, 256);
    
    if (g_SearchState.bActive) {
        g_SearchState.bActive = FALSE;
        if (g_SearchState.workerThread.joinable()) {
            g_SearchState.workerThread.join();
        }
    }
    
    g_SearchState.bRegex = IsDlgButtonChecked(hDlg, IDC_SEARCH_REGEX);
    g_SearchState.bFuzzy = IsDlgButtonChecked(hDlg, IDC_SEARCH_FUZZY);
    g_SearchState.query = szQuery;
    if (!g_SearchState.bRegex && !g_SearchState.bFuzzy) {
        g_SearchState.queryLower = ToLowerCase(g_SearchState.query);
    }
    g_SearchState.pMatches = &matches;
    g_SearchState.hDlg = hDlg;
    g_SearchState.itemsPerPage = 100; 
    g_SearchState.pCurrentPage = pCurrentPage;
    g_SearchState.lastUpdateTime = GetTickCount();
    g_SearchState.bFirstPageFilled = FALSE;
    
    int nLen = SendMessage(g_hEditor, WM_GETTEXTLENGTH, 0, 0);
    g_SearchState.buffer.resize(nLen + 1);
    SendMessage(g_hEditor, WM_GETTEXT, nLen + 1, (LPARAM)g_SearchState.buffer.data());
    
    matches.clear();
    *pCurrentPage = 1;

    g_SearchState.bActive = TRUE;
    
    SendDlgItemMessage(hDlg, IDC_SEARCH_PROGRESS, PBM_SETRANGE32, 0, nLen);
    SendDlgItemMessage(hDlg, IDC_SEARCH_PROGRESS, PBM_SETPOS, 0, 0);
    ShowWindow(GetDlgItem(hDlg, IDC_SEARCH_PROGRESS), SW_SHOW);
    
    g_SearchState.workerThread = std::thread(SearchWorker);
}

void UpdateSearchList(HWND hDlg, const std::vector<SearchMatch>& matches, int page, int itemsPerPage, BOOL bAppendIfPossible)
{
    int startIdx = (page - 1) * itemsPerPage;
    int endIdx = min((int)matches.size(), startIdx + itemsPerPage);
    
    int currentCount = (int)SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETCOUNT, 0, 0);
    
    BOOL bNoMatchesDisplayed = FALSE;
    if (currentCount == 1) {
        int len = (int)SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETTEXTLEN, 0, 0);
        if (len > 0) {
            std::vector<TCHAR> buf(len + 1);
            SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETTEXT, 0, (LPARAM)buf.data());
            if (_tcscmp(buf.data(), _T("No matches found.")) == 0) bNoMatchesDisplayed = TRUE;
        }
    }

    if (matches.empty())
    {
        if (!bNoMatchesDisplayed) {
            SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_RESETCONTENT, 0, 0);
            SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_ADDSTRING, 0, (LPARAM)_T("No matches found."));
            SetDlgItemText(hDlg, IDC_PAGE_INFO, _T("0"));
            SetDlgItemText(hDlg, IDC_PAGE_TOTAL, _T("/ 0"));
            EnableWindow(GetDlgItem(hDlg, IDC_PREV_PAGE), FALSE);
            EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), FALSE);
        }
        return;
    }

    if (bNoMatchesDisplayed) {
        SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_RESETCONTENT, 0, 0);
        currentCount = 0;
    }

    BOOL bDoAppend = FALSE;
    if (bAppendIfPossible && currentCount >= 0 && currentCount < (endIdx - startIdx))
    {
        if (currentCount > 0) {
            int firstMatchIdx = (int)SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETITEMDATA, 0, 0);
            if (firstMatchIdx == startIdx) {
                bDoAppend = TRUE;
            }
        } else {
            bDoAppend = TRUE;
        }
    }

    if (bDoAppend)
    {
        SendMessage(GetDlgItem(hDlg, IDC_SEARCH_LIST), WM_SETREDRAW, FALSE, 0);
        for (int i = startIdx + currentCount; i < endIdx; i++)
        {
            const auto& m = matches[i];
            int line = m.lineNum;
            int col = m.matchStart + 1;
            TCHAR buf[512];
            StringCchPrintf(buf, 512, _T("Ln %d, Col %d: %s"), line + 1, col, m.lineText.c_str());
            int idx = SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_ADDSTRING, 0, (LPARAM)buf);
            SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_SETITEMDATA, idx, (LPARAM)i);
        }
        SendMessage(GetDlgItem(hDlg, IDC_SEARCH_LIST), WM_SETREDRAW, TRUE, 0);
    }
    else
    {
        SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_RESETCONTENT, 0, 0);
        SendMessage(GetDlgItem(hDlg, IDC_SEARCH_LIST), WM_SETREDRAW, FALSE, 0);
        for (int i = startIdx; i < endIdx; i++)
        {
            const auto& m = matches[i];
            int line = m.lineNum;
            int col = m.matchStart + 1;
            TCHAR buf[512];
            StringCchPrintf(buf, 512, _T("Ln %d, Col %d: %s"), line + 1, col, m.lineText.c_str());
            int idx = SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_ADDSTRING, 0, (LPARAM)buf);
            SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_SETITEMDATA, idx, (LPARAM)i);
        }
        SendMessage(GetDlgItem(hDlg, IDC_SEARCH_LIST), WM_SETREDRAW, TRUE, 0);
    }
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
        SetDlgItemText(hDlg, IDC_PAGE_INFO, _T("1"));
        SetDlgItemText(hDlg, IDC_PAGE_TOTAL, _T("/ 1"));
        
        RECT rcScreen;
        GetWindowRect(GetDesktopWindow(), &rcScreen);
        int screenW = rcScreen.right - rcScreen.left;
        int screenH = rcScreen.bottom - rcScreen.top;
        int width = screenW * 50 / 100;
        int height = screenH * 50 / 100;
        int x = rcScreen.left + (screenW - width) / 2;
        int y = rcScreen.top + (screenH - height) / 2;
        
        SetWindowPos(hDlg, NULL, x, y, width, height, SWP_NOZORDER);
        
        CheckDlgButton(hDlg, IDC_SEARCH_REGEX, g_bRegex ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(hDlg, IDC_SEARCH_FUZZY, g_bFuzzy ? BST_CHECKED : BST_UNCHECKED);

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
        
        SetWindowPos(GetDlgItem(hDlg, IDC_STATIC_FIND_LABEL), NULL, margin, row1Y + 4, labelWidth, labelHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_TEXT), NULL, editX, row1Y, editW, editHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_BTN), NULL, btnX, row1Y, btnWidth, btnHeight, SWP_NOZORDER);
        
        SetWindowPos(GetDlgItem(hDlg, IDC_STATIC_REPLACE_LABEL), NULL, margin, row2Y + 4, labelWidth, labelHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_REPLACE_TEXT), NULL, editX, row2Y, editW, editHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_REPLACE_ALL_BTN), NULL, btnX, row2Y, btnWidth, btnHeight, SWP_NOZORDER);
        
        SetWindowPos(GetDlgItem(hDlg, IDC_STATIC_OPTIONS_GROUP), NULL, margin, groupY, width - margin*2, groupH, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_REGEX), NULL, margin + 10, groupY + 20, 120, 20, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_FUZZY), NULL, margin + 140, groupY + 20, 100, 20, SWP_NOZORDER);
        
        SetWindowPos(GetDlgItem(hDlg, IDC_STATIC_RESULTS_LABEL), NULL, margin, resultsY, 50, labelHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_PROGRESS), NULL, margin + 55, resultsY, width - margin*2 - 55, 15, SWP_NOZORDER);
        
        SetWindowPos(GetDlgItem(hDlg, IDC_SEARCH_LIST), NULL, margin, listY, width - margin*2, bottomY - listY - 10, SWP_NOZORDER);
        
        SetWindowPos(GetDlgItem(hDlg, IDC_PREV_PAGE), NULL, margin, bottomY, btnWidth, btnHeight, SWP_NOZORDER);
        
        int pageInfoX = margin + btnWidth + 10;
        int pageInfoW = 40;
        int pageTotalX = pageInfoX + pageInfoW + 5;
        int pageTotalW = 55;
        
        SetWindowPos(GetDlgItem(hDlg, IDC_PAGE_INFO), NULL, pageInfoX, bottomY + 2, pageInfoW, 20, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_PAGE_TOTAL), NULL, pageTotalX, bottomY + 4, pageTotalW, labelHeight, SWP_NOZORDER);
        
        SetWindowPos(GetDlgItem(hDlg, IDC_NEXT_PAGE), NULL, margin + btnWidth + 120, bottomY, btnWidth, btnHeight, SWP_NOZORDER);
        
        SetWindowPos(GetDlgItem(hDlg, IDCANCEL), NULL, width - margin - btnWidth, bottomY, btnWidth, btnHeight, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hDlg, IDC_REPLACE_BTN), NULL, width - margin - btnWidth - 10 - 110, bottomY, 110, btnHeight, SWP_NOZORDER);
        
        return (INT_PTR)TRUE;
    }
    case WM_SEARCH_UPDATE:
    {
        size_t bufferPos = (size_t)wParam;
        SendDlgItemMessage(hDlg, IDC_SEARCH_PROGRESS, PBM_SETPOS, bufferPos, 0);
        
        if (currentPage == 1 && !g_SearchState.bFirstPageFilled)
        {
            std::lock_guard<std::mutex> lock(g_SearchMutex);
            if (matches.size() >= (size_t)itemsPerPage)
            {
                UpdateSearchList(hDlg, matches, currentPage, itemsPerPage, TRUE);
                g_SearchState.bFirstPageFilled = TRUE;
            }
            else if (GetTickCount() - g_SearchState.lastUpdateTime > 200)
            {
                UpdateSearchList(hDlg, matches, currentPage, itemsPerPage, TRUE);
                g_SearchState.lastUpdateTime = GetTickCount();
            }
        }
        
        static DWORD lastPageInfoUpdate = 0;
        if (GetTickCount() - lastPageInfoUpdate > 200) {
            std::lock_guard<std::mutex> lock(g_SearchMutex);
            int totalPages = (int)((matches.size() + itemsPerPage - 1) / itemsPerPage);
            if (totalPages < 1) totalPages = 1;
            
            SetDlgItemInt(hDlg, IDC_PAGE_INFO, currentPage, FALSE);
            TCHAR pageTotal[64];
            StringCchPrintf(pageTotal, 64, _T("/ %d"), totalPages);
            SetDlgItemText(hDlg, IDC_PAGE_TOTAL, pageTotal);
            
            EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), currentPage < totalPages);
            lastPageInfoUpdate = GetTickCount();
        }
        return (INT_PTR)TRUE;
    }
    case WM_SEARCH_COMPLETE:
    {
        if (g_SearchState.workerThread.joinable()) {
            g_SearchState.workerThread.join();
        }
        g_SearchState.bActive = FALSE;
        
        if (g_SearchState.bFuzzy)
        {
            std::sort(matches.begin(), matches.end(), [](const SearchMatch& a, const SearchMatch& b) {
                return a.score > b.score;
            });
            UpdateSearchList(hDlg, matches, currentPage, itemsPerPage);
        }
        else
        {
            UpdateSearchList(hDlg, matches, currentPage, itemsPerPage, TRUE);
        }
        
        int totalPages = (int)((matches.size() + itemsPerPage - 1) / itemsPerPage);
        if (totalPages < 1) totalPages = 1;
        
        SetDlgItemInt(hDlg, IDC_PAGE_INFO, currentPage, FALSE);
        TCHAR pageTotal[64];
        StringCchPrintf(pageTotal, 64, _T("/ %d"), totalPages);
        SetDlgItemText(hDlg, IDC_PAGE_TOTAL, pageTotal);
        
        EnableWindow(GetDlgItem(hDlg, IDC_PREV_PAGE), currentPage > 1);
        EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), currentPage < totalPages);
        return (INT_PTR)TRUE;
    }
    case WM_DESTROY:
        if (g_SearchState.bActive) {
            g_SearchState.bActive = FALSE;
            if (g_SearchState.workerThread.joinable()) {
                g_SearchState.workerThread.join();
            }
        }
        if (hDialogListFont) DeleteObject(hDialogListFont);
        matches.clear();
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_SEARCH_REGEX) {
            g_bRegex = IsDlgButtonChecked(hDlg, IDC_SEARCH_REGEX);
        }
        else if (LOWORD(wParam) == IDC_SEARCH_FUZZY) {
            g_bFuzzy = IsDlgButtonChecked(hDlg, IDC_SEARCH_FUZZY);
        }
        
        if (LOWORD(wParam) == IDC_SEARCH_BTN)
        {
            HWND hFocus = GetFocus();
            if (hFocus == GetDlgItem(hDlg, IDC_PAGE_INFO)) {
                 SendMessage(hDlg, WM_COMMAND, MAKEWPARAM(IDC_PAGE_INFO, EN_KILLFOCUS), (LPARAM)hFocus);
                 return (INT_PTR)TRUE;
            }
            StartSearch(hDlg, matches, &currentPage);
        }
        else if (LOWORD(wParam) == IDC_PAGE_INFO && HIWORD(wParam) == EN_KILLFOCUS)
        {
            BOOL bTranslated = FALSE;
            int newPage = GetDlgItemInt(hDlg, IDC_PAGE_INFO, &bTranslated, FALSE);
            if (bTranslated)
            {
                int totalPages = (int)((matches.size() + itemsPerPage - 1) / itemsPerPage);
                if (totalPages < 1) totalPages = 1;
                
                if (newPage < 1) newPage = 1;
                if (newPage > totalPages) newPage = totalPages;
                
                if (newPage != currentPage)
                {
                    currentPage = newPage;
                    UpdateSearchList(hDlg, matches, currentPage, itemsPerPage);
                }
                SetDlgItemInt(hDlg, IDC_PAGE_INFO, currentPage, FALSE);
                EnableWindow(GetDlgItem(hDlg, IDC_PREV_PAGE), currentPage > 1);
                EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), currentPage < totalPages);
            }
        }
        else if (LOWORD(wParam) == IDC_PREV_PAGE)
        {
            if (currentPage > 1)
            {
                currentPage--;
                UpdateSearchList(hDlg, matches, currentPage, itemsPerPage);
                SetDlgItemInt(hDlg, IDC_PAGE_INFO, currentPage, FALSE);
                int totalPages = (int)((matches.size() + itemsPerPage - 1) / itemsPerPage);
                EnableWindow(GetDlgItem(hDlg, IDC_PREV_PAGE), currentPage > 1);
                EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), currentPage < totalPages);
            }
        }
        else if (LOWORD(wParam) == IDC_NEXT_PAGE)
        {
            int totalPages = (int)((matches.size() + itemsPerPage - 1) / itemsPerPage);
            if (currentPage < totalPages)
            {
                currentPage++;
                UpdateSearchList(hDlg, matches, currentPage, itemsPerPage);
                SetDlgItemInt(hDlg, IDC_PAGE_INFO, currentPage, FALSE);
                EnableWindow(GetDlgItem(hDlg, IDC_PREV_PAGE), currentPage > 1);
                EnableWindow(GetDlgItem(hDlg, IDC_NEXT_PAGE), currentPage < totalPages);
            }
        }
        else if (LOWORD(wParam) == IDC_REPLACE_BTN)
        {
            int idx = SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETCURSEL, 0, 0);
            if (idx != LB_ERR)
            {
                int matchIdx = (int)SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETITEMDATA, idx, 0);
                if (matchIdx >= 0 && matchIdx < (int)matches.size())
                {
                    BOOL bFuzzy = IsDlgButtonChecked(hDlg, IDC_SEARCH_FUZZY);
                    if (bFuzzy) {
                        MessageBox(hDlg, _T("Replace not supported in Fuzzy mode."), _T("Info"), MB_ICONINFORMATION);
                        return TRUE;
                    }

                    const auto& m = matches[matchIdx];
                    
                    TCHAR szReplace[256];
                    GetDlgItemText(hDlg, IDC_REPLACE_TEXT, szReplace, 256);
                    
                    if (m.matchStart != -1)
                    {
                        SendMessage(g_hEditor, EM_SETSEL, m.absolutePos, m.absolutePos + m.matchLen);
                        SendMessage(g_hEditor, EM_REPLACESEL, TRUE, (LPARAM)szReplace);
                        
                        StartSearch(hDlg, matches, &currentPage);
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

            if (g_SearchState.bActive)
            {
                 MessageBox(hDlg, _T("Please wait for search to complete."), _T("Info"), MB_ICONINFORMATION);
                 return TRUE;
            }
            
            if (matches.empty()) 
            {
                 StartSearch(hDlg, matches, &currentPage);
                 return TRUE;
            }
            
            TCHAR szReplace[256];
            GetDlgItemText(hDlg, IDC_REPLACE_TEXT, szReplace, 256);
            
            std::sort(matches.begin(), matches.end(), [](const SearchMatch& a, const SearchMatch& b) {
                if (a.lineNum != b.lineNum) return a.lineNum > b.lineNum;
                return a.matchStart > b.matchStart;
            });
            
            int count = 0;
            for (const auto& m : matches)
            {
                if (m.matchStart != -1)
                {
                    SendMessage(g_hEditor, EM_SETSEL, m.absolutePos, m.absolutePos + m.matchLen);
                    SendMessage(g_hEditor, EM_REPLACESEL, TRUE, (LPARAM)szReplace);
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
                    int matchIdx = (int)SendDlgItemMessage(hDlg, IDC_SEARCH_LIST, LB_GETITEMDATA, idx, 0);
                    if (matchIdx >= 0 && matchIdx < (int)matches.size())
                    {
                        const auto& m = matches[matchIdx];
                        
                        if (m.matchStart != -1)
                        {
                            SendMessage(g_hEditor, EM_SETSEL, m.absolutePos, m.absolutePos);
                            SendMessage(g_hEditor, EM_SCROLLCARET, 0, 0);
                            SendMessage(g_hEditor, EM_SETSEL, m.absolutePos, m.absolutePos + m.matchLen);
                        }
                        else
                        {
                            int nIndex = SendMessage(g_hEditor, EM_LINEINDEX, m.lineNum, 0);
                            if (nIndex != -1)
                            {
                                long nLineLen = SendMessage(g_hEditor, EM_LINELENGTH, nIndex, 0);
                                SendMessage(g_hEditor, EM_SETSEL, nIndex, nIndex + nLineLen);
                                SendMessage(g_hEditor, EM_SCROLLCARET, 0, 0);
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

void ShowAdvancedSearch(HWND hEditor) {
    g_hEditor = hEditor;
    g_hMain = GetParent(hEditor);
    DialogBox(g_hInst, MAKEINTRESOURCE(IDD_ADVANCED_SEARCH), g_hMain, AdvancedSearchDlgProc);
}

PluginMenuItem g_MenuItems[] = {
    { _T("Advanced Search..."), ShowAdvancedSearch, _T("Ctrl+Shift+F") }
};

extern "C" {

PLUGIN_API const TCHAR* GetPluginName() {
    return _T("Advanced Search");
}

PLUGIN_API const TCHAR* GetPluginDescription() {
    return _T("Provides advanced search capabilities including Regex and Fuzzy matching.");
}

PLUGIN_API const TCHAR* GetPluginVersion() {
    return _T("1.0");
}

PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
    *count = 1;
    return g_MenuItems;
}

} // extern "C"

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}
