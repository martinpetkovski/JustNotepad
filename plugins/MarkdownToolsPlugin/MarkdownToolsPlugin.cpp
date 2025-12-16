#include "../../src/PluginInterface.h"
#include <windows.h>
#include <richedit.h>
#include <shellapi.h>
#include <string>
#include <vector>
#include <sstream>

// Global variables
HINSTANCE g_hModule = NULL;
HMODULE g_hRichEdit = NULL;
const wchar_t* CLASS_NAME = L"MDViewerWindowClass";

struct ViewerContext {
    HWND hRichEdit;
    std::vector<std::pair<CHARRANGE, std::wstring>> links;
};

// Helper to append text with format
void AppendText(HWND hEdit, const std::wstring& text, CHARFORMAT2W& cf, PARAFORMAT2& pf) {
    // Set selection to end
    int len = GetWindowTextLength(hEdit);
    SendMessage(hEdit, EM_SETSEL, len, len);

    // Set format for insertion
    SendMessage(hEdit, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
    SendMessage(hEdit, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
    
    // Insert text
    SendMessageW(hEdit, EM_REPLACESEL, FALSE, (LPARAM)text.c_str());
}

// Helper to parse inline formatting
void ParseInlineFormatting(HWND hEdit, const std::wstring& text, CHARFORMAT2W& cf, PARAFORMAT2& pf, ViewerContext* ctx, bool appendNewline = true) {
    size_t pos = 0;
    size_t len = text.length();
    
    while (pos < len) {
        // Check for bold+italic (*** or ___)
        if (pos + 2 < len && ((text[pos] == L'*' && text[pos+1] == L'*' && text[pos+2] == L'*') || (text[pos] == L'_' && text[pos+1] == L'_' && text[pos+2] == L'_'))) {
            size_t end = text.find(text.substr(pos, 3), pos + 3);
            if (end != std::wstring::npos) {
                // Found bold+italic text
                cf.dwEffects |= CFE_BOLD | CFE_ITALIC;
                // Recursively parse content inside
                ParseInlineFormatting(hEdit, text.substr(pos + 3, end - (pos + 3)), cf, pf, ctx, false);
                cf.dwEffects &= ~(CFE_BOLD | CFE_ITALIC);
                pos = end + 3;
                continue;
            }
        }

        // Check for bold (** or __)
        if (pos + 1 < len && ((text[pos] == L'*' && text[pos+1] == L'*') || (text[pos] == L'_' && text[pos+1] == L'_'))) {
            size_t end = text.find(text.substr(pos, 2), pos + 2);
            if (end != std::wstring::npos) {
                // Found bold text
                cf.dwEffects |= CFE_BOLD;
                // Recursively parse content inside bold
                ParseInlineFormatting(hEdit, text.substr(pos + 2, end - (pos + 2)), cf, pf, ctx, false);
                cf.dwEffects &= ~CFE_BOLD;
                pos = end + 2;
                continue;
            }
        }
        
        // Check for italic (* or _)
        if ((text[pos] == L'*' || text[pos] == L'_')) {
            size_t end = text.find(text[pos], pos + 1);
            if (end != std::wstring::npos) {
                // Found italic text
                cf.dwEffects |= CFE_ITALIC;
                // Recursively parse content inside italic
                ParseInlineFormatting(hEdit, text.substr(pos + 1, end - (pos + 1)), cf, pf, ctx, false);
                cf.dwEffects &= ~CFE_ITALIC;
                pos = end + 1;
                continue;
            }
        }

        // Check for code (` )
        if (text[pos] == L'`') {
            size_t end = text.find(L'`', pos + 1);
            if (end != std::wstring::npos) {
                // Found inline code
                wcscpy_s(cf.szFaceName, L"Consolas");
                cf.crBackColor = RGB(240, 240, 240);
                AppendText(hEdit, text.substr(pos + 1, end - (pos + 1)), cf, pf);
                // Reset font
                wcscpy_s(cf.szFaceName, L"Segoe UI");
                cf.crBackColor = RGB(255, 255, 255);
                pos = end + 1;
                continue;
            }
        }

        // Check for link [text](url)
        if (text[pos] == L'[') {
            size_t endText = text.find(L']', pos + 1);
            if (endText != std::wstring::npos && endText + 1 < len && text[endText + 1] == L'(') {
                size_t endUrl = text.find(L')', endText + 2);
                if (endUrl != std::wstring::npos) {
                    // Found link
                    std::wstring linkText = text.substr(pos + 1, endText - (pos + 1));
                    std::wstring linkUrl = text.substr(endText + 2, endUrl - (endText + 2));
                    
                    // Render link text
                    cf.dwEffects |= CFE_LINK | CFE_UNDERLINE;
                    cf.crTextColor = RGB(0, 0, 255);
                    
                    long startLen = GetWindowTextLength(hEdit);

                    // Parse content inside link (recursively)
                    ParseInlineFormatting(hEdit, linkText, cf, pf, ctx, false);
                    
                    long endLen = GetWindowTextLength(hEdit);
                    
                    if (ctx) {
                        CHARRANGE cr;
                        cr.cpMin = startLen;
                        cr.cpMax = endLen;
                        ctx->links.push_back({cr, linkUrl});
                    }
                    
                    // Reset format
                    cf.dwEffects &= ~(CFE_LINK | CFE_UNDERLINE);
                    cf.crTextColor = RGB(0, 0, 0);
                    
                    pos = endUrl + 1;
                    continue;
                }
            }
        }

        // Normal character
        // Find next special char to append chunk
        size_t nextSpecial = std::wstring::npos;
        size_t nextBoldItalic = text.find(L"***", pos);
        size_t nextBold = text.find(L"**", pos);
        size_t nextItalic = text.find_first_of(L"*_", pos);
        size_t nextCode = text.find(L'`', pos);
        size_t nextLink = text.find(L'[', pos);

        size_t next = len;
        if (nextBoldItalic != std::wstring::npos && nextBoldItalic < next) next = nextBoldItalic;
        if (nextBold != std::wstring::npos && nextBold < next) next = nextBold;
        if (nextItalic != std::wstring::npos && nextItalic < next) next = nextItalic;
        if (nextCode != std::wstring::npos && nextCode < next) next = nextCode;
        if (nextLink != std::wstring::npos && nextLink < next) next = nextLink;

        if (next > pos) {
            AppendText(hEdit, text.substr(pos, next - pos), cf, pf);
            pos = next;
        } else {
            AppendText(hEdit, text.substr(pos, 1), cf, pf);
            pos++;
        }
    }
    if (appendNewline) {
        AppendText(hEdit, L"\n", cf, pf);
    }
}

void ParseAndRender(HWND hEdit, const std::wstring& markdown, ViewerContext* ctx) {
    // Clear existing text
    SetWindowTextW(hEdit, L"");
    
    // Clear links
    if (ctx) ctx->links.clear();
    
    // Enable Link detection
    SendMessage(hEdit, EM_AUTOURLDETECT, FALSE, 0); // Disable auto-detect since we handle it manually
    SendMessage(hEdit, EM_SETEVENTMASK, 0, ENM_LINK | ENM_MOUSEEVENTS | ENM_KEYEVENTS);

    std::wstringstream ss(markdown);
    std::wstring line;
    bool inCodeBlock = false;

    CHARFORMAT2W cf = { sizeof(CHARFORMAT2W) };
    cf.dwMask = CFM_FACE | CFM_SIZE | CFM_BOLD | CFM_ITALIC | CFM_COLOR | CFM_BACKCOLOR | CFM_LINK;
    
    PARAFORMAT2 pf = { sizeof(PARAFORMAT2) };
    pf.dwMask = PFM_ALIGNMENT | PFM_NUMBERING | PFM_OFFSET | PFM_STARTINDENT | PFM_TABSTOPS;

    while (std::getline(ss, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();

        // Reset defaults
        cf.dwEffects = 0;
        cf.yHeight = 200; // 10pt
        wcscpy_s(cf.szFaceName, L"Segoe UI");
        cf.crTextColor = RGB(0, 0, 0);
        cf.crBackColor = RGB(255, 255, 255);
        
        pf.wNumbering = 0;
        pf.dxStartIndent = 0;
        pf.dxOffset = 0;
        pf.cTabCount = 0;

        if (line.substr(0, 3) == L"```") {
            inCodeBlock = !inCodeBlock;
            continue;
        }

        if (inCodeBlock) {
            wcscpy_s(cf.szFaceName, L"Consolas");
            cf.crBackColor = RGB(245, 245, 245);
            AppendText(hEdit, line + L"\n", cf, pf);
            continue;
        }

        if (line.empty()) {
            AppendText(hEdit, L"\n", cf, pf);
            continue;
        }

        // Table row detection (starts with |)
        // Allow leading whitespace
        size_t firstChar = line.find_first_not_of(L" \t");
        if (firstChar != std::wstring::npos && line[firstChar] == L'|') {
            // Render as table row
            wcscpy_s(cf.szFaceName, L"Consolas"); // Use monospaced for alignment
            
            // Set tab stops
            pf.cTabCount = 10; // Support up to 10 columns
            for(int i=0; i<10; i++) pf.rgxTabs[i] = (i+1) * 1440 * 1.5; // 1.5 inches per column
            
            std::wstring rowText;
            bool firstPipe = true;
            for (size_t i = firstChar; i < line.length(); i++) {
                if (line[i] == L'|') {
                    if (!firstPipe) {
                        rowText += L"\t";
                    }
                    firstPipe = false;
                } else {
                    rowText += line[i];
                }
            }
            
            AppendText(hEdit, rowText + L"\n", cf, pf);
            continue;
        }

        // Horizontal Rule
        if (line == L"---" || line == L"***" || line == L"___") {
             // Render a line
             AppendText(hEdit, L"__________________________________________________\n", cf, pf);
             continue;
        }

        // Headers
        if (line.substr(0, 2) == L"# ") {
            cf.dwEffects |= CFE_BOLD;
            cf.yHeight = 320; // 16pt
            ParseInlineFormatting(hEdit, line.substr(2), cf, pf, ctx);
        } else if (line.substr(0, 3) == L"## ") {
            cf.dwEffects |= CFE_BOLD;
            cf.yHeight = 280; // 14pt
            ParseInlineFormatting(hEdit, line.substr(3), cf, pf, ctx);
        } else if (line.substr(0, 4) == L"### ") {
            cf.dwEffects |= CFE_BOLD;
            cf.yHeight = 240; // 12pt
            ParseInlineFormatting(hEdit, line.substr(4), cf, pf, ctx);
        } else if (line.substr(0, 5) == L"#### ") {
            cf.dwEffects |= CFE_BOLD;
            cf.yHeight = 220; // 11pt
            ParseInlineFormatting(hEdit, line.substr(5), cf, pf, ctx);
        } else if (line.substr(0, 6) == L"##### ") {
            cf.dwEffects |= CFE_BOLD;
            cf.yHeight = 200; // 10pt
            ParseInlineFormatting(hEdit, line.substr(6), cf, pf, ctx);
        } else if (line.substr(0, 7) == L"###### ") {
            cf.dwEffects |= CFE_BOLD;
            cf.dwEffects |= CFE_ITALIC;
            cf.yHeight = 200; // 10pt
            ParseInlineFormatting(hEdit, line.substr(7), cf, pf, ctx);
        }
        // List
        else if (firstChar != std::wstring::npos && (line.substr(firstChar, 2) == L"- " || line.substr(firstChar, 2) == L"* ")) {
            // Check for task list
            if (line.length() > firstChar + 5 && line.substr(firstChar, 4) == L"- [ ]") {
                pf.dxStartIndent = (LONG)(firstChar * 100); // Indent based on spaces
                AppendText(hEdit, L"\u2610 " + line.substr(firstChar + 6) + L"\n", cf, pf); // Ballot box
            } else if (line.length() > firstChar + 5 && (line.substr(firstChar, 4) == L"- [x]" || line.substr(firstChar, 4) == L"- [X]")) {
                pf.dxStartIndent = (LONG)(firstChar * 100); // Indent based on spaces
                AppendText(hEdit, L"\u2611 " + line.substr(firstChar + 6) + L"\n", cf, pf); // Ballot box with check
            } else {
                pf.wNumbering = PFN_BULLET;
                pf.dxOffset = 100;
                pf.dxStartIndent = (LONG)(firstChar * 100 + 200); // Indent based on spaces
                ParseInlineFormatting(hEdit, line.substr(firstChar + 2), cf, pf, ctx);
            }
        }
        // Blockquote
        else if (line.substr(0, 2) == L"> ") {
            pf.dxStartIndent = 200; // Indent
            cf.crTextColor = RGB(100, 100, 100);
            cf.dwEffects |= CFE_ITALIC;
            ParseInlineFormatting(hEdit, line.substr(2), cf, pf, ctx);
        }
        else {
            // Normal text
            ParseInlineFormatting(hEdit, line, cf, pf, ctx);
        }
    }
}

LRESULT CALLBACK ViewerWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    ViewerContext* ctx = (ViewerContext*)GetWindowLongPtr(hwnd, GWLP_USERDATA);

    switch (uMsg) {
        case WM_CREATE: {
            ctx = new ViewerContext();
            RECT rcClient;
            GetClientRect(hwnd, &rcClient);
            ctx->hRichEdit = CreateWindowExW(0, MSFTEDIT_CLASS, L"", 
                WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
                0, 0, rcClient.right, rcClient.bottom,
                hwnd, NULL, g_hModule, NULL);
            
            // Set margins
            RECT rcMargin = { 10, 10, 10, 10 };
            SendMessage(ctx->hRichEdit, EM_SETRECT, 0, (LPARAM)&rcMargin);
            
            // Store pointer to Context in window user data
            SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)ctx);
            return 0;
        }
        case WM_SIZE: {
            if (ctx && ctx->hRichEdit) {
                MoveWindow(ctx->hRichEdit, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
            }
            return 0;
        }
        case WM_NOTIFY: {
            NMHDR* pnmh = (NMHDR*)lParam;
            if (pnmh->code == EN_LINK) {
                ENLINK* penLink = (ENLINK*)pnmh;
                if (penLink->msg == WM_LBUTTONUP) {
                    if (ctx) {
                        for (const auto& link : ctx->links) {
                            // Check if clicked range is within a stored link range
                            // Note: RichEdit might select a smaller range than the full link text
                            if (penLink->chrg.cpMin >= link.first.cpMin && penLink->chrg.cpMin <= link.first.cpMax) {
                                ShellExecuteW(NULL, L"open", link.second.c_str(), NULL, NULL, SW_SHOWNORMAL);
                                return 0;
                            }
                        }
                    }
                }
            }
            break;
        }
        case WM_SETCURSOR: {
            // If the cursor is over a link, set it to a hand
            // We need to check if the mouse is over a link
            // But RichEdit handles this if EN_LINK is enabled and we return 0?
            // Actually, if we handle EN_LINK, we might need to set cursor manually if auto-detect is off.
            // But CFE_LINK usually handles it.
            // Let's try to force it if we are over a link.
            // Getting char index from point is needed.
            // For now, let's rely on RichEdit default behavior for CFE_LINK.
            break;
        }
        case WM_CLOSE:
            DestroyWindow(hwnd);
            return 0;
        case WM_DESTROY:
            if (ctx) delete ctx;
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

// Helper to wrap selection with prefix/suffix
void WrapSelection(HWND hEditor, const std::wstring& left, const std::wstring& right) {
    DWORD start, end;
    SendMessage(hEditor, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
    
    // Get selected text
    int len = end - start;
    std::wstring text;
    if (len > 0) {
        std::vector<wchar_t> buf(len + 1);
        SendMessage(hEditor, EM_GETSELTEXT, 0, (LPARAM)buf.data());
        text = buf.data();
    }
    
    std::wstring newText = left + text + right;
    SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)newText.c_str());
    
    // Select the inner text
    SendMessage(hEditor, EM_SETSEL, start + left.length(), start + left.length() + text.length());
}

// Helper to insert at start of line(s)
void InsertAtStartOfLines(HWND hEditor, const std::wstring& prefix) {
    DWORD start, end;
    SendMessage(hEditor, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
    
    long startLine = SendMessage(hEditor, EM_LINEFROMCHAR, start, 0);
    long endLine = SendMessage(hEditor, EM_LINEFROMCHAR, end, 0);
    
    // If end is at the start of a line (and not same as start), exclude that line
    if (endLine > startLine) {
        long lineIndex = SendMessage(hEditor, EM_LINEINDEX, endLine, 0);
        if (lineIndex == end) {
            endLine--;
        }
    }
    
    // Process lines from bottom to top to avoid index shifting issues affecting next lines
    for (long i = endLine; i >= startLine; i--) {
        long lineStart = SendMessage(hEditor, EM_LINEINDEX, i, 0);
        SendMessage(hEditor, EM_SETSEL, lineStart, lineStart);
        SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)prefix.c_str());
    }
    
    // Restore selection (approximate)
    // It's hard to restore exact selection because indices shifted.
    // Let's just select the whole block.
    long newStart = SendMessage(hEditor, EM_LINEINDEX, startLine, 0);
    long newEndLineStart = SendMessage(hEditor, EM_LINEINDEX, endLine, 0);
    long newEndLineLen = SendMessage(hEditor, EM_LINELENGTH, newEndLineStart, 0);
    SendMessage(hEditor, EM_SETSEL, newStart, newEndLineStart + newEndLineLen);
}

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() {
        return L"Markdown Tools";
    }
    
    PLUGIN_API const wchar_t* GetPluginDescription() {
        return L"Markdown viewing and editing tools.";
    }
    
    PLUGIN_API const wchar_t* GetPluginVersion() {
        return L"1.2";
    }

    PLUGIN_API const wchar_t* GetPluginLicense() {
        return L"MIT License";
    }

    void RenderMarkdown(HWND hEditorWnd) {
        // Get text length
        int len = GetWindowTextLengthW(hEditorWnd);
        if (len <= 0) return;

        // Get text
        std::vector<wchar_t> buffer(len + 1);
        GetWindowTextW(hEditorWnd, buffer.data(), len + 1);
        std::wstring markdown(buffer.data());

        // Register Window Class
        WNDCLASSEXW wc = {0};
        wc.cbSize = sizeof(WNDCLASSEXW);
        wc.lpfnWndProc = ViewerWndProc;
        wc.hInstance = g_hModule;
        wc.lpszClassName = CLASS_NAME;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        RegisterClassExW(&wc);

        // Create Window
        HWND hViewer = CreateWindowExW(
            WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
            CLASS_NAME,
            L"Markdown Viewer",
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
            hEditorWnd, // Parent
            NULL,
            g_hModule,
            NULL
        );

        if (!hViewer) {
            MessageBoxW(hEditorWnd, L"Failed to create viewer window.", L"Error", MB_ICONERROR);
            return;
        }

        // Get Context
        ViewerContext* ctx = (ViewerContext*)GetWindowLongPtr(hViewer, GWLP_USERDATA);
        if (ctx && ctx->hRichEdit) {
            ParseAndRender(ctx->hRichEdit, markdown, ctx);
        }

        // Show Window
        ShowWindow(hViewer, SW_SHOW);
        UpdateWindow(hViewer);

        // Disable parent to simulate modal behavior
        EnableWindow(hEditorWnd, FALSE);

        // Message Loop
        MSG msg;
        while (GetMessage(&msg, NULL, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }

        // Re-enable parent
        EnableWindow(hEditorWnd, TRUE);
        SetForegroundWindow(hEditorWnd);
    }

    void FormatBold(HWND hEditor) { WrapSelection(hEditor, L"**", L"**"); }
    void FormatItalic(HWND hEditor) { WrapSelection(hEditor, L"*", L"*"); }
    void FormatStrikethrough(HWND hEditor) { WrapSelection(hEditor, L"~~", L"~~"); }
    void FormatCode(HWND hEditor) { WrapSelection(hEditor, L"`", L"`"); }
    
    void FormatLink(HWND hEditor) {
        DWORD start, end;
        SendMessage(hEditor, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
        if (start == end) {
            SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)L"[](url)");
            SendMessage(hEditor, EM_SETSEL, start + 1, start + 1); // Cursor between brackets
        } else {
            WrapSelection(hEditor, L"[", L"](url)");
            // Select "url" part? No, WrapSelection selects inner text.
            // Let's just leave it.
        }
    }
    
    void FormatImage(HWND hEditor) {
        SendMessage(hEditor, EM_REPLACESEL, TRUE, (LPARAM)L"![alt](url)");
    }

    void FormatH1(HWND hEditor) { InsertAtStartOfLines(hEditor, L"# "); }
    void FormatH2(HWND hEditor) { InsertAtStartOfLines(hEditor, L"## "); }
    void FormatH3(HWND hEditor) { InsertAtStartOfLines(hEditor, L"### "); }
    void FormatList(HWND hEditor) { InsertAtStartOfLines(hEditor, L"- "); }
    void FormatTaskList(HWND hEditor) { InsertAtStartOfLines(hEditor, L"- [ ] "); }
    void FormatBlockquote(HWND hEditor) { InsertAtStartOfLines(hEditor, L"> "); }

    PluginMenuItem g_Items[] = {
        { L"MD Viewer", RenderMarkdown, L"Ctrl+Shift+M" },
        { L"Bold", FormatBold, L"Ctrl+B" },
        { L"Italic", FormatItalic, L"Ctrl+I" },
        { L"Strikethrough", FormatStrikethrough, NULL },
        { L"Code", FormatCode, NULL },
        { L"Link", FormatLink, L"Ctrl+K" },
        { L"Image", FormatImage, NULL },
        { L"Heading 1", FormatH1, NULL },
        { L"Heading 2", FormatH2, NULL },
        { L"Heading 3", FormatH3, NULL },
        { L"List", FormatList, NULL },
        { L"Task List", FormatTaskList, NULL },
        { L"Blockquote", FormatBlockquote, NULL }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_Items) / sizeof(g_Items[0]);
        return g_Items;
    }
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        g_hModule = hModule;
        g_hRichEdit = LoadLibraryW(L"Msftedit.dll");
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
        break;
    case DLL_PROCESS_DETACH:
        if (g_hRichEdit) {
            FreeLibrary(g_hRichEdit);
        }
        break;
    }
    return TRUE;
}
