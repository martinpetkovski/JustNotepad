#pragma once
// A terminal panel: PowerShell running on a Windows pseudo console (ConPTY),
// drawn by us. The shell is a real console program, so colours, the prompt,
// Tab completion, history and interactive tools all work as in a console.
#include <windows.h>

#include <string>

#include "theme.h"

// Posted to the parent window: wParam is one of these, lParam the new height
// for TERM_RESIZE.
const UINT WM_APP_TERM = WM_APP + 40;
enum { TERM_CLOSE = 1, TERM_RESIZE = 2, TERM_TOGGLE = 3 };

// Registers the window class and creates the (hidden) panel.
HWND termCreate(HWND parent, HINSTANCE inst);
// Font (fixed pitch), its cell size and colours. Call again when they change.
void termSetStyle(HFONT font, HFONT uiFont, int cw, int ch, int headerH, const Palette& pal);
// Starts the shell in `folder` unless it is already running.
void termStart(const std::wstring& folder);
bool termRunning();
// Makes the shell go to `folder` (when it is not already there).
void termCd(const std::wstring& folder);
// Stops the shell (when the app closes).
void termShutdown();
