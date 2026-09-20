#pragma once
#include <windows.h>

// Messages posted to the main window by the updater worker thread.
#define WM_UPDATE_STATUS   (WM_APP + 310)  // wParam = percent (0..100), -1 = hide progress
#define WM_UPDATE_READY    (WM_APP + 311)  // lParam = wchar_t* zip path (free with Updater_FreeString)
#define WM_UPDATE_DONE     (WM_APP + 312)  // wParam = UpdateResult, lParam = wchar_t* message

enum UpdateResult {
    UPDATE_UP_TO_DATE = 0,
    UPDATE_FAILED = 1,
    UPDATE_DECLINED = 2
};

// Starts the check on a background thread. If bSilent, nothing is reported
// unless an update is actually available.
void Updater_CheckForUpdates(HWND hMain, bool bSilent);

// True while a check/download is in flight.
bool Updater_IsBusy();

// Called on the main thread once the new package has been downloaded.
// Spawns the external applier and returns true if it was started successfully.
bool Updater_LaunchApplier(const wchar_t* pszZipPath, HWND hMain);

void Updater_FreeString(wchar_t* p);

// Clears the busy flag when the user declines or the flow ends early.
void Updater_Reset();
