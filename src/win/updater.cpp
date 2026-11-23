#ifndef UNICODE
#define UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include "updater.h"
#include "version.h"

#include <windows.h>
#include <winhttp.h>
#include <shellapi.h>
#include <bcrypt.h>
#include <string>
#include <vector>
#include <atomic>
#include <thread>
#include <cwctype>
#include <cstdlib>
#include <cstring>

#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "bcrypt.lib")

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------
static const wchar_t* kApiHost = L"api.github.com";
static const wchar_t* kApiPath = L"/repos/martinpetkovski/JustNotepad/releases/latest";
static const wchar_t* kUserAgent = L"JustNotepad-Updater";

#define UPD_WIDEN2(x) L##x
#define UPD_WIDEN(x) UPD_WIDEN2(x)

static std::atomic<bool> g_busy(false);

void Updater_FreeString(wchar_t* p) { free(p); }

bool Updater_IsBusy() { return g_busy.load(); }

void Updater_Reset() { g_busy = false; }

static wchar_t* DupString(const std::wstring& s)
{
    wchar_t* p = (wchar_t*)malloc((s.size() + 1) * sizeof(wchar_t));
    if (p) memcpy(p, s.c_str(), (s.size() + 1) * sizeof(wchar_t));
    return p;
}

// ---------------------------------------------------------------------------
// Tiny helpers: version compare + minimal JSON field extraction
// ---------------------------------------------------------------------------
static std::vector<int> ParseVersion(const std::wstring& in)
{
    std::vector<int> parts;
    size_t i = 0;
    while (i < in.size() && !iswdigit(in[i])) ++i; // skip a leading 'v'
    int cur = 0;
    bool has = false;
    for (; i < in.size(); ++i)
    {
        if (iswdigit(in[i])) { cur = cur * 10 + (in[i] - L'0'); has = true; }
        else if (in[i] == L'.') { parts.push_back(has ? cur : 0); cur = 0; has = false; }
        else break;
    }
    if (has) parts.push_back(cur);
    return parts;
}

// > 0 if a is newer than b
static int CompareVersions(const std::wstring& a, const std::wstring& b)
{
    std::vector<int> va = ParseVersion(a), vb = ParseVersion(b);
    size_t n = va.size() > vb.size() ? va.size() : vb.size();
    for (size_t i = 0; i < n; ++i)
    {
        int x = i < va.size() ? va[i] : 0;
        int y = i < vb.size() ? vb[i] : 0;
        if (x != y) return x < y ? -1 : 1;
    }
    return 0;
}

static std::wstring JsonUnescape(const std::wstring& in)
{
    std::wstring out;
    out.reserve(in.size());
    for (size_t i = 0; i < in.size(); ++i)
    {
        if (in[i] == L'\\' && i + 1 < in.size())
        {
            wchar_t c = in[++i];
            switch (c)
            {
            case L'n': out += L'\n'; break;
            case L'r': out += L'\r'; break;
            case L't': out += L'\t'; break;
            case L'u':
                if (i + 4 < in.size())
                {
                    int v = 0;
                    for (int k = 1; k <= 4; ++k)
                    {
                        wchar_t h = in[i + k];
                        int d = (h >= L'0' && h <= L'9') ? h - L'0'
                              : (h >= L'a' && h <= L'f') ? h - L'a' + 10
                              : (h >= L'A' && h <= L'F') ? h - L'A' + 10 : 0;
                        v = v * 16 + d;
                    }
                    out += (wchar_t)v;
                    i += 4;
                }
                break;
            default: out += c; break;
            }
        }
        else out += in[i];
    }
    return out;
}

// Returns the string value of the first occurrence of "key" at/after "from".
static bool JsonFindString(const std::wstring& json, const wchar_t* key,
                           std::wstring& out, size_t from = 0, size_t* pEnd = nullptr)
{
    std::wstring needle = std::wstring(L"\"") + key + L"\"";
    size_t k = json.find(needle, from);
    if (k == std::wstring::npos) return false;
    size_t c = json.find(L':', k + needle.size());
    if (c == std::wstring::npos) return false;
    size_t q = json.find(L'"', c);
    if (q == std::wstring::npos) return false;
    size_t p = q + 1;
    std::wstring raw;
    while (p < json.size())
    {
        if (json[p] == L'\\' && p + 1 < json.size()) { raw += json[p]; raw += json[p + 1]; p += 2; continue; }
        if (json[p] == L'"') break;
        raw += json[p++];
    }
    out = JsonUnescape(raw);
    if (pEnd) *pEnd = p;
    return true;
}

// ---------------------------------------------------------------------------
// HTTP
// ---------------------------------------------------------------------------
struct UrlParts { std::wstring host, path; INTERNET_PORT port; bool https; };

static bool SplitUrl(const std::wstring& url, UrlParts& out)
{
    URL_COMPONENTS uc = { 0 };
    wchar_t host[256] = { 0 };
    wchar_t path[4096] = { 0 };
    wchar_t extra[2048] = { 0 };
    uc.dwStructSize = sizeof(uc);
    uc.lpszHostName = host;  uc.dwHostNameLength = 255;
    uc.lpszUrlPath = path;   uc.dwUrlPathLength = 4095;
    uc.lpszExtraInfo = extra; uc.dwExtraInfoLength = 2047;
    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &uc)) return false;
    out.host = host;
    out.path = std::wstring(path) + extra;
    out.port = uc.nPort;
    out.https = (uc.nScheme == INTERNET_SCHEME_HTTPS);
    return true;
}

// Downloads url. If pFile is given the body is written there (with progress
// posted to hMain), otherwise it is appended to *pBody as UTF-8 -> wide.
static bool HttpGet(const std::wstring& url, std::wstring* pBody,
                    const std::wstring* pFile, HWND hMain)
{
    UrlParts u;
    if (!SplitUrl(url, u)) return false;

    HINTERNET hSession = WinHttpOpen(kUserAgent, WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY,
                                     WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession)
        hSession = WinHttpOpen(kUserAgent, WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                               WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) return false;

    DWORD timeout = 30000;
    WinHttpSetTimeouts(hSession, 10000, 15000, timeout, timeout);

    bool ok = false;
    HANDLE hFile = INVALID_HANDLE_VALUE;
    HINTERNET hConnect = WinHttpConnect(hSession, u.host.c_str(), u.port, 0);
    if (hConnect)
    {
        HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", u.path.c_str(), NULL,
                                               WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                                               u.https ? WINHTTP_FLAG_SECURE : 0);
        if (hRequest)
        {
            const wchar_t* headers = L"Accept: application/vnd.github+json\r\n";
            if (WinHttpSendRequest(hRequest, headers, (DWORD)-1L, WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
                WinHttpReceiveResponse(hRequest, NULL))
            {
                DWORD status = 0, len = sizeof(status);
                WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                    WINHTTP_HEADER_NAME_BY_INDEX, &status, &len, WINHTTP_NO_HEADER_INDEX);
                if (status == 200)
                {
                    DWORD total = 0; len = sizeof(total);
                    WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_CONTENT_LENGTH | WINHTTP_QUERY_FLAG_NUMBER,
                                        WINHTTP_HEADER_NAME_BY_INDEX, &total, &len, WINHTTP_NO_HEADER_INDEX);

                    if (pFile)
                    {
                        hFile = CreateFileW(pFile->c_str(), GENERIC_WRITE, 0, NULL,
                                            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
                        if (hFile == INVALID_HANDLE_VALUE)
                        {
                            WinHttpCloseHandle(hRequest);
                            WinHttpCloseHandle(hConnect);
                            WinHttpCloseHandle(hSession);
                            return false;
                        }
                    }

                    std::string raw;
                    std::vector<char> buf(64 * 1024);
                    DWORD got = 0, soFar = 0;
                    int lastPct = -1;
                    ok = true;
                    for (;;)
                    {
                        if (!WinHttpReadData(hRequest, buf.data(), (DWORD)buf.size(), &got)) { ok = false; break; }
                        if (got == 0) break;
                        soFar += got;
                        if (hFile != INVALID_HANDLE_VALUE)
                        {
                            DWORD written = 0;
                            if (!WriteFile(hFile, buf.data(), got, &written, NULL) || written != got) { ok = false; break; }
                        }
                        else raw.append(buf.data(), got);

                        if (hMain && total > 0)
                        {
                            int pct = (int)((soFar * 100ULL) / total);
                            if (pct != lastPct) { lastPct = pct; PostMessageW(hMain, WM_UPDATE_STATUS, (WPARAM)pct, 0); }
                        }
                    }
                    if (ok && pBody && !raw.empty())
                    {
                        int need = MultiByteToWideChar(CP_UTF8, 0, raw.data(), (int)raw.size(), NULL, 0);
                        std::wstring w(need, L'\0');
                        MultiByteToWideChar(CP_UTF8, 0, raw.data(), (int)raw.size(), &w[0], need);
                        *pBody = w;
                    }
                }
            }
            WinHttpCloseHandle(hRequest);
        }
        WinHttpCloseHandle(hConnect);
    }
    WinHttpCloseHandle(hSession);
    if (hFile != INVALID_HANDLE_VALUE) CloseHandle(hFile);
    if (!ok && pFile) DeleteFileW(pFile->c_str());
    return ok;
}

// ---------------------------------------------------------------------------
// Integrity: SHA-256 of the downloaded file must match the published one.
// ---------------------------------------------------------------------------
static bool Sha256File(const std::wstring& path, std::wstring& hexOut)
{
    BCRYPT_ALG_HANDLE alg = nullptr;
    BCRYPT_HASH_HANDLE h = nullptr;
    if (BCryptOpenAlgorithmProvider(&alg, BCRYPT_SHA256_ALGORITHM, nullptr, 0) != 0) return false;
    bool ok = BCryptCreateHash(alg, &h, nullptr, 0, nullptr, 0, 0) == 0;
    HANDLE f = ok ? CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr)
                  : INVALID_HANDLE_VALUE;
    ok = ok && f != INVALID_HANDLE_VALUE;
    std::vector<unsigned char> buf(1 << 16);
    while (ok) {
        DWORD got = 0;
        if (!ReadFile(f, buf.data(), (DWORD)buf.size(), &got, nullptr)) { ok = false; break; }
        if (got == 0) break;
        ok = BCryptHashData(h, buf.data(), got, 0) == 0;
    }
    unsigned char digest[32] = {};
    ok = ok && BCryptFinishHash(h, digest, sizeof(digest), 0) == 0;
    if (f != INVALID_HANDLE_VALUE) CloseHandle(f);
    if (h) BCryptDestroyHash(h);
    BCryptCloseAlgorithmProvider(alg, 0);
    if (!ok) return false;
    static const wchar_t* hex = L"0123456789abcdef";
    hexOut.clear();
    for (unsigned char b : digest) { hexOut += hex[b >> 4]; hexOut += hex[b & 15]; }
    return true;
}

// The first 64 hex digits of a ".sha256" file ("<hash>  <name>"), lower case.
static std::wstring ParseSha256(const std::wstring& text)
{
    std::wstring out;
    for (wchar_t c : text) {
        if (iswxdigit(c)) out += (wchar_t)towlower(c);
        else if (!out.empty()) break;
        if (out.size() == 64) break;
    }
    return out.size() == 64 ? out : std::wstring();
}

// ---------------------------------------------------------------------------
// Worker
// ---------------------------------------------------------------------------
static std::wstring GetUpdateTempDir()
{
    wchar_t tmp[MAX_PATH] = { 0 };
    GetTempPathW(MAX_PATH, tmp);
    std::wstring dir = std::wstring(tmp) + L"JustNotepadUpdate";
    CreateDirectoryW(dir.c_str(), NULL);
    return dir;
}

static void Finish(HWND hMain, UpdateResult r, const std::wstring& msg)
{
    PostMessageW(hMain, WM_UPDATE_STATUS, (WPARAM)-1, 0);
    PostMessageW(hMain, WM_UPDATE_DONE, (WPARAM)r, (LPARAM)DupString(msg));
}

static void UpdateWorker(HWND hMain, bool bSilent)
{
    std::wstring json;
    std::wstring apiUrl = std::wstring(L"https://") + kApiHost + kApiPath;

    if (!HttpGet(apiUrl, &json, nullptr, nullptr))
    {
        if (!bSilent)
            Finish(hMain, UPDATE_FAILED, L"Could not reach GitHub to check for updates.\r\nPlease check your internet connection.");
        else
            PostMessageW(hMain, WM_UPDATE_STATUS, (WPARAM)-1, 0);
        g_busy = false;
        return;
    }

    std::wstring tag;
    if (!JsonFindString(json, L"tag_name", tag) || tag.empty())
    {
        if (!bSilent) Finish(hMain, UPDATE_FAILED, L"Could not read the latest release information from GitHub.");
        else PostMessageW(hMain, WM_UPDATE_STATUS, (WPARAM)-1, 0);
        g_busy = false;
        return;
    }

    std::wstring current(L"" UPD_WIDEN(APP_VERSION_STRING));
    if (CompareVersions(tag, current) <= 0)
    {
        if (!bSilent)
        {
            std::wstring m = L"JustNotepad is up to date.\r\n\r\nInstalled version: " + current;
            Finish(hMain, UPDATE_UP_TO_DATE, m);
        }
        else PostMessageW(hMain, WM_UPDATE_STATUS, (WPARAM)-1, 0);
        g_busy = false;
        return;
    }

    // Find the .zip asset.
    std::wstring zipUrl;
    size_t pos = 0;
    for (;;)
    {
        std::wstring candidate;
        size_t end = 0;
        if (!JsonFindString(json, L"browser_download_url", candidate, pos, &end)) break;
        pos = end;
        if (candidate.size() > 4 && _wcsicmp(candidate.c_str() + candidate.size() - 4, L".zip") == 0)
        {
            zipUrl = candidate;
            break;
        }
    }

    // The published SHA-256 of that zip (a ".sha256" asset of the same release).
    std::wstring shaUrl;
    pos = 0;
    if (!zipUrl.empty())
    for (;;)
    {
        std::wstring candidate;
        size_t end = 0;
        if (!JsonFindString(json, L"browser_download_url", candidate, pos, &end)) break;
        pos = end;
        // The zip's own checksum: "<zip url>.sha256" (a release also has the installer's).
        if (_wcsicmp(candidate.c_str(), (zipUrl + L".sha256").c_str()) == 0)
        {
            shaUrl = candidate;
            break;
        }
    }

    if (zipUrl.empty())
    {
        Finish(hMain, UPDATE_FAILED,
               L"Version " + tag + L" is available, but no downloadable package was found.\r\n"
               L"Please download it manually from GitHub.");
        g_busy = false;
        return;
    }

    std::wstring dir = GetUpdateTempDir();
    std::wstring zipPath = dir + L"\\update.zip";

    if (!HttpGet(zipUrl, nullptr, &zipPath, hMain))
    {
        Finish(hMain, UPDATE_FAILED, L"Downloading the update failed. Please try again later.");
        g_busy = false;
        return;
    }

    // Never install something that does not match what was published.
    std::wstring shaText, expected, actual;
    if (shaUrl.empty() || !HttpGet(shaUrl, &shaText, nullptr, nullptr) ||
        (expected = ParseSha256(shaText)).empty() || !Sha256File(zipPath, actual) || actual != expected)
    {
        DeleteFileW(zipPath.c_str());
        Finish(hMain, UPDATE_FAILED,
               shaUrl.empty() ? L"Version " + tag + L" has no published checksum, so it was not installed.\r\n"
                                L"Please download it manually from GitHub."
                              : L"The downloaded update did not match its published checksum and was deleted.\r\n"
                                L"Nothing was changed. Please try again later.");
        g_busy = false;
        return;
    }

    PostMessageW(hMain, WM_UPDATE_STATUS, (WPARAM)-1, 0);
    PostMessageW(hMain, WM_UPDATE_READY, 0, (LPARAM)DupString(zipPath));
    // g_busy stays true until the main thread finishes handling WM_UPDATE_READY.
}

void Updater_CheckForUpdates(HWND hMain, bool bSilent)
{
    bool expected = false;
    if (!g_busy.compare_exchange_strong(expected, true)) return; // already running
    std::thread(UpdateWorker, hMain, bSilent).detach();
}

// ---------------------------------------------------------------------------
// Applying the update
// ---------------------------------------------------------------------------
static bool IsDirWritable(const std::wstring& dir)
{
    std::wstring probe = dir + L"\\.jn_write_test";
    HANDLE h = CreateFileW(probe.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
                           FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, NULL);
    if (h == INVALID_HANDLE_VALUE) return false;
    CloseHandle(h);
    return true;
}

bool Updater_LaunchApplier(const wchar_t* pszZipPath, HWND hMain)
{
    wchar_t exePath[MAX_PATH] = { 0 };
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    std::wstring exe = exePath;
    std::wstring dest = exe;
    size_t slash = dest.find_last_of(L'\\');
    if (slash != std::wstring::npos) dest.erase(slash);

    std::wstring dir = GetUpdateTempDir();
    std::wstring script = dir + L"\\apply_update.ps1";

    // The applier waits for this process to exit, unpacks the archive over the
    // installation folder and restarts the application.
    std::wstring ps =
        L"param([int]$ProcId,[string]$Zip,[string]$Dest,[string]$Exe)\r\n"
        L"$ErrorActionPreference='Stop'\r\n"
        L"try { Wait-Process -Id $ProcId -Timeout 120 -ErrorAction SilentlyContinue } catch {}\r\n"
        L"Start-Sleep -Milliseconds 500\r\n"
        L"$extract = Join-Path (Split-Path $Zip) 'extract'\r\n"
        L"if (Test-Path $extract) { Remove-Item -Recurse -Force $extract }\r\n"
        L"try {\r\n"
        L"  Expand-Archive -LiteralPath $Zip -DestinationPath $extract -Force\r\n"
        L"  $src = $extract\r\n"
        L"  $items = @(Get-ChildItem -LiteralPath $extract)\r\n"
        L"  if ($items.Count -eq 1 -and $items[0].PSIsContainer) { $src = $items[0].FullName }\r\n"
        L"  Copy-Item -Path (Join-Path $src '*') -Destination $Dest -Recurse -Force\r\n"
        L"} catch {\r\n"
        L"  [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')\r\n"
        L"  [System.Windows.Forms.MessageBox]::Show(\"The update could not be applied:`r`n\" + $_.Exception.Message, 'JustNotepad')\r\n"
        L"}\r\n"
        L"Remove-Item -Recurse -Force $extract -ErrorAction SilentlyContinue\r\n"
        L"Remove-Item -Force $Zip -ErrorAction SilentlyContinue\r\n"
        L"Start-Process -FilePath $Exe\r\n";

    HANDLE h = CreateFileW(script.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
                           FILE_ATTRIBUTE_NORMAL, NULL);
    if (h == INVALID_HANDLE_VALUE) return false;
    // UTF-8 with BOM so PowerShell reads it correctly.
    int need = WideCharToMultiByte(CP_UTF8, 0, ps.c_str(), (int)ps.size(), NULL, 0, NULL, NULL);
    std::vector<char> utf8(need);
    WideCharToMultiByte(CP_UTF8, 0, ps.c_str(), (int)ps.size(), utf8.data(), need, NULL, NULL);
    const unsigned char bom[3] = { 0xEF, 0xBB, 0xBF };
    DWORD written = 0;
    WriteFile(h, bom, 3, &written, NULL);
    WriteFile(h, utf8.data(), (DWORD)utf8.size(), &written, NULL);
    CloseHandle(h);

    std::wstring args =
        L"-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File \"" + script +
        L"\" -ProcId " + std::to_wstring(GetCurrentProcessId()) +
        L" -Zip \"" + pszZipPath + L"\" -Dest \"" + dest + L"\" -Exe \"" + exe + L"\"";

    const wchar_t* verb = IsDirWritable(dest) ? L"open" : L"runas";

    SHELLEXECUTEINFOW sei = { 0 };
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOASYNC;
    sei.hwnd = hMain;
    sei.lpVerb = verb;
    sei.lpFile = L"powershell.exe";
    sei.lpParameters = args.c_str();
    sei.nShow = SW_HIDE;

    bool ok = ShellExecuteExW(&sei) ? true : false;
    if (ok && sei.hProcess) CloseHandle(sei.hProcess);
    g_busy = false;
    return ok;
}
