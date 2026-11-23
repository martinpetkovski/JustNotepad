#ifndef UNICODE
#define UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <algorithm>
#include <mutex>

#include "vcs.h"

using jnp::u8;

std::wstring dirOf(const std::wstring& path) {
    size_t i = path.find_last_of(L"\\/");
    return i == std::wstring::npos ? L"" : path.substr(0, i);
}

std::wstring nameOf(const std::wstring& path) {
    size_t i = path.find_last_of(L"\\/");
    return i == std::wstring::npos ? path : path.substr(i + 1);
}

RunResult runProcess(const std::wstring& cmdline, const std::wstring& cwd, unsigned long timeoutMs,
                     bool mergeStderr, size_t maxOut, const std::string& input) {
    RunResult r;
    SECURITY_ATTRIBUTES sa = {};
    sa.nLength = sizeof(sa);
    sa.bInheritHandle = TRUE;
    HANDLE rd = nullptr, wr = nullptr;
    if (!CreatePipe(&rd, &wr, &sa, 1 << 16)) return r;
    SetHandleInformation(rd, HANDLE_FLAG_INHERIT, 0);
    HANDLE nul = CreateFileW(L"NUL", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, &sa,
                             OPEN_EXISTING, 0, nullptr);

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESTDHANDLES;
    HANDLE inRd = nullptr, inWr = nullptr;
    if (!input.empty() && CreatePipe(&inRd, &inWr, &sa, (DWORD)input.size() + 1024))
        SetHandleInformation(inWr, HANDLE_FLAG_INHERIT, 0);
    si.hStdInput = inRd ? inRd : nul;
    si.hStdOutput = wr;
    si.hStdError = mergeStderr ? wr : nul;
    PROCESS_INFORMATION pi = {};
    std::vector<wchar_t> cmd(cmdline.begin(), cmdline.end());
    cmd.push_back(0);
    BOOL ok = CreateProcessW(nullptr, cmd.data(), nullptr, nullptr, TRUE, CREATE_NO_WINDOW, nullptr,
                             cwd.empty() ? nullptr : cwd.c_str(), &si, &pi);
    CloseHandle(wr);
    if (inRd) {
        CloseHandle(inRd);
        if (ok) {
            DWORD put = 0;
            WriteFile(inWr, input.data(), (DWORD)input.size(), &put, nullptr);
        }
        CloseHandle(inWr);  // end of input
    }
    if (nul != INVALID_HANDLE_VALUE) CloseHandle(nul);
    if (!ok) {
        CloseHandle(rd);
        return r;
    }
    r.started = true;

    const DWORD start = GetTickCount();
    std::vector<char> buf(1 << 16);
    auto readSome = [&](DWORD avail) -> bool {
        DWORD got = 0;
        DWORD want = std::min<DWORD>(avail, (DWORD)buf.size());
        if (!ReadFile(rd, buf.data(), want, &got, nullptr) || got == 0) return false;
        if (r.out.size() + got > maxOut) {
            r.truncated = true;
            return false;
        }
        r.out.append(buf.data(), got);
        return true;
    };
    for (;;) {
        DWORD avail = 0;
        if (!PeekNamedPipe(rd, nullptr, 0, nullptr, &avail, nullptr)) break;  // all writers gone
        if (avail) {
            if (!readSome(avail)) break;
            continue;
        }
        if (WaitForSingleObject(pi.hProcess, 15) == WAIT_OBJECT_0) {
            // Exited: collect whatever is still in the pipe.
            while (PeekNamedPipe(rd, nullptr, 0, nullptr, &avail, nullptr) && avail)
                if (!readSome(avail)) break;
            break;
        }
        if (GetTickCount() - start > timeoutMs) {
            r.timedOut = true;
            break;
        }
    }
    if (r.timedOut || r.truncated) TerminateProcess(pi.hProcess, 1);
    WaitForSingleObject(pi.hProcess, 2000);
    DWORD code = 1;
    GetExitCodeProcess(pi.hProcess, &code);
    r.exitCode = code;
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    CloseHandle(rd);
    return r;
}

namespace {

std::mutex g_connMutex;
VcsConnection g_conn;

std::wstring quoteArg(const std::wstring& s) {
    if (!s.empty() && s.find_first_of(L" \t\"") == std::wstring::npos) return s;
    std::wstring q = L"\"";
    for (wchar_t c : s) q += c == L'"' ? L'\'' : c;
    return q + L"\"";
}

bool isFile(const std::wstring& p) {
    DWORD a = GetFileAttributesW(p.c_str());
    return a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY);
}

std::wstring env(const wchar_t* name) {
    wchar_t b[MAX_PATH];
    DWORD n = GetEnvironmentVariableW(name, b, MAX_PATH);
    return n && n < MAX_PATH ? std::wstring(b) : L"";
}

std::wstring regString(HKEY root, const wchar_t* key, const wchar_t* value) {
    wchar_t b[MAX_PATH];
    DWORD size = sizeof(b);
    for (DWORD view : {RRF_SUBKEY_WOW6464KEY, RRF_SUBKEY_WOW6432KEY}) {
        size = sizeof(b);
        if (RegGetValueW(root, key, value, RRF_RT_REG_SZ | view, nullptr, b, &size) == ERROR_SUCCESS && b[0])
            return b;
    }
    return L"";
}

std::wstring onPath(const wchar_t* exe) {
    wchar_t b[MAX_PATH];
    DWORD n = SearchPathW(nullptr, exe, nullptr, MAX_PATH, b, nullptr);
    return n && n < MAX_PATH ? std::wstring(b) : L"";
}

std::wstring firstExisting(const std::vector<std::wstring>& candidates) {
    for (const std::wstring& c : candidates)
        if (!c.empty() && isFile(c)) return c;
    return L"";
}

std::wstring detectGit() {
    std::wstring p = onPath(L"git.exe");
    if (!p.empty()) return p;
    std::vector<std::wstring> c;
    for (HKEY root : {HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER}) {
        std::wstring dir = regString(root, L"SOFTWARE\\GitForWindows", L"InstallPath");
        if (!dir.empty()) c.push_back(dir + L"\\cmd\\git.exe");
    }
    for (const wchar_t* pf : {L"ProgramFiles", L"ProgramW6432", L"ProgramFiles(x86)"}) {
        std::wstring d = env(pf);
        if (!d.empty()) c.push_back(d + L"\\Git\\cmd\\git.exe");
    }
    std::wstring local = env(L"LOCALAPPDATA");
    if (!local.empty()) c.push_back(local + L"\\Programs\\Git\\cmd\\git.exe");
    std::wstring home = env(L"USERPROFILE");
    if (!home.empty()) c.push_back(home + L"\\scoop\\shims\\git.exe");
    return firstExisting(c);
}

std::wstring detectP4() {
    std::wstring p = onPath(L"p4.exe");
    if (!p.empty()) return p;
    std::vector<std::wstring> c;
    for (HKEY root : {HKEY_LOCAL_MACHINE, HKEY_CURRENT_USER}) {
        std::wstring dir = regString(root, L"SOFTWARE\\Perforce\\Environment", L"P4INSTROOT");
        if (!dir.empty()) c.push_back(dir + L"\\p4.exe");
    }
    for (const wchar_t* pf : {L"ProgramFiles", L"ProgramW6432", L"ProgramFiles(x86)"}) {
        std::wstring d = env(pf);
        if (!d.empty()) c.push_back(d + L"\\Perforce\\p4.exe");
    }
    return firstExisting(c);
}

}  // namespace

std::wstring findGit() {
    static std::once_flag once;
    static std::wstring path;
    std::call_once(once, [] { path = detectGit(); });
    return path;
}

std::wstring findP4() {
    static std::once_flag once;
    static std::wstring path;
    std::call_once(once, [] { path = detectP4(); });
    return path;
}

void setVcsConnection(const VcsConnection& c) {
    std::lock_guard<std::mutex> lock(g_connMutex);
    g_conn = c;
}

VcsConnection vcsConnection() {
    std::lock_guard<std::mutex> lock(g_connMutex);
    return g_conn;
}

std::wstring gitCmd() {
    VcsConnection c = vcsConnection();
    std::wstring exe = c.gitExe.empty() ? findGit() : c.gitExe;
    return exe.empty() ? L"git" : quoteArg(exe);
}

std::wstring p4Cmd() {
    VcsConnection c = vcsConnection();
    std::wstring exe = c.p4Exe.empty() ? findP4() : c.p4Exe;
    std::wstring s = exe.empty() ? L"p4" : quoteArg(exe);
    if (!c.p4Port.empty()) s += L" -p " + quoteArg(c.p4Port);
    if (!c.p4User.empty()) s += L" -u " + quoteArg(c.p4User);
    if (!c.p4Client.empty()) s += L" -c " + quoteArg(c.p4Client);
    return s;
}

namespace {

bool exists(const std::wstring& p) { return GetFileAttributesW(p.c_str()) != INVALID_FILE_ATTRIBUTES; }

std::wstring widen(const std::string& s) {
    std::wstring w;
    if (s.empty()) return w;
    int k = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
    if (k <= 0) return w;
    w.resize((size_t)k);
    MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), &w[0], k);
    return w;
}

std::string trim(std::string s) {
    while (!s.empty() && (s.back() == '\n' || s.back() == '\r' || s.back() == ' ')) s.pop_back();
    return s;
}

// Value of a "... key value" line in `p4 -ztag` output, or "".
std::string ztag(const std::string& out, const char* key) {
    std::string needle = std::string("... ") + key + " ";
    size_t pos = 0;
    while (pos < out.size()) {
        size_t eol = out.find('\n', pos);
        if (eol == std::string::npos) eol = out.size();
        if (out.compare(pos, needle.size(), needle) == 0)
            return trim(out.substr(pos + needle.size(), eol - pos - needle.size()));
        pos = eol + 1;
    }
    return "";
}

bool insideGitRepo(std::wstring dir) {
    for (;;) {
        if (exists(dir + L"\\.git")) return true;
        size_t i = dir.find_last_of(L"\\/");
        if (i == std::wstring::npos || i < 2) return false;
        dir.resize(i);
    }
}

}  // namespace

std::vector<std::string> ztagAll(const std::string& out, const char* key) {
    std::vector<std::string> v;
    std::string needle = std::string("... ") + key + " ";
    size_t pos = 0;
    while (pos < out.size()) {
        size_t eol = out.find('\n', pos);
        if (eol == std::string::npos) eol = out.size();
        if (out.compare(pos, needle.size(), needle) == 0)
            v.push_back(trim(out.substr(pos + needle.size(), eol - pos - needle.size())));
        pos = eol + 1;
    }
    return v;
}

VcsInfo queryPerforce(const std::wstring& path, const std::wstring& dir, size_t maxBase);

VcsInfo queryVcs(const std::wstring& path, size_t maxBase) {
    VcsInfo v;
    const std::wstring dir = dirOf(path), name = nameOf(path);
    if (dir.empty() || name.empty()) return v;

    const VcsKind only = vcsConnection().only;
    VcsInfo git;
    if (only != VcsKind::Perforce && insideGitRepo(dir)) {
        git.kind = VcsKind::Git;
        RunResult br = runProcess(gitCmd() + L" rev-parse --abbrev-ref HEAD", dir, 5000, false, 4096);
        if (!br.started) {
            git.toolMissing = true;
        } else {
            git.label = widen(trim(br.out));
            if (git.label.empty()) git.label = L"(no commits)";
            // HEAD:./name is the committed version of this file, relative to its folder.
            RunResult show = runProcess(gitCmd() + L" show \"HEAD:./" + name + L"\"", dir, 15000, false, maxBase);
            if (show.started && !show.timedOut && show.exitCode == 0) {
                git.tracked = true;
                if (!show.truncated) {
                    git.haveBase = true;
                    git.base.assign(show.out.begin(), show.out.end());
                }
            }
        }
        // A file Git does not track may still be a Perforce file: workspaces
        // often sit inside (or contain) a Git repository. Ask Perforce too,
        // unless Git was chosen explicitly.
        if (git.tracked || only == VcsKind::Git) return git;
    }
    if (only == VcsKind::Git) return v;
    VcsInfo p4 = queryPerforce(path, dir, maxBase);
    if (p4.kind == VcsKind::Perforce || git.kind == VcsKind::None) return p4;
    return git;  // untracked in Git and unknown to Perforce: a new Git file
}

// Perforce part of queryVcs.
VcsInfo queryPerforce(const std::wstring& path, const std::wstring& dir, size_t maxBase) {
    VcsInfo v;
    // Perforce: ask the server about the file. This fails fast when p4.exe is
    // not installed and is bounded by the timeout when the server is away.
    RunResult st = runProcess(p4Cmd() + L" -ztag fstat \"" + path + L"\"", dir, 6000, false, 1 << 16);
    if (!st.started || st.timedOut || ztag(st.out, "depotFile").empty()) return v;
    v.kind = VcsKind::Perforce;
    std::string have = ztag(st.out, "haveRev"), head = ztag(st.out, "headRev"), action = ztag(st.out, "action");
    std::string depot = ztag(st.out, "depotFile"), headAction = ztag(st.out, "headAction");
    std::string pathU8;
    {
        int k = WideCharToMultiByte(CP_UTF8, 0, path.c_str(), (int)path.size(), nullptr, 0, nullptr, nullptr);
        pathU8.resize((size_t)std::max(k, 0));
        if (k > 0) WideCharToMultiByte(CP_UTF8, 0, path.c_str(), (int)path.size(), &pathU8[0], k, nullptr, nullptr);
    }
    v.label = L"#" + widen(have.empty() ? "none" : have) + L"/" + widen(head.empty() ? "?" : head);
    v.action = widen(action);
    v.opened = !action.empty();
    v.change = widen(ztag(st.out, "change"));

    // The revision to compare with, best first. #have is what the workspace
    // was synced to. Files opened as move/add, branch or re-add after a delete
    // have no #have of their own, but the depot still has their history.
    std::vector<std::string> candidates;
    if (!have.empty() && have != "none" && have != "0") candidates.push_back(depot + "#" + have);
    if (action == "move/add") {
        std::string moved = ztag(st.out, "movedFile");
        if (!moved.empty()) {
            std::string movedRev = ztag(st.out, "movedRev");
            if (!movedRev.empty()) candidates.push_back(moved + "#" + movedRev);
            candidates.push_back(moved + "#have");
            candidates.push_back(moved + "#head");
        }
    }
    if (action == "branch" || action == "integrate" || action == "add" || action == "move/add") {
        // Integrated / branched files: compare with the source of the pending resolve.
        RunResult rs = runProcess(p4Cmd() + L" -ztag resolved \"" + path + L"\"", dir, 6000, false, 1 << 16);
        std::string from = ztag(rs.out, "fromFile"), endFrom = ztag(rs.out, "endFromRev");
        if (!from.empty()) candidates.push_back(from + (endFrom.empty() ? "" : "#" + endFrom));
    }
    if (!head.empty() && headAction.find("delete") == std::string::npos) candidates.push_back(depot + "#" + head);

    // Last resorts: let the server resolve the local path itself (the same
    // lookup `p4 diff` and P4V use), then the newest depot revision.
    candidates.push_back(pathU8 + "#have");
    if (head.empty()) {
        RunResult fl = runProcess(p4Cmd() + L" -ztag files \"" + widen(depot) + L"\"", dir, 6000, false, 1 << 16);
        std::string rev = ztag(fl.out, "rev"), act = ztag(fl.out, "action");
        if (!rev.empty() && act.find("delete") == std::string::npos) candidates.push_back(depot + "#" + rev);
    }
    // Drop repeats so a failing revision is not fetched twice.
    std::vector<std::string> uniq;
    for (const std::string& c : candidates)
        if (std::find(uniq.begin(), uniq.end(), c) == uniq.end()) uniq.push_back(c);
    candidates.swap(uniq);

    v.tracked = true;  // settled after the fetch below
    std::string lastErr;
    for (const std::string& rev : candidates) {
        // Depot syntax, so the answer does not depend on the workspace mapping.
        RunResult pr = runProcess(p4Cmd() + L" print -q \"" + widen(rev) + L"\"", dir, 20000, false, maxBase);
        if (pr.started && !pr.timedOut && !pr.truncated && pr.exitCode == 0) {
            v.haveBase = true;
            v.baseFrom = widen(rev);
            v.base.assign(pr.out.begin(), pr.out.end());
            break;
        }
        if (pr.truncated) {
            lastErr = "The base revision is too large to compare.";
            break;
        }
        // Re-run with stderr to learn why (rare path, so the extra call is fine).
        RunResult why = runProcess(p4Cmd() + L" print -q \"" + widen(rev) + L"\"", dir, 20000, true, 4096);
        lastErr = pr.timedOut ? "The Perforce server did not answer." : trim(why.out);
        if (lastErr.size() > 300) lastErr.resize(300);
    }
    // A file opened for add with nothing in the depot really is new.
    v.tracked = v.haveBase || !have.empty() || !head.empty() || action != "add";
    if (!v.haveBase && v.tracked) {
        // Show what the server said, so a missing base can be diagnosed from the diff pane.
        std::string f = "fstat: depotFile " + depot + ", haveRev " + (have.empty() ? "-" : have) + ", headRev " +
                        (head.empty() ? "-" : head) + ", headAction " + (headAction.empty() ? "-" : headAction) +
                        ", action " + (action.empty() ? "-" : action);
        v.baseError = widen((lastErr.empty() ? std::string("The base revision could not be fetched.") : lastErr) + "\n" + f);
    }
    return v;
}
