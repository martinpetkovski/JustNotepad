#ifndef UNICODE
#define UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <algorithm>
#include <cstdlib>
#include <cstdio>
#include <cctype>
#include <cwchar>
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

// ---------------------------------------------------------------- history

namespace {

// A `p4 -ztag` value that may run over several lines (descriptions).
std::string ztagLong(const std::string& out, const std::string& key) {
    const std::string needle = "... " + key + " ";
    size_t p = 0;
    while (p < out.size()) {
        size_t eol = out.find('\n', p);
        if (eol == std::string::npos) eol = out.size();
        if (out.compare(p, needle.size(), needle) == 0) {
            std::string v = out.substr(p + needle.size(), eol - p - needle.size());
            // Following lines up to the next "... " belong to it.
            size_t q = eol + 1;
            while (q < out.size() && out.compare(q, 4, "... ") != 0) {
                size_t e2 = out.find('\n', q);
                if (e2 == std::string::npos) e2 = out.size();
                v += "\n" + out.substr(q, e2 - q);
                q = e2 + 1;
            }
            while (!v.empty() && (v.back() == '\n' || v.back() == '\r' || v.back() == ' ')) v.pop_back();
            for (size_t k = 0; (k = v.find('\r', k)) != std::string::npos;) v.erase(k, 1);
            return v;
        }
        p = eol + 1;
    }
    return "";
}

std::string trimNl(std::string s) {
    while (!s.empty() && (s.back() == '\n' || s.back() == '\r' || s.back() == ' ')) s.pop_back();
    size_t a = 0;
    while (a < s.size() && (s[a] == '\n' || s[a] == '\r')) a++;
    return s.substr(a);
}

}  // namespace

bool fileHistory(const std::wstring& path, VcsKind kind, std::vector<Revision>& out, std::wstring& err, int max) {
    out.clear();
    const std::wstring dir = dirOf(path), name = nameOf(path);
    if (kind == VcsKind::Perforce) {
        wchar_t m[16];
        swprintf(m, 16, L"%d", max);
        RunResult r = runProcess(p4Cmd() + L" -ztag filelog -l -m " + m + L" \"" + path + L"\"", dir, 30000, true, 16u << 20);
        if (!r.started) return err = L"p4.exe could not be started.", false;
        if (r.timedOut) return err = L"The Perforce server did not answer.", false;
        const std::string depot = ztag(r.out, "depotFile");
        if (r.exitCode != 0 || depot.empty()) return err = widen(trimNl(r.out)), false;
        for (int i = 0; i < max; i++) {
            const std::string k = std::to_string(i);
            std::string rev = ztag(r.out, ("rev" + k).c_str());
            if (rev.empty()) break;
            Revision v;
            v.rev = widen(rev);
            v.change = widen(ztag(r.out, ("change" + k).c_str()));
            v.user = widen(ztag(r.out, ("user" + k).c_str()));
            v.action = widen(ztag(r.out, ("action" + k).c_str()));
            v.time = atoll(ztag(r.out, ("time" + k).c_str()).c_str());
            v.desc = widen(ztagLong(r.out, "desc" + k));
            v.spec = widen(depot) + L"#" + v.rev;
            out.push_back(std::move(v));
        }
        return true;
    }
    if (kind == VcsKind::Git) {
        // Records split by 0x1E, fields by 0x1F; --name-status gives the file's
        // path in each commit (it may have been renamed since).
        wchar_t m[16];
        swprintf(m, 16, L"%d", max);
        RunResult r = runProcess(gitCmd() + L" log --follow -n " + m +
                                     L" --format=%x1e%H%x1f%h%x1f%an%x1f%at%x1f%B%x1f --name-status -- \"" + name + L"\"",
                                 dir, 30000, false, 32u << 20);
        if (!r.started) return err = L"git.exe could not be started.", false;
        if (r.timedOut) return err = L"git did not answer.", false;
        if (r.exitCode != 0) return err = L"git log failed.", false;
        size_t p = 0;
        while ((p = r.out.find('\x1e', p)) != std::string::npos) {
            size_t e = r.out.find('\x1e', p + 1);
            std::string rec = r.out.substr(p + 1, (e == std::string::npos ? r.out.size() : e) - p - 1);
            p = e == std::string::npos ? r.out.size() : e;
            std::vector<std::string> f;
            size_t a = 0;
            for (int k = 0; k < 5; k++) {
                size_t b = rec.find('\x1f', a);
                if (b == std::string::npos) b = rec.size();
                f.push_back(rec.substr(a, b - a));
                a = b + 1;
            }
            if (f.size() < 5 || f[0].empty()) continue;
            // After the message: "M\tpath" or "R100\told\tnew".
            std::string status = a < rec.size() ? trimNl(rec.substr(a)) : "";
            size_t nl = status.find_last_of('\n');
            if (nl != std::string::npos) status = status.substr(nl + 1);
            std::string action, file;
            size_t t1 = status.find('\t');
            if (t1 != std::string::npos) {
                action = status.substr(0, t1);
                file = status.substr(status.find_last_of('\t') + 1);
            }
            Revision v;
            v.change = widen(f[0]);
            v.rev = widen(f[1]);
            v.user = widen(f[2]);
            v.time = atoll(f[3].c_str());
            v.desc = widen(trimNl(f[4]));
            v.action = widen(action.empty() ? "" : action.substr(0, 1));
            v.spec = widen(f[0]) + L":" + widen(file);  // path from the repository root
            out.push_back(std::move(v));
        }
        if (out.empty()) return err = L"This file has no commits yet.", false;
        return true;
    }
    err = L"This file is not under Git or Perforce.";
    return false;
}

bool fetchRevision(const std::wstring& path, VcsKind kind, const std::wstring& spec, size_t maxBytes,
                   std::vector<jnp::u8>& out, std::wstring& err) {
    const std::wstring dir = dirOf(path);
    RunResult r = kind == VcsKind::Perforce ? runProcess(p4Cmd() + L" print -q \"" + spec + L"\"", dir, 30000, false, maxBytes)
                                            : runProcess(gitCmd() + L" show \"" + spec + L"\"", dir, 30000, false, maxBytes);
    if (!r.started) return err = L"The version control tool could not be started.", false;
    if (r.timedOut) return err = L"The server did not answer.", false;
    if (r.truncated) return err = L"That revision is too large to compare.", false;
    if (r.exitCode != 0) {
        RunResult why = kind == VcsKind::Perforce ? runProcess(p4Cmd() + L" print -q \"" + spec + L"\"", dir, 30000, true, 4096)
                                                  : runProcess(gitCmd() + L" show \"" + spec + L"\"", dir, 30000, true, 4096);
        return err = widen(trimNl(why.out)), false;
    }
    out.assign(r.out.begin(), r.out.end());
    return true;
}

// ---------------------------------------------------------------- blame

namespace {
long long parseP4Date(const std::string& d) {  // "2026/11/20" -> seconds (UTC midnight)
    int y = 0, m = 0, dd = 0;
    if (sscanf_s(d.c_str(), "%d/%d/%d", &y, &m, &dd) != 3) return 0;
    // Days from civil (Howard Hinnant), good for any Gregorian date.
    y -= m <= 2;
    const long long era = (y >= 0 ? y : y - 399) / 400;
    const unsigned yoe = (unsigned)(y - era * 400);
    const unsigned doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + dd - 1;
    const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return (era * 146097 + (long long)doe - 719468) * 86400;
}
}  // namespace

bool blameFile(const std::wstring& path, VcsKind kind, Blame& out, std::wstring& err) {
    out = Blame();
    const std::wstring dir = dirOf(path), name = nameOf(path);
    if (kind == VcsKind::Git) {
        RunResult r = runProcess(gitCmd() + L" blame --porcelain HEAD -- \"" + name + L"\"", dir, 60000, false, 256u << 20);
        if (!r.started || r.timedOut || r.exitCode != 0 || r.truncated) return err = L"git blame failed.", false;
        std::vector<std::pair<std::string, int>> seen;  // hash -> info index
        int cur = -1;
        size_t p = 0;
        const std::string& o = r.out;
        while (p < o.size()) {
            size_t e = o.find('\n', p);
            if (e == std::string::npos) e = o.size();
            std::string l = o.substr(p, e - p);
            p = e + 1;
            if (!l.empty() && l[0] == '\t') {  // the line itself
                out.line.push_back(cur);
                out.text += l.substr(1) + "\n";
                continue;
            }
            if (l.size() >= 41 && l[40] == ' ' && isxdigit((unsigned char)l[0])) {  // "<hash> <orig> <final> [n]"
                std::string hash = l.substr(0, 40);
                cur = -1;
                for (auto& s : seen)
                    if (s.first == hash) cur = s.second;
                if (cur < 0) {
                    BlameInfo bi;
                    bi.change = widen(hash.substr(0, 8));
                    out.infos.push_back(bi);
                    cur = (int)out.infos.size() - 1;
                    seen.push_back({hash, cur});
                }
                continue;
            }
            if (cur < 0) continue;
            BlameInfo& bi = out.infos[(size_t)cur];
            if (l.rfind("author ", 0) == 0) bi.who = widen(l.substr(7));
            else if (l.rfind("author-time ", 0) == 0) bi.time = atoll(l.c_str() + 12);
            else if (l.rfind("summary ", 0) == 0) bi.summary = widen(l.substr(8));
        }
        return true;
    }
    if (kind == VcsKind::Perforce) {
        // Change number per line (-c), no header (-q). Users and dates come
        // from filelog: works with every server, unlike annotate -u.
        RunResult r = runProcess(p4Cmd() + L" annotate -c -q \"" + path + L"#have\"", dir, 60000, false, 256u << 20);
        if (!r.started) return err = L"p4.exe could not be started.", false;
        if (r.timedOut) return err = L"p4 annotate timed out.", false;
        if (r.truncated) return err = L"The file is too large to annotate.", false;
        if (r.exitCode != 0) {
            RunResult why = runProcess(p4Cmd() + L" annotate -c -q \"" + path + L"#have\"", dir, 30000, true, 4096);
            return err = widen(trimNl(why.out)), false;
        }
        std::vector<std::pair<std::string, int>> seen;
        size_t p = 0;
        const std::string& o = r.out;
        while (p < o.size()) {
            size_t e = o.find('\n', p);
            if (e == std::string::npos) e = o.size();
            std::string l = o.substr(p, e - p);
            p = e + 1;
            if (!l.empty() && l.back() == '\r') l.pop_back();
            // "1234: text"
            size_t colon = l.find(':');
            if (colon == std::string::npos || colon == 0) continue;
            bool digits = true;
            for (size_t k = 0; k < colon; k++) digits = digits && isdigit((unsigned char)l[k]);
            if (!digits) continue;
            std::string change = l.substr(0, colon);
            std::string text = colon + 2 <= l.size() ? l.substr(colon + 2) : "";
            int idx = -1;
            for (auto& s : seen)
                if (s.first == change) idx = s.second;
            if (idx < 0) {
                BlameInfo bi;
                bi.change = widen(change);
                out.infos.push_back(bi);
                idx = (int)out.infos.size() - 1;
                seen.push_back({change, idx});
            }
            out.line.push_back(idx);
            out.text += text + "\n";
        }
        if (out.line.empty()) return err = L"p4 annotate returned nothing.", false;
        // Who submitted each change, and when.
        std::vector<Revision> hist;
        std::wstring herr;
        if (fileHistory(path, kind, hist, herr, 5000)) {
            for (BlameInfo& bi : out.infos)
                for (const Revision& rv : hist)
                    if (rv.change == bi.change) {
                        bi.who = rv.user;
                        bi.time = rv.time;
                        size_t nl = rv.desc.find_first_of(L"\r\n");
                        bi.summary = nl == std::wstring::npos ? rv.desc : rv.desc.substr(0, nl);
                        break;
                    }
        }
        // Changes that reached the file through an integration are not in
        // its own filelog: ask for those directly.
        int asked = 0;
        for (BlameInfo& bi : out.infos) {
            if (!bi.who.empty() || ++asked > 40) continue;
            RunResult d = runProcess(p4Cmd() + L" -ztag describe -s " + bi.change, dir, 15000, false, 1u << 20);
            bi.who = widen(ztag(d.out, "user"));
            bi.time = atoll(ztag(d.out, "time").c_str());
        }
        return true;
    }
    return err = L"Not under version control.", false;
}
