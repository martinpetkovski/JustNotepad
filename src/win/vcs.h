#pragma once
// Version control support. JustNotepad does not link any VCS library: it runs
// the git.exe / p4.exe already installed on the machine, so there is nothing to
// configure and the user's existing credentials are used.
#include <string>
#include <vector>

#include "../core/types.h"

struct RunResult {
    bool started = false;   // false: the program could not be launched (not installed / not on PATH)
    bool timedOut = false;
    bool truncated = false; // output exceeded maxOut and the process was stopped
    unsigned long exitCode = 0;
    std::string out;
};

// Runs a command line hidden, capturing stdout (and stderr when asked).
// `input`, if given, is written to the program's stdin (keep it under 64 KB).
RunResult runProcess(const std::wstring& cmdline, const std::wstring& cwd, unsigned long timeoutMs,
                     bool mergeStderr, size_t maxOut, const std::string& input = std::string());

enum class VcsKind { None, Git, Perforce };

// How to reach version control, set in Source Control > Connect. Empty fields
// fall back to what git / p4 already use (P4CONFIG, P4PORT, `p4 set`, ...).
struct VcsConnection {
    VcsKind only = VcsKind::None;  // None: detect; Git / Perforce: use only that system
    std::wstring gitExe;           // empty: detected (see findGit / findP4)
    std::wstring p4Exe;            // empty: detected
    std::wstring p4Port, p4User, p4Client;
    bool autoCheckout = true;      // perforce: run `p4 edit` as soon as a file is edited
};

// Thread-safe; every later query and command uses the new settings.
void setVcsConnection(const VcsConnection& c);
VcsConnection vcsConnection();

// Locates the installed tools: PATH first, then the installers' registry
// entries, then the usual install folders. Empty when not found. The result
// is cached after the first call.
std::wstring findGit();
std::wstring findP4();

// Command-line prefixes for the configured tools, e.g. `p4 -p host:1666 -u me -c ws`.
std::wstring gitCmd();
std::wstring p4Cmd();

struct VcsInfo {
    VcsKind kind = VcsKind::None;
    std::wstring label;       // git: branch; perforce: "#have/head"
    std::wstring action;      // perforce: edit / add / ..., empty when not opened
    bool toolMissing = false; // repository found but the command-line tool is not available
    bool tracked = false;     // the file exists in the base revision
    bool opened = false;      // perforce: checked out
    bool haveBase = false;
    std::vector<jnp::u8> base;  // raw bytes of the base revision (git HEAD / perforce #have)
    std::wstring change;      // perforce: changelist the file is open in ("default" or a number)
    std::wstring baseFrom;    // perforce: the depot revision the base came from, e.g. //depot/a.txt#3
    std::wstring baseError;   // why there is no base although the file has history
};

// Works out which system manages `path` and fetches the base revision of the
// file. Slow (starts processes, may talk to a server): call it off the UI thread.
VcsInfo queryVcs(const std::wstring& path, size_t maxBase);

// Every value of `key` in `p4 -ztag` output, in order.
std::vector<std::string> ztagAll(const std::string& out, const char* key);

std::wstring dirOf(const std::wstring& path);
std::wstring nameOf(const std::wstring& path);
