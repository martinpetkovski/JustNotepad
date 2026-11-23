# JustNotepad

A plain text editor for Windows that starts instantly and stays instant at any file size.
One small exe, no installer, no plugins.

## Build

Needs Visual Studio 2022 (or the Build Tools) with the C++ workload. Nothing else.

```powershell
.\build.ps1          # release build -> build\JustNotepad.exe
.\build.ps1 -Run     # build and start
.\build.ps1 -Test    # build and run the core tests + benchmark
```

## Why it is fast

- **Opening copies nothing.** Large files are memory-mapped; the document is a
  piece tree that points into the file. Opening 10 GB costs the same as 1 KB.
- **Edits are O(log n).** Typed text goes into an append-only buffer and the
  tree gets a new piece. Nothing is ever shifted or re-read. Undo and redo move
  piece descriptors around.
- **Only the screen is laid out.** Scrolling is anchored to a byte offset, so
  painting, the caret and the scrollbar never need to know how many lines the
  file has.
- **Line numbers arrive in the background.** A worker thread counts newlines
  (SSE2, several GB/s) into a small index; until it finishes you can already
  scroll, edit and search.
- **Nothing loads before the first frame** except the window and the font.
  Dialog and shell libraries are delay-loaded.

## What works today

- Open, edit and save any file: UTF-8 (with or without BOM), UTF-16 LE/BE,
  single-byte ANSI, and binary data (bytes are preserved exactly)
- Unlimited undo/redo, cut/copy/paste, drag selection, word navigation
- Find, replace, replace all (match case, whole word, both directions)
- Go to line, line numbers, word wrap, auto-indent, zoom, font choice
- CRLF/LF detection, drag and drop, per-monitor DPI, settings remembered

## Version control

Nothing to set up: when a file is opened, JustNotepad looks for a Git
repository or asks Perforce about it, in the background, using the `git.exe` or
`p4.exe` already on your PATH (and therefore your existing login).

- The status bar shows the branch (Git) or have/head revision (Perforce) and
  the number of changes against the base revision (Git `HEAD`, Perforce `#have`).
- Changed lines are marked in the gutter: green added, blue modified, red where
  lines were removed.
- **Source Control** menu: stage / check out, commit / submit this file, revert
  the file, refresh. Editing a Perforce file that is not checked out offers to
  check it out.

## Live diff

**F8** splits the window: the base revision on the left, your working text on
the right. The right side is the normal editor, so you edit inside the diff and
it updates as you type. **F7** / **Shift+F7** jump between changes and
**Ctrl+R** puts the base lines back for the change under the caret (one undo
step).

## Releasing

```powershell
.\publish.ps1           # version, build, test, zip, tag, push, GitHub release
.\publish.ps1 -DryRun   # everything up to the zip; nothing committed or pushed
```

Versions are `<major>.<yyMMdd>.<n>`, where n counts releases on the same day.
Needs a clean working tree, an `origin` remote and the GitHub CLI (`gh`).

## Roadmap

1. Tabs
2. Code features: syntax highlighting, multi-cursor, column select, bracket
   matching, comment toggle, regex search on a background thread
3. Padded line alignment in the diff view, staging individual changes

## Source layout

| Path | What |
|---|---|
| `src/core/` | Portable core, no OS headers: piece tree, line index, layout, diff, UTF helpers |
| `src/win/` | Win32 front end: window, painting, input, file I/O, dialogs, version control |
| `tests/` | Differential tests and benchmark for the core |
| `docs/icon.ico` | Application icon |

## Known limits

- The grid is fixed-pitch. Scripts that need shaping (Arabic, Indic) and colour
  emoji do not render properly yet.
- A single line longer than 16 KB is shown as several rows.
- Saving a memory-mapped file (over 32 MB) clears its undo history.
- Case-insensitive search folds ASCII letters only.
- Change markers and the diff view are off for files over 16 MB.
- The diff view lines the two sides up at the middle of the window; it does not
  pad them, so rows above and below a change can sit at different heights.
- Perforce is detected for files the server already knows; brand-new files
  cannot be added from the editor yet.
- UTF-16 files are converted in memory on load, so they do not open instantly
  when very large.
