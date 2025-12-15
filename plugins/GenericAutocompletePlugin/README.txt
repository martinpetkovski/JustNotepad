Generic Autocomplete Plugin (Ctags)
===================================

This plugin provides autocomplete suggestions using Universal Ctags.

Setup:
1. Download Universal Ctags from https://github.com/universal-ctags/ctags-win32/releases
2. Extract the zip file.
3. Place `ctags.exe` in `plugins/GenericAutocompletePlugin/ctags/` folder (create the folder if it doesn't exist).
   OR ensure `ctags.exe` is in your system PATH.

Usage:
- Press Ctrl+Space to see suggestions for the current file.
- The plugin currently supports C++, Python, and JavaScript explicitly, and defaults to C++ for others.
- Suggestions are shown in a popup message box (basic implementation).

Building:
- The plugin is built as part of the JustNotepad solution.
