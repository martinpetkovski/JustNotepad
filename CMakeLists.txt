cmake_minimum_required(VERSION 3.20)
project(JustNotepad LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_EXPORT_COMPILE_COMMANDS ON)

# Windows-specific definitions
add_definitions(-DUNICODE -D_UNICODE)

if(MSVC)
    add_compile_options(/utf-8)
endif()

add_executable(JustNotepad WIN32
    src/main.cpp
    src/resources.rc
    src/TextHelpers.cpp
    src/PluginManager.cpp
    src/Updater.cpp
)

# Link Windows libs
if(MSVC)
    target_link_libraries(JustNotepad PRIVATE user32 gdi32 comdlg32 ole32 comctl32 shlwapi shell32 dwmapi uxtheme winhttp delayimp)
    # Enable Common Controls v6 (needed for SysLink and visual styles)
    target_link_options(JustNotepad PRIVATE "/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'")

    # Startup cost: every statically imported DLL is touched (and scanned by
    # Defender) before the first frame is drawn. These are only needed once the
    # user actually opens a dialog / checks for updates, so load them on demand.
    target_link_options(JustNotepad PRIVATE
        /DELAYLOAD:comdlg32.dll
        /DELAYLOAD:ole32.dll
        /DELAYLOAD:dwmapi.dll
        /DELAYLOAD:uxtheme.dll
        /DELAYLOAD:winhttp.dll
    )

    # Smaller image = less to page in on a cold start.
    target_link_options(JustNotepad PRIVATE $<$<CONFIG:Release>:/OPT:REF /OPT:ICF>)
else()
    target_link_libraries(JustNotepad PRIVATE user32 gdi32 comdlg32 ole32 comctl32 shlwapi shell32 dwmapi uxtheme winhttp)
endif()

# Tests
enable_testing()

add_executable(UnitTests
    tests/UnitTests.cpp
    src/TextHelpers.cpp
)

if(MSVC)
    target_link_libraries(UnitTests PRIVATE user32)
else()
    target_link_libraries(UnitTests PRIVATE user32)
endif()

add_test(NAME UnitTests COMMAND UnitTests)

add_executable(PluginTests
    tests/PluginTests.cpp
)

if(MSVC)
    target_link_libraries(PluginTests PRIVATE user32)
else()
    target_link_libraries(PluginTests PRIVATE user32)
endif()

add_test(NAME PluginTests COMMAND PluginTests)

add_subdirectory(plugins/DateTimePlugin)
add_subdirectory(plugins/GuidPlugin)
add_subdirectory(plugins/EmojiPlugin)
add_subdirectory(plugins/GitPlugin)
add_subdirectory(plugins/CharacterMapPlugin)
add_subdirectory(plugins/WhitespaceVisualizerPlugin)

# Output directory
set_target_properties(JustNotepad PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}"
    OUTPUT_NAME "Just Notepad"
)
