cmake_minimum_required(VERSION 3.20)
project(JustNotepad LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Windows-specific definitions
add_definitions(-DUNICODE -D_UNICODE)

add_executable(JustNotepad WIN32
    src/main.cpp
    src/resources.rc
    src/TextHelpers.cpp
)

# Link Windows libs
if(MSVC)
    target_link_libraries(JustNotepad PRIVATE user32 gdi32 comdlg32 ole32 comctl32)
else()
    target_link_libraries(JustNotepad PRIVATE user32 gdi32 comdlg32 ole32 comctl32)
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




# Output directory
set_target_properties(JustNotepad PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}"
)
