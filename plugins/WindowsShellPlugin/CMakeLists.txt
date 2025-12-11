add_library(WindowsShellPlugin SHARED
    WindowsShellPlugin.cpp
)
target_compile_definitions(WindowsShellPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(WindowsShellPlugin PRIVATE user32 advapi32 shell32 shlwapi)
set_target_properties(WindowsShellPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET WindowsShellPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:WindowsShellPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
