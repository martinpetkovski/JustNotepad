add_library(ClangFormatPlugin SHARED ClangFormatPlugin.cpp ClangFormatPlugin.rc)
target_compile_definitions(ClangFormatPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(ClangFormatPlugin PRIVATE user32 kernel32 shlwapi)

set_target_properties(ClangFormatPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET ClangFormatPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:ClangFormatPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
