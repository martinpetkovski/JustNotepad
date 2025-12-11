add_library(GitPlugin SHARED GitPlugin.cpp GitPlugin.rc)
target_compile_definitions(GitPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(GitPlugin PRIVATE user32 kernel32 shlwapi)

set_target_properties(GitPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET GitPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:GitPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)