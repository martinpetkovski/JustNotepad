add_library(PerforcePlugin SHARED PerforcePlugin.cpp PerforcePlugin.rc)
target_compile_definitions(PerforcePlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(PerforcePlugin PRIVATE user32 kernel32 shlwapi)

set_target_properties(PerforcePlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET PerforcePlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:PerforcePlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
