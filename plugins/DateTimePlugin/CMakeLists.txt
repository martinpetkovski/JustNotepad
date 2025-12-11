add_library(DateTimePlugin SHARED DateTimePlugin.cpp DateTimePlugin.def)
target_compile_definitions(DateTimePlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(DateTimePlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET DateTimePlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:DateTimePlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
