add_library(MarkdownToolsPlugin SHARED MarkdownToolsPlugin.cpp MarkdownToolsPlugin.def)
target_compile_definitions(MarkdownToolsPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(MarkdownToolsPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET MarkdownToolsPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:MarkdownToolsPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
