add_library(JsonViewerPlugin SHARED 
    JsonViewerPlugin.cpp 
    JsonViewerPlugin.def
    JsonViewerPlugin.rc
)
target_compile_definitions(JsonViewerPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(JsonViewerPlugin PRIVATE user32 comctl32 xmllite shlwapi)
set_target_properties(JsonViewerPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET JsonViewerPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:JsonViewerPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
