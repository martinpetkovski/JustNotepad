add_library(DataViewerPlugin SHARED 
    DataViewerPlugin.cpp 
    DataViewerPlugin.def
    DataViewerPlugin.rc
)
target_compile_definitions(DataViewerPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(DataViewerPlugin PRIVATE user32 comctl32 xmllite shlwapi)
set_target_properties(DataViewerPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET DataViewerPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:DataViewerPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
