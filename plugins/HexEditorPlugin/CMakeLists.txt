add_library(HexEditorPlugin SHARED HexEditorPlugin.cpp HexEditorPlugin.def)
target_compile_definitions(HexEditorPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(HexEditorPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET HexEditorPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:HexEditorPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
