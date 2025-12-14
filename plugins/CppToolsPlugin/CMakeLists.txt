add_library(CppToolsPlugin SHARED CppToolsPlugin.cpp CppToolsPlugin.def)
target_compile_definitions(CppToolsPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(CppToolsPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET CppToolsPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:CppToolsPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
