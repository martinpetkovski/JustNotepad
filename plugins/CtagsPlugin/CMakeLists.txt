add_library(CtagsPlugin SHARED CtagsPlugin.cpp CtagsPlugin.def)
target_compile_definitions(CtagsPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(CtagsPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET CtagsPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:CtagsPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
    COMMAND ${CMAKE_COMMAND} -E copy_directory ${CMAKE_CURRENT_SOURCE_DIR}/ctags $<TARGET_FILE_DIR:JustNotepad>/plugins/ctags
)
