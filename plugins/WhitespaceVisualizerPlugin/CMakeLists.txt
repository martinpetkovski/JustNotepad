add_library(WhitespaceVisualizerPlugin SHARED WhitespaceVisualizerPlugin.cpp WhitespaceVisualizerPlugin.def)
target_compile_definitions(WhitespaceVisualizerPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(WhitespaceVisualizerPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET WhitespaceVisualizerPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:WhitespaceVisualizerPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
