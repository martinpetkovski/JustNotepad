add_library(EmojiPlugin SHARED EmojiPlugin.cpp EmojiPlugin.rc EmojiPlugin.def)
target_compile_definitions(EmojiPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(EmojiPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET EmojiPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:EmojiPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
