add_library(CharacterMapPlugin SHARED CharacterMapPlugin.cpp CharacterMapPlugin.rc CharacterMapPlugin.def)
target_compile_definitions(CharacterMapPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(CharacterMapPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET CharacterMapPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:CharacterMapPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
