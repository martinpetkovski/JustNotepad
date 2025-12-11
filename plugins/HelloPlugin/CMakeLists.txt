add_library(HelloPlugin SHARED HelloPlugin.cpp HelloPlugin.def)
target_compile_definitions(HelloPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(HelloPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET HelloPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:HelloPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
