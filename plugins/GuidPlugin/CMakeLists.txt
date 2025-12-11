add_library(GuidPlugin SHARED GuidPlugin.cpp GuidPlugin.def)
target_compile_definitions(GuidPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(GuidPlugin PRIVATE ole32)
set_target_properties(GuidPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET GuidPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:GuidPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
