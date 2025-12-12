add_library(Neon SHARED Neon.cpp Neon.rc Neon.def)
target_compile_definitions(Neon PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(Neon PRIVATE shlwapi comctl32)

set_target_properties(Neon PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET Neon POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:Neon> $<TARGET_FILE_DIR:JustNotepad>/plugins/
    COMMAND ${CMAKE_COMMAND} -E copy_directory ${CMAKE_CURRENT_SOURCE_DIR}/highlight $<TARGET_FILE_DIR:JustNotepad>/plugins/highlight
)
