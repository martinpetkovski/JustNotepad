add_library(MacrosPlugin SHARED
    MacrosPlugin.cpp
)
target_compile_definitions(MacrosPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(MacrosPlugin PRIVATE user32 advapi32 shell32 shlwapi)
set_target_properties(MacrosPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET MacrosPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:MacrosPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/macros
    COMMAND ${CMAKE_COMMAND} -E copy_directory ${CMAKE_SOURCE_DIR}/macros $<TARGET_FILE_DIR:JustNotepad>/macros
)
