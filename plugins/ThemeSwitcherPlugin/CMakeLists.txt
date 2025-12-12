add_library(ThemeSwitcherPlugin SHARED ThemeSwitcherPlugin.cpp)
target_compile_definitions(ThemeSwitcherPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
set_target_properties(ThemeSwitcherPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET ThemeSwitcherPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:ThemeSwitcherPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
