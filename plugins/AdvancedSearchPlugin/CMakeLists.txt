add_library(AdvancedSearchPlugin SHARED
    AdvancedSearchPlugin.cpp
    AdvancedSearchPlugin.rc
)
target_compile_definitions(AdvancedSearchPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(AdvancedSearchPlugin PRIVATE user32 comctl32)
set_target_properties(AdvancedSearchPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

add_custom_command(TARGET AdvancedSearchPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:AdvancedSearchPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
)
