add_library(NeonPlugin SHARED NeonPlugin.cpp NeonPlugin.rc NeonPlugin.def)
target_compile_definitions(NeonPlugin PRIVATE JUSTNOTEPAD_PLUGIN_EXPORTS)
target_link_libraries(NeonPlugin PRIVATE shlwapi comctl32)

set_target_properties(NeonPlugin PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/plugins"
)

# Download and extract precompiled highlight
set(HIGHLIGHT_VERSION "4.18")
set(HIGHLIGHT_URL "http://andre-simon.de/zip/highlight-${HIGHLIGHT_VERSION}-x64.zip")
set(HIGHLIGHT_ZIP "${CMAKE_CURRENT_BINARY_DIR}/highlight.zip")
set(HIGHLIGHT_EXTRACT_DIR "${CMAKE_CURRENT_BINARY_DIR}/highlight_extracted")
set(HIGHLIGHT_ROOT "${HIGHLIGHT_EXTRACT_DIR}/highlight-${HIGHLIGHT_VERSION}-x64")

if(NOT EXISTS "${HIGHLIGHT_ROOT}/highlight.exe")
    message(STATUS "Downloading Highlight ${HIGHLIGHT_VERSION} from ${HIGHLIGHT_URL}...")
    file(DOWNLOAD ${HIGHLIGHT_URL} ${HIGHLIGHT_ZIP} SHOW_PROGRESS)
    
    message(STATUS "Extracting Highlight...")
    file(ARCHIVE_EXTRACT INPUT ${HIGHLIGHT_ZIP} DESTINATION ${HIGHLIGHT_EXTRACT_DIR})
endif()

file(GLOB HIGHLIGHT_DLLS "${HIGHLIGHT_ROOT}/*.dll")

add_custom_command(TARGET NeonPlugin POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins
    COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:NeonPlugin> $<TARGET_FILE_DIR:JustNotepad>/plugins/
    
    # Copy downloaded highlight executable and data
    COMMAND ${CMAKE_COMMAND} -E make_directory $<TARGET_FILE_DIR:JustNotepad>/plugins/highlight
    COMMAND ${CMAKE_COMMAND} -E copy "${HIGHLIGHT_ROOT}/highlight.exe" $<TARGET_FILE_DIR:JustNotepad>/plugins/highlight/
    COMMAND ${CMAKE_COMMAND} -E copy_directory "${HIGHLIGHT_ROOT}/langDefs" $<TARGET_FILE_DIR:JustNotepad>/plugins/highlight/langDefs
    COMMAND ${CMAKE_COMMAND} -E copy_directory "${HIGHLIGHT_ROOT}/themes" $<TARGET_FILE_DIR:JustNotepad>/plugins/highlight/themes
    COMMAND ${CMAKE_COMMAND} -E copy_directory "${HIGHLIGHT_ROOT}/plugins" $<TARGET_FILE_DIR:JustNotepad>/plugins/highlight/plugins
    COMMAND ${CMAKE_COMMAND} -E copy "${HIGHLIGHT_ROOT}/filetypes.conf" $<TARGET_FILE_DIR:JustNotepad>/plugins/highlight/
)

foreach(DLL ${HIGHLIGHT_DLLS})
    add_custom_command(TARGET NeonPlugin POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy "${DLL}" $<TARGET_FILE_DIR:JustNotepad>/plugins/highlight/
    )
endforeach()
