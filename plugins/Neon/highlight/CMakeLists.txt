### export CMAKE_PREFIX_PATH=/usr/local/Qt-6.5.3/

### Build and install highlight (lib, cli and gui).
### This cmake build file is NOT mandatory, you can still use the old makefile
### if preferred.
### The script tries to compile the GUI using Qt6, with Qt5 5.15 as fallback.

cmake_minimum_required(VERSION 3.10)

cmake_policy(SET CMP0167 NEW)

project(highlight-all)

option(BUILD_GUI "Build and install the Qt gui" TRUE)

# Specify C++ standard
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

# Compile options
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -O2 -DNDEBUG -D_FILE_OFFSET_BITS=64")

# Directories
set(CORE_DIR src/core)
set(CLI_DIR src/cli)
set(GUI_QT_DIR src/gui-qt)
set(INCLUDE_DIR src/include)

# Include Boost
find_package(Boost REQUIRED)

# Lua detection
find_package(PkgConfig REQUIRED)
pkg_check_modules(LUA lua5.4)
if (NOT LUA_FOUND)
    pkg_check_modules(LUA lua)
endif()

find_package(Lua REQUIRED)

# Include directories
include_directories(${Boost_INCLUDE_DIRS} ${INCLUDE_DIR} ${LUA_INCLUDE_DIRS} )

# Set installation directories
set(PREFIX /usr)
set(DATA_DIR ${PREFIX}/share)

set(BIN_DIR ${PREFIX}/bin)
set(LIB_DIR ${PREFIX}/lib)
set(MAN_DIR ${DATA_DIR}/man)
set(DOC_DIR ${DATA_DIR}/doc)

set(HIGHLIGHT_DATA_DIR ${DATA_DIR}/highlight/)
set(HIGHLIGHT_DOC_DIR ${DOC_DIR}/highlight/)
set(CONF_DIR /etc/highlight/)

set(BASH_COMP_DIR ${DATA_DIR}/bash-completion/completions)
set(FISH_COMP_DIR ${DATA_DIR}/fish/vendor_completions.d)
set(ZSH_COMP_DIR ${DATA_DIR}/zsh/site-functions)

set(DESKTOP_APPS_DIR ${DATA_DIR}/applications)
set(DESKTOP_ICONS_DIR ${DATA_DIR}/icons/hicolor/256x256/apps)
set(GUI_FILES_DIR ${HIGHLIGHT_DATA_DIR}/gui_files)

# Source files
set(CORE_OBJECTS
    ${CORE_DIR}/stylecolour.cpp
    ${CORE_DIR}/stringtools.cpp
    ${CORE_DIR}/xhtmlgenerator.cpp
    ${CORE_DIR}/latexgenerator.cpp
    ${CORE_DIR}/texgenerator.cpp
    ${CORE_DIR}/rtfgenerator.cpp
    ${CORE_DIR}/htmlgenerator.cpp
    ${CORE_DIR}/ansigenerator.cpp
    ${CORE_DIR}/svggenerator.cpp
    ${CORE_DIR}/codegenerator.cpp
    ${CORE_DIR}/xterm256generator.cpp
    ${CORE_DIR}/pangogenerator.cpp
    ${CORE_DIR}/bbcodegenerator.cpp
    ${CORE_DIR}/odtgenerator.cpp
    ${CORE_DIR}/syntaxreader.cpp
    ${CORE_DIR}/elementstyle.cpp
    ${CORE_DIR}/themereader.cpp
    ${CORE_DIR}/keystore.cpp
    ${CORE_DIR}/lspclient.cpp
    ${CORE_DIR}/datadir.cpp
    ${CORE_DIR}/preformatter.cpp
    ${CORE_DIR}/platform_fs.cpp
    ${CORE_DIR}/astyle/ASStreamIterator.cpp
    ${CORE_DIR}/astyle/ASResource.cpp
    ${CORE_DIR}/astyle/ASFormatter.cpp
    ${CORE_DIR}/astyle/ASBeautifier.cpp
    ${CORE_DIR}/astyle/ASEnhancer.cpp
    ${CORE_DIR}/Diluculum/InternalUtils.cpp
    ${CORE_DIR}/Diluculum/LuaExceptions.cpp
    ${CORE_DIR}/Diluculum/LuaFunction.cpp
    ${CORE_DIR}/Diluculum/LuaState.cpp
    ${CORE_DIR}/Diluculum/LuaUserData.cpp
    ${CORE_DIR}/Diluculum/LuaUtils.cpp
    ${CORE_DIR}/Diluculum/LuaValue.cpp
    ${CORE_DIR}/Diluculum/LuaVariable.cpp
    ${CORE_DIR}/Diluculum/LuaWrappers.cpp
)

set(CLI_OBJECTS
    ${CLI_DIR}/arg_parser.cc
    ${CLI_DIR}/cmdlineoptions.cpp
    ${CLI_DIR}/help.cpp
    ${CLI_DIR}/main.cpp
)

set(GUI_OBJECTS
    ${GUI_QT_DIR}/main.cpp
    ${GUI_QT_DIR}/mainwindow.cpp
    ${GUI_QT_DIR}/io_report.cpp
    ${GUI_QT_DIR}/showtextfile.cpp
    ${GUI_QT_DIR}/syntax_chooser.cpp
)

# Library
add_library(highlight-lib STATIC ${CORE_OBJECTS})

# CLI executable
add_executable(highlight ${CLI_OBJECTS})
target_link_libraries(highlight highlight-lib ${LUA_LIBRARIES} ${Boost_LIBRARIES} dl)
set_target_properties(highlight PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)

# Include directories
target_include_directories(highlight-lib PRIVATE ${INCLUDE_DIR} ${Boost_INCLUDE_DIRS})
target_include_directories(highlight PRIVATE ${INCLUDE_DIR} ${Boost_INCLUDE_DIRS})

# Define custom directories
target_compile_definitions(highlight-lib PRIVATE HL_CONFIG_DIR="${CONF_DIR}" HL_DATA_DIR="${HIGHLIGHT_DATA_DIR}" HL_DOC_DIR="${HIGHLIGHT_DOC_DIR}")
target_compile_definitions(highlight PRIVATE HL_CONFIG_DIR="${CONF_DIR}" HL_DATA_DIR="${HIGHLIGHT_DATA_DIR}" HL_DOC_DIR="${HIGHLIGHT_DOC_DIR}")

# Install locations
install(DIRECTORY DESTINATION ${HIGHLIGHT_DATA_DIR})
install(DIRECTORY DESTINATION ${CONF_DIR})

if(BUILD_GUI)
    install(DIRECTORY DESTINATION ${GUI_FILES_DIR})
    install(DIRECTORY DESTINATION ${GUI_FILES_DIR}/ext)
    install(DIRECTORY DESTINATION ${GUI_FILES_DIR}/l10n)
endif()

install(FILES filetypes.conf DESTINATION ${CONF_DIR})
install(DIRECTORY langDefs DESTINATION ${HIGHLIGHT_DATA_DIR})
install(DIRECTORY themes DESTINATION ${HIGHLIGHT_DATA_DIR})
install(DIRECTORY plugins DESTINATION ${HIGHLIGHT_DATA_DIR})
install(DIRECTORY extras DESTINATION ${HIGHLIGHT_DOC_DIR})
install(FILES AUTHORS README.adoc README_LSP_CLIENT.adoc README_PLUGINS.adoc ChangeLog.adoc COPYING INSTALL DESTINATION ${HIGHLIGHT_DOC_DIR})
install(FILES man/highlight.1 DESTINATION ${MAN_DIR}/man1)
install(FILES man/filetypes.conf.5 DESTINATION ${MAN_DIR}/man5)
install(
    FILES sh-completion/highlight.bash
    RENAME _highlight
    DESTINATION ${BASH_COMP_DIR}
)
install(FILES sh-completion/highlight.fish DESTINATION ${FISH_COMP_DIR})
install(
    FILES sh-completion/highlight.zsh
    RENAME _highlight
    DESTINATION ${ZSH_COMP_DIR}
)

if(BUILD_GUI)
    install(FILES highlight.desktop DESTINATION ${DESKTOP_APPS_DIR} COMPONENT gui)
    install(FILES src/gui-qt/highlight.png DESTINATION ${DESKTOP_ICONS_DIR} COMPONENT gui)
    install(FILES gui_files/ext/fileopenfilter.conf DESTINATION ${GUI_FILES_DIR}/ext COMPONENT gui)
    install(DIRECTORY gui_files/l10n/ DESTINATION ${GUI_FILES_DIR}/l10n FILES_MATCHING PATTERN "*.qm" )
endif()

# Installation of binaries
install(TARGETS highlight DESTINATION ${BIN_DIR} )

# Messages for the user
add_custom_target(post-install
    COMMAND ${CMAKE_COMMAND} -E echo "Done."
    COMMAND ${CMAKE_COMMAND} -E echo "Type highlight --help or man highlight for instructions."
    COMMAND ${CMAKE_COMMAND} -E echo "Take a look at ${HIGHLIGHT_DOC_DIR}/extras for scripts, SWIG and TCL bindings."
    COMMAND ${CMAKE_COMMAND} -E echo "Execute 'make install-gui' to install the highlight GUI."
    COMMAND ${CMAKE_COMMAND} -E echo "Do not hesitate to report problems. Unknown bugs are hard to fix."
)

# Uninstall target
add_custom_target(uninstall
    COMMAND ${CMAKE_COMMAND} -E remove_directory ${HIGHLIGHT_DATA_DIR}
    COMMAND ${CMAKE_COMMAND} -E remove_directory ${HIGHLIGHT_DOC_DIR}
    COMMAND ${CMAKE_COMMAND} -E remove_directory ${CONF_DIR}
    COMMAND ${CMAKE_COMMAND} -E remove ${MAN_DIR}/man1/highlight.1.gz
    COMMAND ${CMAKE_COMMAND} -E remove ${MAN_DIR}/man5/filetypes.conf.5.gz
    COMMAND ${CMAKE_COMMAND} -E remove ${BASH_COMP_DIR}/highlight
    COMMAND ${CMAKE_COMMAND} -E remove ${FISH_COMP_DIR}/highlight.fish
    COMMAND ${CMAKE_COMMAND} -E remove ${ZSH_COMP_DIR}/_highlight
    COMMAND ${CMAKE_COMMAND} -E remove ${BIN_DIR}/highlight
    COMMAND ${CMAKE_COMMAND} -E remove ${BIN_DIR}/highlight-gui
    COMMAND ${CMAKE_COMMAND} -E remove ${DESKTOP_APPS_DIR}/highlight.desktop
    COMMAND ${CMAKE_COMMAND} -E remove ${DESKTOP_ICONS_DIR}/highlight.png
    COMMAND ${CMAKE_COMMAND} -E echo "Done."
)

# qt gui
if(BUILD_GUI)
    set(CMAKE_AUTOMOC ON)
    set(CMAKE_AUTOUIC ON)
    set(CMAKE_AUTORCC ON)

    INCLUDE_DIRECTORIES( ${CMAKE_CURRENT_BINARY_DIR} )

    # Qt version detection and setup
    find_package(Qt6 COMPONENTS Core Gui Widgets)

    if (NOT Qt6_FOUND)
        find_package(Qt5 5.15 COMPONENTS Core Gui Widgets)
    endif()

    # Add the sources
    set(SOURCES
        ${GUI_QT_DIR}/main.cpp
        ${GUI_QT_DIR}/mainwindow.cpp
        ${GUI_QT_DIR}/io_report.cpp
        ${GUI_QT_DIR}/showtextfile.cpp
        ${GUI_QT_DIR}/syntax_chooser.cpp
    )

    # Add the headers
    set(HEADERS
        ${GUI_QT_DIR}/mainwindow.h
        ${GUI_QT_DIR}/precomp.h
        ${GUI_QT_DIR}/io_report.h
        ${GUI_QT_DIR}/showtextfile.h
        ${GUI_QT_DIR}/syntax_chooser.h
    )

    # Add the UI forms
    set(FORMS
        ${GUI_QT_DIR}/mainwindow.ui
        ${GUI_QT_DIR}/io_report.ui
        ${GUI_QT_DIR}/showtextfile.ui
        ${GUI_QT_DIR}/syntax_chooser.ui
    )

    qt_wrap_ui(UIS_HDRS
        ${GUI_QT_DIR}/mainwindow.ui
        ${GUI_QT_DIR}/io_report.ui
        ${GUI_QT_DIR}/showtextfile.ui
        ${GUI_QT_DIR}/syntax_chooser.ui
    )

    # Add the resources
    set(RESOURCES
        ${GUI_QT_DIR}/highlight-gui.qrc
    )

    # Add translations
    set(TRANSLATIONS
        ${GUI_QT_DIR}/highlight_de_DE.ts
        ${GUI_QT_DIR}/highlight_es_ES.ts
        ${GUI_QT_DIR}/highlight_cs_CZ.ts
        ${GUI_QT_DIR}/highlight_zh_CN.ts
        ${GUI_QT_DIR}/highlight_it_IT.ts
        ${GUI_QT_DIR}/highlight_fr_FR.ts
        ${GUI_QT_DIR}/highlight_bg_BG.ts
        ${GUI_QT_DIR}/highlight_ja_JP.ts
    )

    qt_add_resources(GUI_QRC_SOURCES ${GUI_QT_DIR}/highlight-gui.qrc)

    # Create the executable
    add_executable(highlight-gui ${SOURCES} ${HEADERS} ${FORMS} ${RESOURCES} ${UIS_HDRS})

    target_link_libraries(highlight-gui Qt::Core Qt::Gui Qt::Widgets highlight-lib ${LUA_LIBRARIES} ${Boost_LIBRARIES})

    set_target_properties(highlight-gui PROPERTIES AUTOMOC TRUE)
    set_target_properties(highlight-gui PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)

    target_include_directories(highlight-gui PRIVATE ${INCLUDE_DIR} ${LUA_INCLUDE_DIRS} ${Boost_INCLUDE_DIRS})

    target_compile_definitions(highlight-gui PRIVATE HL_CONFIG_DIR="${CONF_DIR}" HL_DATA_DIR="${HIGHLIGHT_DATA_DIR}" HL_DOC_DIR="${HIGHLIGHT_DOC_DIR}"  -DQT)

    install(TARGETS highlight-gui DESTINATION ${BIN_DIR} OPTIONAL COMPONENT gui)
endif()

# Platform-specific configurations
if (WIN32)
    #set(DESTDIR ../..)
    #include_directories(F:/libs/lua_5.4.0/ ../../include F:/libs/boost_1_82_0)
    #add_definitions(-DWIN32)
    #remove_definitions(-DNOCRYPT)
    #set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} /STATIC")
    #target_link_libraries(highlight-gui -LF:/git/highlight/src -lhighlight)
    #if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    #    target_link_libraries(highlight-gui -LF:/libs/lua_5.4.0/x64 -llua)
    #else()
    #    target_link_libraries(highlight-gui -LF:/libs/lua_5.4.0/x32 -llua)
    #endif()
    #set_target_properties(highlight-gui PROPERTIES RESOURCE highlight-gui.rc)
    #add_custom_command(TARGET highlight-gui POST_BUILD COMMAND F:/upx393w/upx.exe --best --force F:/git/highlight/highlight-gui.exe)
elseif (APPLE)
    set(CMAKE_C_COMPILER clang)
    set(CMAKE_CXX_COMPILER clang++)
    set(CMAKE_OSX_DEPLOYMENT_TARGET 14.0)
    set(CMAKE_MACOSX_BUNDLE ON)
    set_target_properties(highlight-gui PROPERTIES MACOSX_BUNDLE_ICON_FILE highlight.icns)
    set_source_files_properties(${CMAKE_CURRENT_SOURCE_DIR}/highlight.icns PROPERTIES MACOSX_PACKAGE_LOCATION "Resources")
endif()
