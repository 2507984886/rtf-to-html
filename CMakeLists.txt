cmake_minimum_required(VERSION 3.15)
project(rtf2html VERSION 1.0.0 LANGUAGES CXX)

# C++17 标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# 构建类型
if(NOT CMAKE_BUILD_TYPE)
  set(CMAKE_BUILD_TYPE Release)
endif()

# 编译选项
if(MSVC)
  add_compile_options(/W4 /utf-8)
else()
  add_compile_options(-Wall -Wextra -Wpedantic)
endif()

# ============================================================
# 库目标: rtf2html
# ============================================================
add_library(rtf2html
  src/codepage.cpp
  src/tokenizer.cpp
  src/converter.cpp
  src/rtf2html.cpp
)

target_include_directories(rtf2html
  PUBLIC
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
    $<INSTALL_INTERFACE:include>
  PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
)

set_target_properties(rtf2html PROPERTIES
  VERSION ${PROJECT_VERSION}
  SOVERSION 1
  POSITION_INDEPENDENT_CODE ON
)

# ============================================================
# CLI 工具目标: rtf2html_cli
# ============================================================
add_executable(rtf2html_cli
  tools/rtf2html_cli/main.cpp
)

target_link_libraries(rtf2html_cli PRIVATE rtf2html)

set_target_properties(rtf2html_cli PROPERTIES
  OUTPUT_NAME rtf2html
)

# ============================================================
# 测试（可选）
# ============================================================
option(BUILD_TESTS "构建测试" OFF)
if(BUILD_TESTS)
  enable_testing()
  add_subdirectory(tests)
endif()

# ============================================================
# 安装规则
# ============================================================
include(GNUInstallDirs)

install(TARGETS rtf2html rtf2html_cli
  EXPORT rtf2htmlTargets
  LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
  ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
  RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
)

install(DIRECTORY include/ DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})

install(EXPORT rtf2htmlTargets
  FILE rtf2htmlTargets.cmake
  NAMESPACE rtf2html::
  DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/rtf2html
)
