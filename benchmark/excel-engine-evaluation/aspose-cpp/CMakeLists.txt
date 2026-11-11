cmake_minimum_required(VERSION 3.16)
project(aspose_cpp_smoke LANGUAGES CXX)

if(NOT DEFINED ASPOSE_SDK_ROOT)
    message(FATAL_ERROR "Pass -DASPOSE_SDK_ROOT=/path/to/Aspose.Cells")
endif()

find_package(Aspose.Cells REQUIRED CONFIG PATHS "${ASPOSE_SDK_ROOT}" NO_DEFAULT_PATH)
add_executable(aspose_cpp_smoke src/aspose_cpp_smoke.cpp)
target_compile_features(aspose_cpp_smoke PRIVATE cxx_std_17)
target_link_libraries(aspose_cpp_smoke PRIVATE Aspose.Cells)
set_target_properties(aspose_cpp_smoke PROPERTIES
    BUILD_RPATH "$ORIGIN/../lib"
    INSTALL_RPATH "$ORIGIN/../lib")
install(TARGETS aspose_cpp_smoke RUNTIME DESTINATION bin)
