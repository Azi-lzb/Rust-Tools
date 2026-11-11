cmake_minimum_required(VERSION 3.16)
project(libxl_linux_smoke LANGUAGES CXX)

if(NOT DEFINED LIBXL_ROOT)
    message(FATAL_ERROR "Set LIBXL_ROOT to the extracted official LibXL SDK directory")
endif()

add_executable(libxl_smoke src/libxl_smoke.cpp)
target_compile_features(libxl_smoke PRIVATE cxx_std_17)
target_include_directories(libxl_smoke PRIVATE "${LIBXL_ROOT}/include_cpp")
target_link_libraries(libxl_smoke PRIVATE "${LIBXL_ROOT}/lib64/libxl.so" stdc++fs)
set_target_properties(libxl_smoke PROPERTIES
    BUILD_RPATH "${LIBXL_ROOT}/lib64"
    INSTALL_RPATH "$ORIGIN/../lib"
)

install(TARGETS libxl_smoke RUNTIME DESTINATION bin)
install(FILES "${LIBXL_ROOT}/lib64/libxl.so" DESTINATION lib)
