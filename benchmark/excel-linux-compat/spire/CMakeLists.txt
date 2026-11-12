cmake_minimum_required(VERSION 3.16)
project(spire_smoke LANGUAGES CXX)

if(NOT DEFINED SPIRE_ROOT)
    message(FATAL_ERROR "Pass -DSPIRE_ROOT=/path/to/Spire.XLS.Cpp.Linux/build/native")
endif()

add_executable(spire_smoke src/spire_smoke.cpp)
target_compile_features(spire_smoke PRIVATE cxx_std_17)
target_include_directories(spire_smoke PRIVATE "${SPIRE_ROOT}/include")
target_link_directories(spire_smoke PRIVATE "${SPIRE_ROOT}/lib/linux")
target_link_libraries(spire_smoke PRIVATE SpireXlsCpp SpireXlsBase dl)
set_target_properties(spire_smoke PROPERTIES
    BUILD_RPATH "$ORIGIN/../lib"
    INSTALL_RPATH "$ORIGIN/../lib")

install(TARGETS spire_smoke RUNTIME DESTINATION bin)
