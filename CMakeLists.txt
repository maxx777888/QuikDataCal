cmake_minimum_required (VERSION 3.8)
project ("QuikAS")
set(CMAKE_CXX_STANDARD 17)

add_subdirectory(OpenXLSX)
#add_subdirectory(external/nowide EXCLUDE_FROM_ALL)
# Include the OpenXLSX headers
include_directories(OpenXLSX/headers)

# Define the executable with source files
add_executable (QuikAS "QuikAS.cpp" "QuikAS.h" "stock.h" "stock.cpp" "menu.h" "menu.cpp"    "dataInputXLSX.h" "dataInputXLSX.cpp")
#target_include_directories(QuikAS PRIVATE ${Boost_INCLUDE_DIRS})
# Link with Boost libraries before defining the executable
#target_link_libraries(QuikAS ${Boost_LIBRARIES})
find_package(XLnt REQUIRED) 
target_link_libraries(QuikAS OpenXLSX::OpenXLSX xlnt::xlnt)
 # Use REQUIRED to ensure XLnt is found

#target_link_libraries(QuikAS xlnt::xlnt)
set_target_properties(QuikAS PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/output)