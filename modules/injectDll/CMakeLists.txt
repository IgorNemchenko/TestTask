cmake_minimum_required(VERSION 3.9.0)

project(injectDll)

# header files
set(HEADERS
	include/targetver.h
)

# source files
set(SOURCES
	src/dllmain.cpp
)

# misc headers/sources
set(MISC
	src/pch.h
	src/pch.cpp
)

source_group("" FILES ${MISC})
source_group("Header Files" FILES ${HEADERS})
source_group("Source Files" FILES ${SOURCES})

#
# Target.
#

add_library(${PROJECT_NAME} SHARED
	${HEADERS}
	${SOURCES}
	${MISC}
)

target_include_directories(${PROJECT_NAME}
	PRIVATE
		include
		src
		${CMAKE_SOURCE_DIR}/external

)

target_link_libraries(${PROJECT_NAME}
	debug ${CMAKE_SOURCE_DIR}/external/Blackbone/lib/BlackBone_d.lib
    optimized ${CMAKE_SOURCE_DIR}/external/Blackbone/lib/BlackBone.lib
    )
	
if (MSVC)
	# Use link time code generation in the Release builds. 
	set_target_properties(${PROJECT_NAME} PROPERTIES STATIC_LIBRARY_FLAGS_RELEASE "/LTCG")
endif()

#
# Compiler options.
#

if (MSVC)
	add_definitions(/wd4714)
	add_definitions(-D_SCL_SECURE_NO_WARNINGS)

	# Ignore warning for missing PDBs.
	set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} /ignore:4099")
	SET(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /std:c++latest" )
	
	# Disable SAFESEH check.
	if(NOT (CMAKE_CL_64))
		message(STATUS "Disabling SAFESEH")
		SET(CMAKE_EXE_LINKER_FLAGS_DEBUG "${CMAKE_EXE_LINKER_FLAGS_DEBUG} /SAFESEH:NO")
		SET(CMAKE_EXE_LINKER_FLAGS_RELEASE "${CMAKE_EXE_LINKER_FLAGS_RELEASE} /SAFESEH:NO")
	endif()
endif()

#
# Precompiled headers.
#

#set(COTIRE_MINIMUM_NUMBER_OF_TARGET_SOURCES 0)
set_target_properties(${PROJECT_NAME} PROPERTIES COTIRE_CXX_PREFIX_HEADER_INIT "src/pch.h")
set_target_properties(${PROJECT_NAME} PROPERTIES COTIRE_UNITY_LINK_LIBRARIES_INIT "COPY")
cotire(${PROJECT_NAME})