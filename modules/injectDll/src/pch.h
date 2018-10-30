#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files
#include <windows.h>

#ifdef WIN32
#pragma warning(push)

// warning C4512: assignment operator could not be generated.
#pragma warning(disable: 4512)
// warning C4706: assignment within conditional expression.
#pragma warning(disable: 4706)
// warning C4996: 'name': This function or variable may be unsafe.
#pragma warning(disable: 4996)

#endif // #ifdef WIN32

#ifdef WIN32
#pragma warning(pop)
#endif // #ifdef WIN32



