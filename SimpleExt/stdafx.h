// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit


#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

// Custom below
#include <strsafe.h>
#include <atlstr.h> // For Cstring handling
#include <ShlObj.h> // for Shell COM interfaces
#include <new.h>	// for bad alloc
#include <Windows.h> // for windows process control
#include <io.h>
#include <string> // for cstring
#include <algorithm> // for string replace
#include <vector>
#include <exception> // for std::exception
#include <iostream>
#include <sstream> // use for stream string
#include <fstream>
