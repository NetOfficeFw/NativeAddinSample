// pch.h: This is a precompiled header file.
// Files listed below are compiled only once, improving build performance for future builds.
// This also affects IntelliSense performance, including code completion and many code browsing features.
// However, files listed here are ALL re-compiled if any one of them is updated between builds.
// Do not add files here that you will be updating frequently as this negates the performance advantage.

#ifndef PCH_H
#define PCH_H

// add headers that you want to pre-compile here
#include "framework.h"

#pragma warning( disable : 4278 )
#pragma warning( disable : 4146 )
	// import the MSADDNDR.DLL typelib which we need for IDTExtensibility2
	#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4" raw_interfaces_only named_guids
#pragma warning( default : 4146 )
#pragma warning( default : 4278 )

#include "resource.h"
#include "NativeAddinTypeLib_h.h"

#include "gtest/gtest.h"

#endif //PCH_H
