
#ifndef __XINDOWS_TESTXINDOWS_H__
#define __XINDOWS_TESTXINDOWS_H__

#ifdef XINDOWS_EXPORT
#define XINDOWS_PUBLIC	__declspec(dllexport)
#else
#define XINDOWS_PUBLIC	__declspec(dllimport)
#endif

#include "core/testcore.h"
#include "site/testsite.h"

#endif //__XINDOWS_TESTXINDOWS_H__