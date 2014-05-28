
#ifndef __XINDOWS_H__
#define __XINDOWS_H__

#ifndef XINDOWS_PUBLIC
#ifdef XINDOWS_EXPORT
#define XINDOWS_PUBLIC	__declspec(dllexport)
#else
#define XINDOWS_PUBLIC	__declspec(dllimport)
#endif
#endif

#include "core/core.h"
#include "site/site.h"

#endif //__XINDOWS_H__