
#ifndef __XINDOWSUTIL_H__
#define __XINDOWSUTIL_H__

#ifndef XINDOWS_PUBLIC
#ifdef XINDOWS_EXPORT
#define XINDOWS_PUBLIC	__declspec(dllexport)
#else
#define XINDOWS_PUBLIC	__declspec(dllimport)
#endif
#endif

#include <Windows.h>
#include <atlbase.h>

#pragma warning ( disable : 4200 )  // nonstandard extension used : zero-sized array in struct/union
#pragma warning ( disable : 4251 )  // needs to have dll-interface to be used by clients of class ***
#pragma warning ( disable : 4996 )  // function or variable may be unsafe

#define NO_COPY(cls) \
    cls(const cls&); \
    cls& operator=(const cls&)

#define DYNCAST(Dest_type, Source_Value) \
    (static_cast<Dest_type*>(Source_Value))

#include "debug/Debug.h"

#include "global/Win32Error.h"
#include "global/Math.h"
#include "global/Void.h"

#include "memory/MemUtil.h"

#include "mt/GlobalLock.h"
#include "mt/CriticalSection.h"
#include "mt/ReadWriteSection.h"

#include "string/String.h"
#include "string/Buffer.h"
#include "string/BufferedString.h"
#include "string/StringUtil.h"

#include "number/NumConv.h"

#include "com/Unknown.h"
#include "com/Variant.h"
#include "com/TypeInfoNav.h"

#include "collection/Array.h"
#include "collection/Array2.h"
#include "collection/Assoc.h"
#include "collection/AtomTable.h"
#include "collection/HtPvPv.h"

#include "geom/Size.h"
#include "geom/Point.h"
#include "geom/Rect.h"
#include "geom/Region.h"
#include "geom/RegionStack.h"

#include "stream/StreamDef.h"
#include "stream/Substream.h"
#include "stream/BufferStream.h"
#include "stream/ROStmOnBuffer.h"
#include "stream/FatStg.h"
#include "stream/DataStream.h"

#include "encode/Mime64.h"

#include "dll/DynamicLibrary.h"

#include "international/Codepage.h"

#endif //__XINDOWSUTIL_H__