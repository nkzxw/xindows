
#ifndef __XINDOWSUTIL_MEMORY_MEMUTIL_H__
#define __XINDOWSUTIL_MEMORY_MEMUTIL_H__

XINDOWS_PUBLIC void*   _MemAlloc(ULONG cb);
XINDOWS_PUBLIC void*   _MemAllocClear(ULONG cb);
XINDOWS_PUBLIC HRESULT _MemRealloc(void** ppv, ULONG cb);
XINDOWS_PUBLIC ULONG   _MemGetSize(void* pv);
XINDOWS_PUBLIC void    _MemFree(void* pv);
XINDOWS_PUBLIC HRESULT _MemAllocString(LPCTSTR pchSrc, LPTSTR* ppchDst);
XINDOWS_PUBLIC HRESULT _MemAllocString(ULONG cch, LPCTSTR pchSrc, LPTSTR* ppchDst);
XINDOWS_PUBLIC HRESULT _MemReplaceString(LPCTSTR pchSrc, LPTSTR* ppchDest);

#define MemAlloc(cb)                            _MemAlloc(cb)
#define MemAllocClear(cb)                       _MemAllocClear(cb)
#define MemRealloc(ppv, cb)                     _MemRealloc(ppv, cb)
#define MemGetSize(pv)                          _MemGetSize(pv)
#define MemFree(pv)	                            _MemFree(pv)
#define MemAllocString(pch, ppch)               _MemAllocString(pch, ppch)
#define MemAllocStringBuffer(cch, pch, ppch)    _MemAllocString(cch, pch, ppch)
#define MemReplaceString(pch, ppch)             _MemReplaceString(pch, ppch)
#define MemFreeString(pch)                      _MemFree(pch)

#ifdef _DEBUG
XINDOWS_PUBLIC void __cdecl _MemSetName(void* pv, char* szFmt, ...);
XINDOWS_PUBLIC char*   _MemGetName(void* pv);

#define MemGetName(pv)  _MemGetName(pv)
#define MemSetName(x)   _MemSetName x
#else //_DEBUG
#define MemGetName(pv)  ((char*)0)
#define MemSetName(x)
#endif //!_DEBUG


#define DECLARE_MEMALLOC_NEW_DELETE() \
    inline void* __cdecl operator new(size_t cb)    { return(MemAlloc(cb)); } \
    inline void* __cdecl operator new[](size_t cb)  { return(MemAlloc(cb)); } \
    inline void __cdecl operator delete(void* pv)   { MemFree(pv); }          \
    inline void __cdecl operator delete[](void* pv) { MemFree(pv); }

#define DECLARE_MEMCLEAR_NEW_DELETE() \
    inline void* __cdecl operator new(size_t cb)    { return(MemAllocClear(cb)); } \
    inline void* __cdecl operator new[](size_t cb)  { return(MemAllocClear(cb)); } \
    inline void __cdecl operator delete(void* pv)   { MemFree(pv); }               \
    inline void __cdecl operator delete[](void* pv) { MemFree(pv); }

#define DECLARE_GLOBAL_MEMALLOC_NEW_DELETE \
    inline void* __cdecl operator new(size_t cb)    { return(MemAlloc(cb)); } \
    inline void* __cdecl operator new[](size_t cb)  { return(MemAlloc(cb)); } \
    inline void __cdecl operator delete(void* pv)   { MemFree(pv); }          \
    inline void __cdecl operator delete[](void* pv) { MemFree(pv); }

#endif //__XINDOWSUTIL_MEMORY_MEMUTIL_H__