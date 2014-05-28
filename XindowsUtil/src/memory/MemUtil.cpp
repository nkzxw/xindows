
#include "stdafx.h"
#include "MemUtil.h"

#include <atlbase.h>

extern HANDLE g_hProcessHeap;

#ifdef _DEBUG
DWORD   g_dwTlsDbg          = (DWORD)-1;    // 内存调试的Tls索引

size_t  g_cbTotalAllocated  = 0;            // 记录当前请求内存分配总和
size_t  g_cbMaxAllocated    = 0;            // 记录请求内存分配峰值
ULONG   g_cAllocated        = 0;            // 记录请求内存分配次数

CRITICAL_SECTION g_csDebug;

class CDbgLockGlobal
{
public:
    CDbgLockGlobal() { EnterCriticalSection(&g_csDebug); }
    ~CDbgLockGlobal() { LeaveCriticalSection(&g_csDebug); }
};

#define DBG_LOCK_GLOBALS    CDbgLockGlobal glock

// 线程局部对象结构体，辅助实现每个线程的请求内存分配记录
struct DBGTHREADSTATE
{
    DBGTHREADSTATE* ptsNext;
    DBGTHREADSTATE* ptsPrev;

    // Add globals below
    void*           pvRequest;  // 线程最后一次请求分配内存的指针
    size_t          cbRequest;  // 线程最后一次请求分配内存的大小
};
static DBGTHREADSTATE* s_pts = NULL;

//+--------------------------------------------------------------
//
// 每个请求分配内存块的前缀结构体
// 用来跟踪所有请求分配块以及请求分配名称
//
//---------------------------------------------------------------
struct DBGALLOCHDR
{
    DBGALLOCHDR*    pdbgahPrev; // 前一个内存块头
    DBGALLOCHDR*    pdbgahNext; // 后一个内存块头
    DWORD           iAllocated; // 记录是第几次请求分配操作
    DWORD           tid;        // 请求分配线程的ID
    size_t          cbRequest;  // 请求分配大小
    char            szName[64]; // 请求分配块名称
    DWORD           adwGuard[4];// 保护头
};

// 内存跟踪块的根，通过根可以获取所有分配块
DBGALLOCHDR g_dbgahRoot =
{
    &g_dbgahRoot,
    &g_dbgahRoot,
    0,
    (DWORD)-1
};

//+--------------------------------------------------------------
//
// 每个请求分配内存块的后缀结构体
// 使用特定的数据填充用来检测指针是合法
//
//---------------------------------------------------------------
struct DBGALLOCFOOT
{
    DWORD adwGuard[4];
};

const BYTE  FILL_CLEAN      = 0x9D;         // 分配内存后填充内存值
const BYTE  FILL_DEAD       = 0xAD;         // 释放内存前填充内存值
const DWORD FILL_HEAD_GUARD = 0xBDBDBDBD;   // 实际内存头填充保护值
const DWORD FILL_FOOT_GUARD = 0xCDCDCDCD;   // 实际内存尾填充保护值
//+--------------------------------------------------------------
//
//  内存保护区填充字节
//
//---------------------------------------------------------------
DWORD g_adwHeadGuardFill[4] = { FILL_HEAD_GUARD, FILL_HEAD_GUARD, FILL_HEAD_GUARD, FILL_HEAD_GUARD };
DWORD g_adwFootGuardFill[4] = { FILL_FOOT_GUARD, FILL_FOOT_GUARD, FILL_FOOT_GUARD, FILL_FOOT_GUARD };

void    _DbgDllProcessAttach();
void    _DbgDllProcessDetach();
HRESULT _DbgDllThreadAttach();
void    _DbgDllThreadDetach();
void    _DbgDllThreadDetach(DBGTHREADSTATE* pts);

// 确保线程局部对象创建
HRESULT _EnsureThreadState()
{
    if(!TlsGetValue(g_dwTlsDbg))
    {
        return _DbgDllThreadAttach();
    }
    return S_OK;
}

// 返回线程局部对象
DBGTHREADSTATE* _DbgGetThreadState()
{
    return (DBGTHREADSTATE*)TlsGetValue(g_dwTlsDbg);
}

// 调试期间实际分配内存大小=请求分配+分配头+分配尾
size_t _ActualSizeFromRequestSize(size_t cb)
{
    return cb+sizeof(DBGALLOCHDR)+sizeof(DBGALLOCFOOT);
}

// 根据分配块头获取请求分配内存地址
void* _RequestFromActual(DBGALLOCHDR* pdbgah)
{
    return pdbgah+1;
}

// 根据请求分配内存地址获取分配块头
DBGALLOCHDR* _ActualFromRequest(void* pv)
{
    return ((DBGALLOCHDR*)pv)-1;
}

// 根据分配块头获取分配块尾
DBGALLOCFOOT* _FooterFromBlock(DBGALLOCHDR* pdbgah)
{
    return (DBGALLOCFOOT*)(((BYTE*)pdbgah)+sizeof(DBGALLOCHDR)+pdbgah->cbRequest);
}

// 根据请求内存的保护区判断内存块是否合法
BOOL _BlockIsValid(DBGALLOCHDR* pdbgah)
{
    DBGALLOCFOOT* pdbgft;

    Assert(sizeof(pdbgah->adwGuard) == sizeof(g_adwHeadGuardFill));
    if(memcmp(pdbgah->adwGuard, g_adwHeadGuardFill, sizeof(g_adwHeadGuardFill)))
    {
        return FALSE;
    }

    pdbgft = _FooterFromBlock(pdbgah);
    Assert(sizeof(pdbgft->adwGuard) == sizeof(g_adwFootGuardFill));
    if(memcmp(pdbgft->adwGuard, g_adwFootGuardFill, sizeof(g_adwFootGuardFill)))
    {
        return FALSE;
    }

    return TRUE;
}

// Traces分配或者释放内存信息
void _DbgTraceAlloc(TCHAR* szTag, size_t cbAlloc, size_t cbFree, DWORD iAllocated)
{
    Trace((TEXT("%s +%5d -%5d = [%7d]  #=%-4d\n"), szTag, cbAlloc, cbFree, g_cbTotalAllocated, iAllocated));
}

// Traces所有分配块检查内存泄漏
void WINAPI _DbgExTraceMemoryLeaks()
{
    DBGALLOCHDR*    pdbgah;
    char            achBuf[512];

    _EnsureThreadState();

    DBG_LOCK_GLOBALS;

    USES_CONVERSION;

    wsprintfA(achBuf, "---------- Leaked Memory Blocks ----------\n");
    Trace((A2T(achBuf)));
    for(pdbgah=g_dbgahRoot.pdbgahNext; pdbgah!=&g_dbgahRoot; pdbgah=pdbgah->pdbgahNext)
    {
        wsprintfA(achBuf, "p=0x%08x  cb=%-4d #=%-4d TID:0x%x %s\n",
            _RequestFromActual(pdbgah),
            pdbgah->cbRequest,
            pdbgah->iAllocated,
            pdbgah->tid,
            pdbgah->szName);

        Trace((A2T(achBuf)));
    }


    wsprintfA(achBuf, "total size %d, peak size %d\n", g_cbTotalAllocated, g_cbMaxAllocated);
    Trace((A2T(achBuf)));
    wsprintfA(achBuf, "---------- Leaked Memory Blocks End ------\n");
    Trace((A2T(achBuf)));
}

// 初始化内存调试的CriticalSection
// 创建内存调试的Tls索引
void _DbgDllProcessAttach()
{
    ::InitializeCriticalSection(&g_csDebug);

    g_dwTlsDbg = TlsAlloc();
    if(g_dwTlsDbg == (DWORD)(-1))
    {
        AssertSz(0, "TlsAlloc DBG failed!!!");
        return;
    }

    _EnsureThreadState();
}

// 跟踪内存泄漏情况
// 释放内存调试的Tls索引
// 删除内存调试的CriticalSection
void _DbgDllProcessDetach()
{
    _EnsureThreadState();

    _DbgExTraceMemoryLeaks();

    if(g_dwTlsDbg != (DWORD)(-1))
    {
        while(s_pts)
        {
            _DbgDllThreadDetach(s_pts);
        }
        TlsFree(g_dwTlsDbg);
    }

    ::DeleteCriticalSection(&g_csDebug);
}

// 初始化线程局部对象
HRESULT _DbgDllThreadAttach()
{
    // 本地分配线程局部对象，防止引起循环分配
    DBGTHREADSTATE* pts = (DBGTHREADSTATE*)LocalAlloc(LMEM_FIXED, sizeof(DBGTHREADSTATE));

    if(!pts)
    {
        AssertSz(0, "Debug Thread initialization failed");
        return E_OUTOFMEMORY;
    }

    memset(pts, 0, sizeof(DBGTHREADSTATE));

    // 链接到全局的现场局部对象链中
    {
        DBG_LOCK_GLOBALS;

        pts->ptsNext = s_pts;
        if(s_pts)
        {
            s_pts->ptsPrev = pts;
        }
        s_pts = pts;
    }

    // 设置Tls的值为线程局部对象
    TlsSetValue(g_dwTlsDbg, pts);

    return S_OK;
}

// 释放线程局部对象
void _DbgDllThreadDetach()
{
    _DbgDllThreadDetach(_DbgGetThreadState());
}

// 释放线程局部对象
void _DbgDllThreadDetach(DBGTHREADSTATE* pts)
{
    DBGTHREADSTATE** ppts;

    if(!pts)
    {
        return;
    }

    DBG_LOCK_GLOBALS;

    // 从链中取出线程局部对象
    for(ppts=&s_pts; *ppts&&*ppts!=pts; ppts=&((*ppts)->ptsNext));
    if(*ppts)
    {
        *ppts = pts->ptsNext;
        if(pts->ptsNext)
        {
            pts->ptsNext->ptsPrev = pts->ptsPrev;
        }
    }

    // 释放线程局部对象
    LocalFree(pts);
}

// 内存分配前记录请求分配大小，返回实际分配大小
size_t WINAPI _DbgExPreAlloc(size_t cbRequest)
{
    _EnsureThreadState();
    _DbgGetThreadState()->cbRequest = cbRequest;
    return _ActualSizeFromRequestSize(cbRequest);
}

// 内存分配后，记录分配块的信息，更新全局分配信息
void* WINAPI _DbgExPostAlloc(void* pv)
{
    DBGTHREADSTATE* pts;
    DBGALLOCHDR*    pdbgah;
    DBGALLOCFOOT*   pdbgft;

    pts = _DbgGetThreadState();

    pdbgah = (DBGALLOCHDR*)pv;
    if(!pdbgah)
    {
        return NULL;
    }

    // 记录请求大小
    pdbgah->cbRequest = pts->cbRequest;

    // 记录线程ID
    pdbgah->tid = GetCurrentThreadId();

    // 初始化分配名称为空
    pdbgah->szName[0] = 0;

    // 填充保护区
    Assert(sizeof(pdbgah->adwGuard) == sizeof(g_adwHeadGuardFill));
    memcpy(pdbgah->adwGuard, g_adwHeadGuardFill, sizeof(g_adwHeadGuardFill));
    pdbgft = _FooterFromBlock(pdbgah);
    Assert(sizeof(pdbgft->adwGuard) == sizeof(g_adwFootGuardFill));
    memcpy(pdbgft->adwGuard, g_adwFootGuardFill, sizeof(g_adwFootGuardFill));

    // 更新全局分配信息
    {
        DBG_LOCK_GLOBALS;

        g_cbTotalAllocated += pts->cbRequest;
        if(g_cbMaxAllocated < g_cbTotalAllocated)
        {
            g_cbMaxAllocated = g_cbTotalAllocated;
        }

        // 分配总次数
        pdbgah->iAllocated = g_cAllocated++;

        // 更新分配块链
        pdbgah->pdbgahPrev = &g_dbgahRoot;
        pdbgah->pdbgahNext = g_dbgahRoot.pdbgahNext;
        g_dbgahRoot.pdbgahNext->pdbgahPrev = pdbgah;
        g_dbgahRoot.pdbgahNext = pdbgah;
    }

    _DbgTraceAlloc(TEXT("A"), pts->cbRequest, 0, pdbgah->iAllocated);

    pv = _RequestFromActual(pdbgah);

    // 初始化内存值为FILL_CLEAN以便调试
    memset(pv, FILL_CLEAN, pts->cbRequest);

    return pv;
}

// 释放内存前，检查内存是否合法，更新全局内存分配信息以及分配链
void* WINAPI _DbgExPreFree(void* pv)
{
    DBGALLOCHDR* pdbgah;

    _EnsureThreadState();

    if(!pv)
    {
        return NULL;
    }

    pdbgah = _ActualFromRequest(pv);

    Assert(_BlockIsValid(pdbgah));

    {
        // 为了简单没有检查死循环的情况
        DBG_LOCK_GLOBALS;

        pdbgah->pdbgahPrev->pdbgahNext = pdbgah->pdbgahNext;
        pdbgah->pdbgahNext->pdbgahPrev = pdbgah->pdbgahPrev;

        g_cbTotalAllocated -= pdbgah->cbRequest;
    }

    _DbgTraceAlloc(TEXT("F"), 0, pdbgah->cbRequest, pdbgah->iAllocated);

    // 填充实际内存块为FILL_DEAD以便调试
    memset(pdbgah, FILL_DEAD, _ActualSizeFromRequestSize(pdbgah->cbRequest));

    return pdbgah;
}

// 内存释放后
void WINAPI _DbgExPostFree()
{
}

// 重新分配内存前
size_t WINAPI _DbgExPreRealloc(void* pvRequest, size_t cbRequest, void** ppv)
{
    size_t          cb;
    DBGTHREADSTATE* pts;
    DBGALLOCHDR*    pdbgah = _ActualFromRequest(pvRequest);

    _EnsureThreadState();
    pts = _DbgGetThreadState();
    pts->cbRequest = cbRequest;
    pts->pvRequest = pvRequest;

    if(pvRequest == NULL)
    {
        // 不分配内存，返回实际分配内存大小
        *ppv = NULL;
        cb = _ActualSizeFromRequestSize(cbRequest);
    }
    else if(cbRequest == 0)
    {
        // 释放内存，返回0
        *ppv = _DbgExPreFree(pvRequest);
        cb = 0;
    }
    else
    {
        // 把内存块从全局链中取出，并更新全局信息，返回实际分配内存大小
        Assert(_BlockIsValid(pdbgah));

        {
            DBG_LOCK_GLOBALS;

            pdbgah->pdbgahPrev->pdbgahNext = pdbgah->pdbgahNext;
            pdbgah->pdbgahNext->pdbgahPrev = pdbgah->pdbgahPrev;

            g_cbTotalAllocated -= pdbgah->cbRequest;
        }

        *ppv = pdbgah;
        cb = _ActualSizeFromRequestSize(cbRequest);
    }

    return cb;
}

// 重新分配内存后
void* WINAPI _DbgExPostRealloc(void* pv)
{
    void*           pvReturn;
    DBGTHREADSTATE* pts;
    DBGALLOCHDR*    pdbgah;
    DBGALLOCFOOT*   pdbgft;

    pts = _DbgGetThreadState();

    if(pts->pvRequest == NULL)
    {
        // 这种情况等同于直接分配
        pvReturn = _DbgExPostAlloc(pv);
    }
    else if(pts->cbRequest == 0)
    {
        // 这种情况等同于释放内存
        Assert(pv == NULL);
        pvReturn = NULL;
    }
    else
    {
        DBG_LOCK_GLOBALS;

        if(pv == NULL)
        {
            // 分配失败，把以前的内存块重新放回全局链中
            Assert(pts->pvRequest);
            pdbgah = _ActualFromRequest(pts->pvRequest);

            pdbgah->pdbgahPrev = &g_dbgahRoot;
            pdbgah->pdbgahNext = g_dbgahRoot.pdbgahNext;
            g_dbgahRoot.pdbgahNext->pdbgahPrev = pdbgah;
            g_dbgahRoot.pdbgahNext = pdbgah;

            g_cbTotalAllocated += pts->cbRequest;
            if(g_cbMaxAllocated < g_cbTotalAllocated)
            {
                g_cbMaxAllocated = g_cbTotalAllocated;
            }

            pvReturn = NULL;
        }
        else
        {
            pdbgah = (DBGALLOCHDR*)pv;

            // 更新全局分配信息
            g_cbTotalAllocated += pts->cbRequest;
            if(g_cbMaxAllocated < g_cbTotalAllocated)
            {
                g_cbMaxAllocated = g_cbTotalAllocated;
            }

            _DbgTraceAlloc(TEXT("R"), pts->cbRequest, pdbgah->cbRequest, pdbgah->iAllocated);
            pdbgah->cbRequest = pts->cbRequest;

            // 加入分配链中
            pdbgah->pdbgahPrev = &g_dbgahRoot;
            pdbgah->pdbgahNext = g_dbgahRoot.pdbgahNext;
            g_dbgahRoot.pdbgahNext->pdbgahPrev = pdbgah;
            g_dbgahRoot.pdbgahNext = pdbgah;

            // 填充保护区
            pdbgft = _FooterFromBlock(pdbgah);
            Assert(sizeof(pdbgft->adwGuard) == sizeof(g_adwFootGuardFill));
            memcpy(pdbgft->adwGuard, g_adwFootGuardFill, sizeof(g_adwFootGuardFill));

            pvReturn = _RequestFromActual(pdbgah);
        }
    }

    return pvReturn;
}

// 返回实际内存块的指针地址
void* WINAPI _DbgExPreGetSize(void* pvRequest)
{
    _EnsureThreadState();
    _DbgGetThreadState()->pvRequest = pvRequest;
    return (pvRequest)?_ActualFromRequest(pvRequest):NULL;
}

// 请求块的大小
size_t WINAPI _DbgExPostGetSize(size_t cb)
{
    DBGTHREADSTATE* pts;
    pts = _DbgGetThreadState();
    return (pts->pvRequest)?_ActualFromRequest(pts->pvRequest)->cbRequest:0;
}
#endif //_DEBUG


#ifdef _DEBUG
#define DbgPreAlloc                 _DbgExPreAlloc
#define DbgPostAlloc                _DbgExPostAlloc
#define DbgPreFree                  _DbgExPreFree
#define DbgPostFree                 _DbgExPostFree
#define DbgPreRealloc               _DbgExPreRealloc
#define DbgPostRealloc              _DbgExPostRealloc
#define DbgPreGetSize               _DbgExPreGetSize
#define DbgPostGetSize              _DbgExPostGetSize
#else //_DEBUG
#define DbgPreAlloc(cb)             cb
#define DbgPostAlloc(pv)            pv
#define DbgPreFree(pv)              pv
#define DbgPostFree()
#define DbgPreRealloc(pv, cb, ppv)  cb
#define DbgPostRealloc(pv)          pv
#define DbgPreGetSize(pv)           pv
#define DbgPostGetSize(cb)          cb
#endif //!_DEBUG


//+--------------------------------------------------------------
//
//  Function:   _MemAlloc
//
//  Synopsis:   Allocate block of memory.
//
//              The contents of the block are undefined.  If the requested size
//              is zero, this function returns a valid pointer.  The returned
//              pointer is guaranteed to be suitably aligned for storage of any
//              object type.
//
//  Arguments:  [cb] - Number of bytes to allocate.
//
//  Returns:    Pointer to the allocated block, or NULL on error.
//
//---------------------------------------------------------------
void* _MemAlloc(ULONG cb)
{
    AssertSz(cb, "Requesting zero sized block.");
    if(cb == 0)
    {
        cb = 1;
    }

    Assert(g_hProcessHeap);

    return DbgPostAlloc(HeapAlloc(g_hProcessHeap, 0, DbgPreAlloc(cb)));
}

//+--------------------------------------------------------------
//  Function:   _MemAllocClear
//
//  Synopsis:   Allocate a zero filled block of memory.
//
//              If the requested size is zero, this function returns a valid
//              pointer. The returned pointer is guaranteed to be suitably
//              aligned for storage of any object type.
//
//  Arguments:  [cb] - Number of bytes to allocate.
//
//  Returns:    Pointer to the allocated block, or NULL on error.
//
//---------------------------------------------------------------
void* _MemAllocClear(ULONG cb)
{
    AssertSz(cb, "Allocating zero sized block.");

    // The small-block heap will lose its mind if this ever happens, so we
    // protect against the possibility.
    if(cb == 0)
    {
        cb = 1;
    }

    void* pv;
    Assert(g_hProcessHeap);

    pv = DbgPostAlloc(HeapAlloc(g_hProcessHeap, HEAP_ZERO_MEMORY,
        DbgPreAlloc(cb)));

    // In debug, DbgPostAlloc set the memory so we need to clear it again.
#ifdef _DEBUG
    if(pv)
    {
        memset(pv, 0, cb);
    }
#endif //_DEBUG

    return pv;
}

//+--------------------------------------------------------------
//  Function:   _MemRealloc
//
//  Synopsis:   Change the size of an existing block of memory, allocate a
//              block of memory, or free a block of memory depending on the
//              arguments.
//
//              If cb is zero, this function always frees the block of memory
//              and *ppv is set to zero.
//
//              If cb is not zero and *ppv is zero, then this function allocates
//              cb bytes.
//
//              If cb is not zero and *ppv is non-zero, then this function
//              changes the size of the block, possibly by moving it.
//
//              On error, *ppv is left unchanged.  The block contents remains
//              unchanged up to the smaller of the new and old sizes.  The
//              contents of the block beyond the old size is undefined.
//              The returned pointer is guaranteed to be suitably aligned for
//              storage of any object type.
//
//              The signature of this function is different than thy typical
//              realloc-like function to avoid the following common error:
//                  pv = realloc(pv, cb);
//              If realloc fails, then null is returned and the pointer to the
//              original block of memory is leaked.
//
//  Arguments:  [cb] - Requested size in bytes. A value of zero always frees
//                  the block.
//              [ppv] - On input, pointer to existing block pointer or null.
//                  On output, pointer to new block pointer.
//
//  Returns:    HRESULT
//
//---------------------------------------------------------------
HRESULT _MemRealloc(void** ppv, ULONG cb)
{
    void* pv;
    Assert(g_hProcessHeap);

    if(cb == 0)
    {
        _MemFree(*ppv);
        *ppv = 0;
    }
    else if(*ppv == NULL)
    {
        *ppv = _MemAlloc(cb);
        if(*ppv == NULL)
        {
            return E_OUTOFMEMORY;
        }
    }
    else
    {
#ifdef _DEBUG
        ULONG cbReq = cb;
        cb = DbgPreRealloc(*ppv, cb, &pv);
#else //_DEBUG
        pv = *ppv;
#endif //!_DEBUG

        pv = DbgPostRealloc(HeapReAlloc(g_hProcessHeap, 0, pv, cb));

        if(pv == NULL)
        {
            return E_OUTOFMEMORY;
        }
        *ppv = pv;
    }

    return S_OK;
}

//+--------------------------------------------------------------
//
//  Function:   _MemGetSize
//
//  Synopsis:   Get size of block allocated with MemAlloc/MemRealloc.
//
//              Note that MemAlloc/MemRealloc can allocate more than
//              the requested number of bytes. Therefore the size returned
//              from this function is possibly greater than the size
//              passed to MemAlloc/Realloc.
//
//  Arguments:  [pv] - Return size of this block.
//
//  Returns:    The size of the block, or zero of pv == NULL.
//
//---------------------------------------------------------------
ULONG _MemGetSize(void* pv)
{
    if(pv == NULL)
    {
        return 0;
    }

    Assert(g_hProcessHeap);

    return DbgPostGetSize(HeapSize(g_hProcessHeap, 0, DbgPreGetSize(pv)));
}

//+--------------------------------------------------------------
//
//  Function:   _MemFree
//
//  Synopsis:   Free a block of memory allocated with MemAlloc,
//              MemAllocFree or MemRealloc.
//
//  Arguments:  [pv] - Pointer to block to free.  A value of zero is
//              is ignored.
//
//---------------------------------------------------------------
void _MemFree(void* pv)
{
    // The null check is required for HeapFree.
    if(pv == NULL)
    {
        return;
    }

    Assert(g_hProcessHeap);

#ifdef _DEBUG
    pv = DbgPreFree(pv);
#endif //_DEBUG

    HeapFree(g_hProcessHeap, 0, pv);
    DbgPostFree();
}

//+------------------------------------------------------------------------
//
//  Function:   _MemAllocString
//
//  Synopsis:   Allocates a string copy using MemAlloc.
//
//              The inline function MemFreeString is provided for symmetry.
//
//  Arguments:  pchSrc    String to copy
//              ppchDest  Copy of string is returned in *ppch
//                        NULL is stored on error
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT _MemAllocString(const TCHAR* pchSrc, TCHAR** ppchDest)
{
    TCHAR* pch;
    size_t cb;

    cb = (_tcsclen(pchSrc)+1) * sizeof(TCHAR);
    *ppchDest = pch = (TCHAR*)_MemAlloc(cb);
    if(!pch)
    {
        return E_OUTOFMEMORY;
    }
    else
    {
        memcpy(pch, pchSrc, cb);
        return S_OK;
    }
}

//+------------------------------------------------------------------------
//
//  Function:   _MemAllocString
//
//  Synopsis:   Allocates a string copy using MemAlloc.  Doesn't require
//              null-terminated input string.
//
//              The inline function MemFreeString is provided for symmetry.
//
//  Arguments:  cch       number of characters in input string,
//                        not including any trailing null character
//              pchSrc    pointer to source string
//              ppchDest  Copy of string is returned in *ppch
//                        NULL is stored on error
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT _MemAllocString(ULONG cch, const TCHAR* pchSrc, TCHAR** ppchDest)
{
    TCHAR* pch;
    size_t cb = cch * sizeof(TCHAR);

    *ppchDest = pch = (TCHAR*)_MemAlloc(cb+sizeof(TCHAR));
    if(!pch)
    {
        return E_OUTOFMEMORY;
    }
    else
    {
        memcpy(pch, pchSrc, cb);
        pch[cch] = 0;
        return S_OK;
    }
}

//+------------------------------------------------------------------------
//
//  Function:   _MemReplaceString
//
//  Synopsis:   Allocates a string using MemAlloc, replacing and freeing
//              another string on success.
//
//  Arguments:  pchSrc    String to copy. May be NULL.
//              ppchDest  On success, original string is freed and copy of
//                        source string is returned here
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT _MemReplaceString(const TCHAR* pchSrc, TCHAR** ppchDest)
{
    HRESULT hr;
    TCHAR* pch;

    if(pchSrc)
    {
        hr = _MemAllocString(pchSrc, &pch);
        if(hr)
        {
            RRETURN(hr);
        }
    }
    else
    {
        pch = NULL;
    }

    MemFreeString(*ppchDest);
    *ppchDest = pch;

    return S_OK;
}

#ifdef _DEBUG
void* WINAPI _DbgExMemSetNameList(void* pvRequest, char* szFmt, va_list va)
{
    DBGALLOCHDR* pdbgah = _ActualFromRequest(pvRequest);

    char szBuf[1024];

    if(!_BlockIsValid(pdbgah))
    {
        return pvRequest;
    }

    if(pvRequest)
    {
        wvsprintfA(szBuf, szFmt, va);

        szBuf[ARRAYSIZE(pdbgah->szName)-1] = 0;
        lstrcpyA(pdbgah->szName, szBuf);
    }

    return pvRequest;
}

//+--------------------------------------------------------------
//
//  Function:   _DbgExMemSetName
//
//  Synopsis:   Sets the descriptive name of a memory block
//
//  Arguments:  pv  Pointer to the requested allocation
//
//  Returns:    Pointer to the requested allocation
//
//---------------------------------------------------------------
void* __cdecl _DbgExMemSetName(void* pvRequest, char* szFmt, ...)
{
    va_list va;
    void* pv;

    va_start(va, szFmt);
    pv = _DbgExMemSetNameList(pvRequest, szFmt, va);
    va_end(va);

    return(pv);
}

void _MemSetName(void* pv, char* szFmt, ...)
{
    char szBuf[1024];
    va_list va;

    va_start(va, szFmt);
    wvsprintfA(szBuf, szFmt, va);
    va_end(va);

    _DbgExMemSetName(pv, "%s", szBuf);
}

//+--------------------------------------------------------------
//
//  Function:   _DbgExMemGetName
//
//  Synopsis:   Gets the descriptive name of a memory block
//
//  Arguments:  pv  Pointer to the requested allocation
//
//  Returns:    Pointer to the requested allocation
//
//---------------------------------------------------------------
char* WINAPI _DbgExMemGetName(void* pvRequest)
{
    return _ActualFromRequest(pvRequest)->szName;
}

char* _MemGetName(void* pv)
{
    return _DbgExMemGetName(pv);
}
#endif //_DEBUG