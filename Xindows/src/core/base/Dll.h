
#ifndef __XINDOWS_CORE_BASE_DLL_H__
#define __XINDOWS_CORE_BASE_DLL_H__

//+------------------------------------------------------------------------
//
//  DLL object count functions.
//
//      The dll object count is used to implement DLLCanUnloadNow.
//
//  NOTE:   For now, incrementing the per-thread object count also increments
//          the global secondary count. This is necessary for DllCanUnloadNow
//          to work correctly. This makes some calls to increment/decrement
//          the secondary object count redundant; we should consider optimizing
//          these away.
//
//-------------------------------------------------------------------------
#define IncrementObjectCount()  _IncrementObjectCount()
#define DecrementObjectCount()  _DecrementObjectCount()

void    _IncrementObjectCount();
void    _DecrementObjectCount();

inline LONG GetPrimaryObjectCount();

#define AddRefThreadState()     _AddRefThreadState()
#define ReleaseThreadState()    _DecrementObjectCount()
HRESULT _AddRefThreadState();

inline THREADSTATE* GetThreadState()
{
    Assert(TlsGetValue(_afxGlobalData._dwTls));
    return (THREADSTATE*)(TlsGetValue(_afxGlobalData._dwTls));
}

#define TLS(x)  (GetThreadState()->x) // Access a per-thread variable

inline LONG GetPrimaryObjectCount()
{
    return TLS(dll.lObjCount);
}

class CEnsureThreadState
{
public:
    CEnsureThreadState()  { _hr = AddRefThreadState(); }
    ~CEnsureThreadState() { if(_hr == S_OK) ReleaseThreadState(); }
    HRESULT _hr;
};

extern ULONG g_ulLcl; // Global incrementing counter
//+---------------------------------------------------------------------------
//
//  Function:   IncrementLcl/GetLcl
//
//  Synopsis:   IncrmentLcl returns a ULONG which is bigger than any
//              previously given, and GetLcl returns a ULONG which
//              is at least as big as any previously given. Used for
//              determining order when syncrhonizing threads.
//
//----------------------------------------------------------------------------
inline ULONG IncrementLcl()
{
    LOCK_GLOBALS;

    // Cannot use InterlockedIncrement because it only returns +/-/0.
    ULONG lcl = ++g_ulLcl;

    AssertSz(lcl, "Lcl overflow");
    return lcl;
}

inline ULONG GetLcl()
{
    return g_ulLcl;
}

#endif //__XINDOWS_CORE_BASE_DLL_H__