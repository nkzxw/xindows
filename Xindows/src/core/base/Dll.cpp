
#include "stdafx.h"
#include "Dll.h"

extern HRESULT DllThreadAttach();
extern void DllThreadDetach(THREADSTATE* pts);
extern void DllThreadPassivate(THREADSTATE* pts);

extern void DeinitSurface();

extern void KillDwnTaskExec();
extern void KillImgTaskExec();

extern void DeinitMultiLanguage();
extern void DeinitUniscribe();

ULONG g_ulLcl = 0;

//+----------------------------------------------------------------------------
//
//  Function:   DllAllThreadsDetach
//
//  Synopsis:   Cleanup when all Trident threads go away.
//
//-----------------------------------------------------------------------------
void DllAllThreadsDetach()
{
    DeinitSurface();
    KillDwnTaskExec();
    KillImgTaskExec();
    DeinitMultiLanguage();
    DeinitUniscribe();

    if(_afxGlobalData._hInstResource)
    {
        ::FreeLibrary(_afxGlobalData._hInstResource);
        _afxGlobalData._hInstResource = NULL;
    }
}

//+------------------------------------------------------------------
//
//  Function:   _IncrementObjectCount
//
//  Synopsis:   Increment the per-thread object count
//
//-------------------------------------------------------------------
void _IncrementObjectCount()
{
    TLS(dll.lObjCount)++;
}

//+------------------------------------------------------------------
//
//  Function:   _DecrementObjectCount
//
//  Synopsis:   Decrement the per-thread object count and passivate the
//              thread when it transitions to zero.
//
//-------------------------------------------------------------------
void _DecrementObjectCount()
{
    THREADSTATE* pts = GetThreadState();

    Assert(pts->dll.lObjCount > 0);

    if(--pts->dll.lObjCount == 0)
    {
        pts->dll.lObjCount += ULREF_IN_DESTRUCTOR;
        DllThreadPassivate(pts);
        pts->dll.lObjCount -= ULREF_IN_DESTRUCTOR;
        Assert(pts->dll.lObjCount == 0);
        DllThreadDetach(pts);
        LOCK_GLOBALS;
        if(_afxGlobalData._pts == NULL)
        {
            DllAllThreadsDetach();
        }
    }
}

//+------------------------------------------------------------------
//
//  Function:   _AddRefThreadState
//
//  Synopsis:   Prepare per-thread variables
//
//  Returns:    S_OK, E_OUTOFMEMORY
//
//-------------------------------------------------------------------
HRESULT _AddRefThreadState()
{
    HRESULT hr;
    if(TlsGetValue(_afxGlobalData._dwTls) == NULL)
    {
        hr = DllThreadAttach();
        if(hr)
        {
            RRETURN(hr);
        }
    }

    IncrementObjectCount();

    return S_OK;
}
