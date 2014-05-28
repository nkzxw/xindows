
#include "stdafx.h"
#include "GlobalLock.h"

CRITICAL_SECTION s_cs;

#ifdef _DEBUG
DWORD           s_dwThreadID;
LONG            s_cNesting;
#endif //_DEBUG

CGlobalLock::CGlobalLock()
{
    EnterCriticalSection(&s_cs);
#ifdef _DEBUG
    if(!s_cNesting)
    {
        s_dwThreadID = GetCurrentThreadId();
    }
    else
    {
        Assert(s_dwThreadID == GetCurrentThreadId());
    }
    Assert(++s_cNesting > 0);
#endif //_DEBUG
}

CGlobalLock::~CGlobalLock()
{
#ifdef _DEBUG
    Assert(s_dwThreadID == GetCurrentThreadId());
    Assert(--s_cNesting >= 0);
    if(!s_cNesting)
    {
        s_dwThreadID = 0;
    }
#endif //_DEBUG
    LeaveCriticalSection(&s_cs);
}

#ifdef _DEBUG
BOOL CGlobalLock::IsThreadLocked()
{
    return (s_dwThreadID==GetCurrentThreadId());
}
#endif //_DEBUG

// Process attach/detach routines
void CGlobalLock::Init()
{
    InitializeCriticalSection(&s_cs);
}

void CGlobalLock::Deinit()
{
#ifdef _DEBUG
    if(s_cNesting)
    {
        Trace1("Global lock count > 0, Count=%0d\n", s_cNesting);
    }
#endif //_DEBUG
    DeleteCriticalSection(&s_cs);
}
