
#include "stdafx.h"
#include "DownBase.h"

CBaseFT::CBaseFT(CRITICAL_SECTION* pcs)
{
    _ulRefs     = 1;
    _ulAllRefs  = 1;
    _pcs        = pcs;
}

CBaseFT::~CBaseFT()
{
    AssertSz(_ulAllRefs==0, "Ref count messed up in derived dtor?");
    AssertSz(_ulRefs==0, "Ref count messed up in derived dtor?");
}

void CBaseFT::Passivate()
{
    AssertSz(_ulRefs==0,
        "CBaseFT::Passivate called unexpectedly or refcnt "
        "messed up in derived Passivate");
}

ULONG CBaseFT::Release()
{
    ULONG ulRefs = (ULONG)InterlockedDecrement((LONG*)&_ulRefs);

    if(ulRefs == 0)
    {
        Passivate();
        AssertSz(_ulRefs==0, "CBaseFT::AddRef occured after last release");
        SubRelease();
    }

    return ulRefs;
}

ULONG CBaseFT::SubRelease()
{
    ULONG ulRefs = (ULONG)InterlockedDecrement((LONG*)&_ulAllRefs);

    if(ulRefs == 0)
    {
        delete this;
    }

    return ulRefs;
}

#ifdef _DEBUG
void CBaseFT::EnterCriticalSection()
{
    if(_pcs)
    {
        ::EnterCriticalSection(_pcs);

        Assert(_dwThread==0 || _dwThread==GetCurrentThreadId());

        if(_dwThread == 0)
        {
            _dwThread = GetCurrentThreadId();
        }

        _cEnter += 1;
    }
}

void CBaseFT::LeaveCriticalSection()
{
    if(_pcs)
    {
        Assert(_dwThread == GetCurrentThreadId());
        Assert(_cEnter > 0);

        if(--_cEnter == 0)
        {
            _dwThread = 0;
        }

        ::LeaveCriticalSection(_pcs);
    }
}

BOOL CBaseFT::EnteredCriticalSection()
{
    return TRUE;
}
#endif


CExecFT::CExecFT(CRITICAL_SECTION* pcs) : CBaseFT(pcs)
{
    _hThread = NULL;
    _hEvent  = NULL;
    _hrInit  = S_OK;
}

CExecFT::~CExecFT()
{
    CloseHandle(_hThread);
    CloseHandle(_hEvent);
}

void CExecFT::Passivate()
{
}

DWORD WINAPI CExecFT::StaticThreadProc(void* pv)
{
    return (((CExecFT*)pv)->ThreadProc());
}

DWORD CExecFT::ThreadProc()
{
    _hrInit = ThreadInit();

    if(_hEvent)
    {
        Verify(SetEvent(_hEvent));
    }

    if(_hrInit == S_OK)
    {
        ThreadExec();
    }

    ThreadTerm();
    ThreadExit();
    return 0;
}

void CExecFT::ThreadExit()
{
    SubRelease();
}

HRESULT CExecFT::Launch(BOOL fWait)
{
    DWORD dwResult;
    DWORD dwStackSize = 0;

    if(fWait)
    {
        _hEvent = CreateEventA(NULL, FALSE, FALSE, NULL);

        if(_hEvent == NULL)
        {
            RRETURN(GetLastWin32Error());
        }
    }

    SubAddRef();

    _hThread = CreateThread(NULL, dwStackSize, &CExecFT::StaticThreadProc, this, 0, &_dwThreadId);

    if(_hThread == NULL)
    {
        SubRelease();
        RRETURN(GetLastWin32Error());
    }

    if(fWait)
    {
        dwResult = WaitForSingleObject(_hEvent, INFINITE);

        Assert(dwResult == WAIT_OBJECT_0);

        CloseHandle(_hEvent);
        _hEvent = NULL;

        RRETURN(_hrInit);
    }

    return S_OK;
}

void CExecFT::Shutdown(DWORD dwTimeOut)
{
    if(_hThread && GetCurrentThreadId()!=_dwThreadId)
    {
        DWORD dwExitCode;

        WaitForSingleObject(_hThread, dwTimeOut);

        if(GetExitCodeThread(_hThread, &dwExitCode) && dwExitCode==STILL_ACTIVE)
        {
            TerminateThread(_hThread, 1);
        }
    }
}
