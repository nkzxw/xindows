
#include "stdafx.h"
#include "Dwn.h"

// Globals --------------------------------------------------------------------
CDwnTaskExec* g_pDwnTaskExec = NULL;

// CDwnTask -------------------------------------------------------------------
void CDwnTask::Passivate()
{
    Terminate();
    super::Passivate();
}

// CDwnTaskExec ---------------------------------------------------------------
CDwnTaskExec::CDwnTaskExec(CRITICAL_SECTION* pcs) : super(pcs)
{
}

CDwnTaskExec::~CDwnTaskExec()
{
}

void CDwnTaskExec::AddTask(CDwnTask* pDwnTask)
{
    BOOL fSignal = FALSE;

    EnterCriticalSection();

    Assert(!pDwnTask->_fEnqueued);
    Assert(!pDwnTask->_pDwnTaskExec);

    pDwnTask->_pDwnTaskExec = this;
    pDwnTask->_pDwnTaskNext = NULL;
    pDwnTask->_pDwnTaskPrev = _pDwnTaskTail;
    pDwnTask->_fEnqueued    = TRUE;
    pDwnTask->SubAddRef();

    if(_pDwnTaskTail)
    {
        _pDwnTaskTail = _pDwnTaskTail->_pDwnTaskNext = pDwnTask;
    }
    else
    {
        _pDwnTaskHead = _pDwnTaskTail = pDwnTask;
    }

    if(pDwnTask->_fActive)
    {
        fSignal = (_cDwnTaskActive == 0);
        _cDwnTaskActive += 1;
    }

    _cDwnTask += 1;

    LeaveCriticalSection();

    if(fSignal)
    {
        Verify(SetEvent(_hevWait));
    }
}

void CDwnTaskExec::SetTask(CDwnTask* pDwnTask, BOOL fActive)
{
    BOOL fSignal = FALSE;

    EnterCriticalSection();

    if(pDwnTask->_fEnqueued)
    {
        if(pDwnTask == _pDwnTaskRun)
        {
            // Making a running task active always wins ... the task will
            // run from the top at least one more time.
            if(_ta != TA_DELETE)
            {
                if(fActive)
                {
                    _ta = TA_ACTIVATE;
                }
                else if(_ta != TA_ACTIVATE)
                {
                    _ta = TA_BLOCK;
                }
            }
        }
        else if(!!fActive != !!pDwnTask->_fActive)
        {
            pDwnTask->_fActive = fActive;

            if(fActive)
            {
                fSignal = (_cDwnTaskActive == 0);
                _cDwnTaskActive += 1;
            }
            else
            {
                _cDwnTaskActive -= 1;
            }
        }
    }

    LeaveCriticalSection();

    if(fSignal && GetThreadId()!=GetCurrentThreadId())
    {
        Verify(SetEvent(_hevWait));
    }
}

void CDwnTaskExec::DelTask(CDwnTask* pDwnTask)
{
    BOOL fRelease = FALSE;

    EnterCriticalSection();

    if(pDwnTask->_fEnqueued)
    {
        if(pDwnTask == _pDwnTaskRun)
        {
            _ta = TA_DELETE;
        }
        else
        {
            if(pDwnTask->_pDwnTaskPrev)
            {
                pDwnTask->_pDwnTaskPrev->_pDwnTaskNext = pDwnTask->_pDwnTaskNext;
            }
            else
            {
                _pDwnTaskHead = pDwnTask->_pDwnTaskNext;
            }

            if(pDwnTask->_pDwnTaskNext)
            {
                pDwnTask->_pDwnTaskNext->_pDwnTaskPrev = pDwnTask->_pDwnTaskPrev;
            }
            else
            {
                _pDwnTaskTail = pDwnTask->_pDwnTaskPrev;
            }

            if(pDwnTask->_fActive)
            {
                _cDwnTaskActive -= 1;
            }

            _cDwnTask -= 1;

            if(_pDwnTaskCur == pDwnTask)
            {
                _pDwnTaskCur = pDwnTask->_pDwnTaskNext;
            }

            pDwnTask->_fEnqueued = FALSE;
            fRelease = TRUE;
        }
    }

    LeaveCriticalSection();

    if(fRelease)
    {
        pDwnTask->SubRelease();
    }
}

BOOL CDwnTaskExec::IsTaskTimeout()
{
    return (_fShutdown || (GetTickCount()-_dwTickRun>_dwTickSlice));
}

HRESULT CDwnTaskExec::Launch()
{
    HRESULT hr;

    _hevWait = CreateEventA(NULL, FALSE, FALSE, NULL);

    if(_hevWait == NULL)
    {
        hr = GetLastWin32Error();
        goto Cleanup;
    }

    hr = super::Launch(FALSE);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

void CDwnTaskExec::Passivate()
{
    if(_hevWait)
    {
        CloseHandle(_hevWait);
    }
}

HRESULT CDwnTaskExec::ThreadInit()
{
    _dwTickTimeout  = 10 * 60 * 1000;   // Ten minutes
    _dwTickSlice    = 200;              // 0.2 seconds

    return S_OK;
}

void CDwnTaskExec::ThreadExec()
{
    CDwnTask* pDwnTask;

    for(;;)
    {
        for(;;)
        {
            EnterCriticalSection();

            pDwnTask     = _pDwnTaskRun;
            _pDwnTaskRun = NULL;

            if(_ta == TA_DELETE)
            {
                DelTask(pDwnTask);
            }
            else if (_ta == TA_BLOCK)
            {
                pDwnTask->SetBlocked(TRUE);
            }

            _ta = TA_NONE;

            if(_cDwnTaskActive && !_fShutdown)
            {
                while(!_pDwnTaskRun)
                {
                    if(_pDwnTaskCur == NULL)
                    {
                        _pDwnTaskCur = _pDwnTaskHead;
                    }

                    if(_pDwnTaskCur->_fActive)
                    {
                        _pDwnTaskRun = _pDwnTaskCur;
                    }

                    _pDwnTaskCur = _pDwnTaskCur->_pDwnTaskNext;
                }
            }

            LeaveCriticalSection();

            if(!_pDwnTaskRun)
            {
                break;
            }

            _dwTickRun = GetTickCount();
            _pDwnTaskRun->Run();
        }

        if(_fShutdown)
        {
            break;
        }

        DWORD dwResult = WaitForSingleObject(_hevWait, _dwTickTimeout);

        if(dwResult == WAIT_TIMEOUT)
        {
            EnterCriticalSection();

            if(_cDwnTask == 0)
            {
                ThreadTimeout();
            }

            LeaveCriticalSection();
        }
    }
}

void CDwnTaskExec::ThreadTimeout()
{
    KillDwnTaskExec();
}

void CDwnTaskExec::ThreadTerm()
{
    while(_pDwnTaskHead)
    {
        _pDwnTaskHead->Terminate();
    }
}

void CDwnTaskExec::Shutdown()
{
    _fShutdown = TRUE;
    Verify(SetEvent(_hevWait));
    super::Shutdown();
}


// External Functions ---------------------------------------------------------
HRESULT StartDwnTask(CDwnTask* pDwnTask)
{
    HRESULT hr = S_OK;

    g_csDwnTaskExec.Enter();

    if(g_pDwnTaskExec == NULL)
    {
        g_pDwnTaskExec = new CDwnTaskExec(g_csDwnTaskExec.GetPcs());

        if(g_pDwnTaskExec == NULL)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        hr = g_pDwnTaskExec->Launch();

        if(hr)
        {
            g_pDwnTaskExec->Release();
            g_pDwnTaskExec = NULL;
            goto Cleanup;
        }
    }

    g_pDwnTaskExec->AddTask(pDwnTask);

Cleanup:
    g_csDwnTaskExec.Leave();
    RRETURN(hr);
}

void KillDwnTaskExec()
{
    g_csDwnTaskExec.Enter();

    CDwnTaskExec* pDwnTaskExec = g_pDwnTaskExec;
    g_pDwnTaskExec = NULL;

    g_csDwnTaskExec.Leave();

    if(pDwnTaskExec)
    {
        pDwnTaskExec->Shutdown();
        pDwnTaskExec->Release();
    }
}