
#include "stdafx.h"
#include "Image.h"

CImgTaskExec* g_pImgTaskExec;

CImgLoad::~CImgLoad()
{
    if(_pImgTask)
    {
        _pImgTask->SubRelease();
    }
}

void CImgLoad::Passivate()
{
    super::Passivate();

    if(_pImgTask)
    {
        // The task is needed by the asynchronous callback methods, but here
        // we want to passivate it by releasing the last reference but
        // maintaining a secondary reference which will be released by the
        // destructor.
        _pImgTask->SubAddRef();
        _pImgTask->Release();
    }
}

HRESULT CImgLoad::Init(DWNLOADINFO* pdli, CDwnInfo* pDwnInfo)
{
    HRESULT hr;

    hr = super::Init(pdli, pDwnInfo, 
        0/*IDS_BINDSTATUS_DOWNLOADINGDATA_PICTURE wlw note*/,
        DWNF_GETMODTIME|DWNF_GETFLAGS);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CImgLoad::OnBindHeaders()
{
    FILETIME ft;
    HRESULT hr = S_OK;

    if(!_pDwnInfo->TstFlags(DWNF_DOWNLOADONLY))
    {
        ft = _pDwnBindData->GetLastMod();

        if(ft.dwLowDateTime || ft.dwHighDateTime)
        {
            if(_pDwnInfo->AttachByLastMod(this, &ft, _pDwnBindData->IsFullyAvail()))
            {
                CDwnDoc* pDwnDoc = _pDwnBindData->GetDwnDoc();

                if(pDwnDoc)
                {
                    DWNPROG DwnProg;
                    _pDwnBindData->GetProgress(&DwnProg);
                    pDwnDoc->AddBytesRead(DwnProg.dwMax);
                }

                _pDwnBindData->Disconnect();
                OnDone(S_OK);

                hr = S_FALSE;
                goto Cleanup;
            }
        }
    }

    _pDwnInfo->SetSecFlags(_pDwnBindData->GetSecFlags());

Cleanup:    
    RRETURN1(hr, S_FALSE);
}

HRESULT CImgLoad::OnBindMime(MIMEINFO* pmi)
{
    CImgTask* pImgTask = NULL;
    HRESULT   hr       = S_OK;

    if(!pmi || _pDwnInfo->TstFlags(DWNF_DOWNLOADONLY))
    {
        hr = S_OK;
        goto Cleanup;
    }

    if(!pmi->pfnImg)
    {
        hr = E_ABORT;
        goto Cleanup;
    }

    _pDwnInfo->SetMimeInfo(pmi);

    pImgTask = pmi->pfnImg();

    if(pImgTask == NULL)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pImgTask->Init(GetImgInfo(), pmi, _pDwnBindData);

    hr = StartImgTask(pImgTask);
    if(hr)
    {
        goto Cleanup;
    }

    EnterCriticalSection();

    if(_fPassive)
    {
        hr = E_ABORT;
    }
    else
    {
        _pImgTask = pImgTask;
        pImgTask  = NULL;
    }

    LeaveCriticalSection();

    if(hr == S_OK)
    {
        GetImgInfo()->OnLoadTask(this, _pImgTask);
    }

Cleanup:
    if(pImgTask)
    {
        pImgTask->Release();
    }
    RRETURN(hr);
}

HRESULT CImgLoad::OnBindData()
{
    HRESULT hr = S_OK;

    if(_pImgTask)
    {
        _pImgTask->SetBlocked(FALSE);
    }
    else if(_pDwnInfo->TstFlags(DWNF_DOWNLOADONLY))
    {
        BYTE  ab[1024];
        ULONG cb;

        do
        {
            hr = _pDwnBindData->Read(ab, sizeof(ab), &cb);
        } while(!hr && cb);
    }
    else
    {
        // If we're getting data but never got a valid mime type that we
        // know how to decode, use the data to figure out if a pluggable
        // decoder should be used.
        BYTE        ab[200];
        ULONG       cb;
        MIMEINFO*   pmi;

        hr = _pDwnBindData->Peek(ab, ARRAYSIZE(ab), &cb);
        if(hr)
        {
            goto Cleanup;
        }

        if(cb<ARRAYSIZE(ab) && _pDwnBindData->IsPending())
        {
            goto Cleanup;
        }

        pmi = GetMimeInfoFromData(ab, cb, _pDwnBindData->GetContentType());

        if(!pmi || !pmi->pfnImg)
        {
            pmi = _pDwnBindData->GetRawMimeInfoPtr();
            if(!pmi || !pmi->pfnImg)
            {
                hr = E_ABORT;
                goto Cleanup;
            }
        }

        hr = OnBindMime(pmi);
        if(hr)
        {
            goto Cleanup;
        }

        _pImgTask->SetBlocked(FALSE);
    }

Cleanup:
    RRETURN(hr);
}

void CImgLoad::OnBindDone(HRESULT hrErr)
{
    if(_pImgTask)
    {
        _pImgTask->SetBlocked(FALSE);
    }

    OnDone(hrErr);
}



void CImgTaskExec::YieldTask(CImgTask* pImgTask, BOOL fBlock)
{
    if(fBlock)
    {
        pImgTask->SetBlocked(TRUE);
    }

    SwitchToFiber(_pvFiberMain);
}

FIBERINFO* CImgTaskExec::GetFiber(CImgTask* pImgTask)
{
    BOOL        fAll = pImgTask->IsFullyAvail() ? 0 : 1;
    FIBERINFO*  pfi  = &_afi[fAll];
    UINT        cfi  = ARRAYSIZE(_afi) - fAll;

    for(; cfi>0; --cfi,++pfi)
    {
        if(pfi->pImgTask == NULL)
        {
            goto found;
        }
    }

    return NULL;

found:
    if(pfi->pvFiber == NULL)
    {
        pfi->pvMain = _pvFiberMain;

        if(pfi->pvMain)
        {
            pfi->pvFiber = CreateFiber(0x8000, FiberProc, pfi);
        }

        if(pfi->pvFiber == NULL)
        {
            return NULL;
        }
    }

    return pfi;
}

void CImgTaskExec::AssignFiber(FIBERINFO* pfi)
{
    BOOL fAll = (pfi == &_afi[0]);

    EnterCriticalSection();

    CImgTask* pImgTask = (CImgTask*)_pDwnTaskHead;

    for(; pImgTask; pImgTask=(CImgTask*)pImgTask->_pDwnTaskNext)
    {
        if(pImgTask->_fWaitForFiber
            && !pImgTask->_fTerminate
            && !pImgTask->_pfi
            && (!fAll || pImgTask->IsFullyAvail()))
        {
            pfi->pImgTask = pImgTask;
            pImgTask->_pfi = pfi;
            pImgTask->_fWaitForFiber = FALSE;
            pImgTask->SubAddRef();
            pImgTask->SetBlocked(FALSE);
            goto Cleanup;
        }
    }

Cleanup:
    LeaveCriticalSection();    
}

void CImgTaskExec::RunTask(CImgTask* pImgTask)
{
    FIBERINFO* pfi = pImgTask->_pfi;

    if(pfi==NULL && !pImgTask->_fTerminate)
    {
        pfi = GetFiber(pImgTask);

        if(pfi)
        {
            pfi->pImgTask = pImgTask;
            pImgTask->_pfi = pfi;
            pImgTask->_fWaitForFiber = FALSE;
            pImgTask->SubAddRef();
        }
        else
        {
            // No fiber available.  Note that when one becomes available
            // there may be a task waiting to hear about it.
            pImgTask->_fWaitForFiber = TRUE;
            pImgTask->SetBlocked(TRUE);
            goto Cleanup;
        }
    }

    if(pfi)
    {
        Assert(pfi->pImgTask == pImgTask);
        Assert(pImgTask->_pfi == pfi);

        SwitchToFiber(pfi->pvFiber);

        if(pImgTask->_pfi == NULL)
        {
            pImgTask->SubRelease();
            AssignFiber(pfi);
        }
    }

    if(pImgTask->_fTerminate)
    {
        pImgTask->_pImgInfo->OnTaskDone(pImgTask);
        super::DelTask(pImgTask);
    }

Cleanup:
    return;
}

HRESULT CImgTaskExec::ThreadInit()
{
    HRESULT hr;

    hr = super::ThreadInit();
    if(hr)
    {
        goto Cleanup;
    }

    if(TRUE)
    {
        _pvFiberMain = ConvertThreadToFiber(0);
    }

Cleanup:
    RRETURN(hr);
}

void CImgTaskExec::ThreadTerm()
{
    FIBERINFO*  pfi = _afi;
    UINT        cfi = ARRAYSIZE(_afi);

    for(; cfi>0; --cfi,++pfi)
    {
        if(pfi->pImgTask)
        {
            pfi->pImgTask->_pfi = NULL;
            pfi->pImgTask->SubRelease();
            pfi->pImgTask = NULL;
        }
        if(pfi->pvFiber)
        {
            DeleteFiber(pfi->pvFiber);
            pfi->pvFiber = NULL;
        }
    }

    // Manually release any tasks remaining on the queue.  Don't call super::ThreadTerm
    // because it tries to call Terminate() on the task and expects it to dequeue.  Our
    // tasks don't dequeue in Terminate() though, so we end up in an infinite loop.
    while(_pDwnTaskHead)
    {
        CImgTask* pImgTask = (CImgTask*)_pDwnTaskHead;
        _pDwnTaskHead = pImgTask->_pDwnTaskNext;
        Assert(pImgTask->_fEnqueued);
        pImgTask->SubRelease();
    }
}

void CImgTaskExec::ThreadExit()
{
    void* pvMain = _pvFiberMain;

    if(_fCoInit)
    {
        CoUninitialize();
    }

    super::ThreadExit();

    if(pvMain)
    {
        // Due to a bug in the Win95 and WinNT implementation of fibers,
        // we don't call FbrDeleteFiber(pvMain) anymore.  Instead we manually
        // free the fiber data for the main fiber on this thread.
        LocalFree(pvMain);
    }
}

void CImgTaskExec::ThreadTimeout()
{
    KillImgTaskExec();
}

void WINAPI CImgTaskExec::FiberProc(void* pv)
{
    FIBERINFO* pfi = (FIBERINFO*)pv;
    CImgTask* pImgTask = NULL;

    while(pfi)
    {
        pImgTask = pfi->pImgTask;
        pImgTask->Exec();
        pfi->pImgTask = NULL;
        pImgTask->_pfi = NULL;

        SwitchToFiber(pfi->pvMain);
    }
}

HRESULT CImgTaskExec::RequestCoInit()
{
    HRESULT hr = S_OK;

    if(!_fCoInit)
    {
        hr = CoInitialize(NULL);

        if(!FAILED(hr))
        {
            _fCoInit = TRUE;
            hr = S_OK;
        }
    }

    RRETURN(hr);
}



HRESULT StartImgTask(CImgTask* pImgTask)
{
    HRESULT hr = S_OK;

    g_csImgTaskExec.Enter();

    if(g_pImgTaskExec == NULL)
    {
        g_pImgTaskExec = new CImgTaskExec(g_csImgTaskExec.GetPcs());

        if(g_pImgTaskExec == NULL)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        hr = g_pImgTaskExec->Launch();

        if(hr)
        {
            g_pImgTaskExec->Release();
            g_pImgTaskExec = NULL;
            goto Cleanup;
        }
    }

    g_pImgTaskExec->AddTask(pImgTask);

Cleanup:
    g_csImgTaskExec.Leave();
    RRETURN(hr);
}

void KillImgTaskExec()
{
    g_csImgTaskExec.Enter();

    CImgTaskExec* pImgTaskExec = g_pImgTaskExec;
    g_pImgTaskExec = NULL;

    g_csImgTaskExec.Leave();

    if(pImgTaskExec)
    {
        pImgTaskExec->Shutdown();
        pImgTaskExec->Release();
    }
}