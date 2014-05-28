
#include "stdafx.h"
#include "Download.h"

#include "Dwn.h"

// Globals --------------------------------------------------------------------
DEFINE_CRITICAL(g_csDwnBindPend, 0); // CDwnBindData (SetPending, SetEof)
DEFINE_CRITICAL(g_csDwnChanSig,  0); // CDwnChan (Signal)
DEFINE_CRITICAL(g_csDwnStm,      0); // CDwnStm
DEFINE_CRITICAL(g_csDwnDoc,      0); // CDwnDoc
DEFINE_CRITICAL(g_csDwnPost,     0); // CDwnPost (GetStgMedium)
DEFINE_CRITICAL(g_csDwnBindSig,  1); // CDwnBind (Signal)
DEFINE_CRITICAL(g_csDwnBindTerm, 1); // CDwnBind (Terminate)
DEFINE_CRITICAL(g_csDwnCache,    1); // CDwnInfo active and cache lists
DEFINE_CRITICAL(g_csDwnTaskExec, 1); // CDwnTaskExec
DEFINE_CRITICAL(g_csImgTaskExec, 1); // CImgTaskExec
DEFINE_CRITICAL(g_csImgTransBlt, 1); // CImgBitsDIB::StretchBlt


// CDwnChan -------------------------------------------------------------------
CDwnChan::CDwnChan(CRITICAL_SECTION* pcs) : super(pcs)
{
    _fSignalled = TRUE;
}

void CDwnChan::Passivate()
{
    Disconnect();
}

void CDwnChan::SetCallback(PFNDWNCHAN pfnCallback, void* pvCallback)
{
    HRESULT hr = AddRefThreadState();

    Disconnect();

    if(hr == S_OK)
    {
        Assert(_fSignalled);

        _pts          = GetThreadState();
        _pfnCallback  = pfnCallback;
        _pvCallback   = pvCallback;
        _fSignalled   = FALSE;
    }
}

void CDwnChan::Disconnect()
{
    if(_pts && (_pts==GetThreadState()))
    {
        THREADSTATE* pts;
        BOOL fSignalled;

        g_csDwnChanSig.Enter();

        fSignalled   = _fSignalled;
        pts          = _pts;
        _pts         = NULL;
        _pfnCallback = NULL;
        _pvCallback  = NULL;
        _fSignalled  = TRUE;

        g_csDwnChanSig.Leave();

        if(fSignalled)
        {
            GWKillMethodCallEx(pts, this, ONCALL_METHOD(CDwnChan, OnMethodCall, onmethodcall), 0);
        }

        ReleaseThreadState();
    }
}

void CDwnChan::Signal()
{
    if(!_fSignalled)
    {
        g_csDwnChanSig.Enter();

        if(!_fSignalled)
        {
            _fSignalled = TRUE;
            GWPostMethodCallEx(_pts, this, ONCALL_METHOD(CDwnChan, OnMethodCall, onmethodcall), 0, FALSE, GetOnMethodCallName());
        }

        g_csDwnChanSig.Leave();
    }
}

void CDwnChan::OnMethodCall(DWORD_PTR dwContext)
{
    if(_fSignalled)
    {
        _fSignalled = FALSE;
        _pfnCallback(this, _pvCallback);
    }
}



// CDwnCtx --------------------------------------------------------------------
#ifdef _DEBUG
void CDwnCtx::EnterCriticalSection()
{
    ((CDwnCrit*)GetPcs())->Enter();
}

void CDwnCtx::LeaveCriticalSection()
{
    ((CDwnCrit*)GetPcs())->Leave();
}

BOOL CDwnCtx::EnteredCriticalSection()
{
    return ((CDwnCrit*)GetPcs())->IsEntered();
}
#endif

void CDwnCtx::Passivate()
{
    SetLoad(FALSE, NULL, FALSE);

    super::Passivate();

    if(_pDwnInfo)
    {
        _pDwnInfo->DelDwnCtx(this);
    }

    ClearInterface(&_pProgSink);
}

LPCTSTR CDwnCtx::GetUrl()
{
    return (_pDwnInfo ? _pDwnInfo->GetUrl() : _afxGlobalData._Zero.ach);
}

MIMEINFO* CDwnCtx::GetMimeInfo()
{
    return(_pDwnInfo ? _pDwnInfo->GetMimeInfo() : NULL);
}

HRESULT CDwnCtx::GetFile(LPTSTR* ppch)
{
    *ppch = NULL;
    RRETURN(_pDwnInfo ? _pDwnInfo->GetFile(ppch) : E_FAIL);
}

FILETIME CDwnCtx::GetLastMod()
{
    if(_pDwnInfo)
    {
        return _pDwnInfo->GetLastMod();
    }
    else
    {
        FILETIME ftZ = { 0 };
        return ftZ;
    }
}

DWORD CDwnCtx::GetSecFlags()
{
    return (_pDwnInfo ? _pDwnInfo->GetSecFlags() : 0);
}

HRESULT CDwnCtx::SetProgSink(IProgSink* pProgSink)
{
    HRESULT hr = S_OK;

    EnterCriticalSection();

#ifdef _DEBUG
    if(pProgSink)
    {
        if(!_pDwnInfo)
        {
            Trace0("CDwnCtx::SetProgSink called with no _pDwnInfo");
        }
        else
        {
            if(_pDwnInfo->GetFlags(DWNF_STATE) & (DWNLOAD_COMPLETE|DWNLOAD_ERROR|DWNLOAD_STOPPED))
            {
                Trace0("CDwnCtx::SetProgSink called when _pDwnInfo is already done");
            }
        }
    }
#endif

    if(_pDwnInfo)
    {
        if(pProgSink)
        {
            hr = _pDwnInfo->AddProgSink(pProgSink);
            if(hr)
            {
                goto Cleanup;
            }
        }

        if(_pProgSink)
        {
            _pDwnInfo->DelProgSink(_pProgSink);
        }

    }

    ReplaceInterface(&_pProgSink, pProgSink);

Cleanup:
    LeaveCriticalSection();
    RRETURN(hr);
}

ULONG CDwnCtx::GetState(BOOL fClear)
{
    DWORD dwState;

    EnterCriticalSection();

    dwState = _wChg;

    if(_pDwnInfo)
    {
        dwState |= _pDwnInfo->GetFlags(DWNF_STATE);
    }

    if(fClear)
    {
        _wChg = 0;
    }

    LeaveCriticalSection();

    return dwState;
}

void CDwnCtx::Signal(WORD wChg)
{
    Assert(EnteredCriticalSection());

    wChg &= _wChgReq;   // Only light up requested bits
    wChg &= ~_wChg;     // Don't light up bits already on

    if(wChg)
    {
        _wChg |= (WORD)wChg;
        super::Signal();
    }
}

void CDwnCtx::SetLoad(BOOL fLoad, DWNLOADINFO* pdli, BOOL fReload)
{
    if(!!fLoad!=!!_fLoad || (fLoad&&_fLoad&&fReload))
    {
        if(fLoad 
            && !pdli->pDwnBindData 
            && !pdli->pmk 
            && !pdli->pstm 
            && !pdli->pchUrl
            && !pdli->fClientData)
        {
            pdli->pchUrl = _pDwnInfo->GetUrl();
        }

        _pDwnInfo->SetLoad(this, fLoad, fReload, pdli);
    }
}

CDwnLoad* CDwnCtx::GetDwnLoad()
{
    CDwnLoad* pDwnLoadRet = NULL;

    EnterCriticalSection();
    if(_pDwnInfo && _pDwnInfo->_pDwnLoad)
    {
        pDwnLoadRet = _pDwnInfo->_pDwnLoad;
        pDwnLoadRet->AddRef();
    }
    LeaveCriticalSection();

    return pDwnLoadRet;
}

HRESULT NewDwnCtx(UINT dt, BOOL fLoad, DWNLOADINFO* pdli, CDwnCtx** ppDwnCtx)
{
    CDwnInfo*   pDwnInfo;
    CDwnCtx*    pDwnCtx;
    HRESULT     hr;

    hr = CDwnInfo::Create(dt, pdli, &pDwnInfo);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pDwnInfo->NewDwnCtx(&pDwnCtx);

    if(hr == S_OK)
    {
        pDwnInfo->AddDwnCtx(pDwnCtx);        
    }

    pDwnInfo->Release();

    if(hr)
    {
        goto Cleanup;
    }

    if(fLoad)
    {
        pDwnCtx->SetLoad(TRUE, pdli, FALSE);
    }

    *ppDwnCtx = pDwnCtx;

Cleanup:
    RRETURN(hr);
}