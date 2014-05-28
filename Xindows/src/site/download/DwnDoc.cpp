
#include "stdafx.h"
#include "Download.h"

#include "Dwn.h"

CDwnDoc::CDwnDoc()
{
}

CDwnDoc::~CDwnDoc()
{
    Assert(_pDoc == NULL);

    if(_aryDwnDocInfo.Size())
    {
        OnDocThreadCallback();
    }

    delete[] _ape;
    delete[] _pbRequestHeaders;

    ReleaseInterface(_pDownloadNotify);
}

void CDwnDoc::SetDoc(CDocument* pDoc)
{
    Assert(_pDoc == NULL);
    Assert(pDoc->_dwTID == GetCurrentThreadId());

    SetCallback(OnDocThreadCallback, this);

    _fCallbacks  = TRUE;
    _dwThreadId  = GetCurrentThreadId();
    _pDoc        = pDoc;
    _pDoc->SubAddRef();

    OnDocThreadCallback();
}

void CDwnDoc::Disconnect()
{
    CDocument* pDoc = _pDoc;

    if(pDoc)
    {
        Assert(IsDocThread());

        super::Disconnect();

        g_csDwnDoc.Enter();

        _pDoc = NULL;
        _fCallbacks = FALSE;

        g_csDwnDoc.Leave();

        if(_aryDwnDocInfo.Size())
        {
            OnDocThreadCallback();
            _aryDwnDocInfo.DeleteAll();
        }

        pDoc->SubRelease();
    }
}

HRESULT CDwnDoc::AddDocThreadCallback(CDwnBindData* pDwnBindData, void* pvArg)
{
    HRESULT hr = S_OK;

    if(IsDocThread())
    {
        pDwnBindData->OnDwnDocCallback(pvArg);
    }
    else
    {
        BOOL fSignal = FALSE;

        g_csDwnDoc.Enter();

        if(_fCallbacks)
        {
            // this was initialized as ddi = { pDwnBindData, pvArg }; 
            // that produces some bogus code for win16 and it generates an
            // extra data segment. so changed to the below - vreddy -7/30/97.
            DWNDOCINFO ddi;
            ddi.pDwnBindData = pDwnBindData;
            ddi.pvArg = pvArg;

            hr = _aryDwnDocInfo.AppendIndirect(&ddi);

            if(hr == S_OK)
            {
                pDwnBindData->SubAddRef();
                fSignal = TRUE;
            }
        }

        g_csDwnDoc.Leave();

        if(fSignal)
        {
            super::Signal();
        }
    }
    RRETURN(hr);
}

void CDwnDoc::OnDocThreadCallback()
{
    DWNDOCINFO ddi;

    while(_aryDwnDocInfo.Size() > 0)
    {
        ddi.pDwnBindData = NULL;
        ddi.pvArg = NULL;

        g_csDwnDoc.Enter();

        if(_aryDwnDocInfo.Size() > 0)
        {
            ddi = _aryDwnDocInfo[0];
            _aryDwnDocInfo.Delete(0);
        }

        g_csDwnDoc.Leave();

        if(ddi.pDwnBindData)
        {
            ddi.pDwnBindData->OnDwnDocCallback(ddi.pvArg);
            ddi.pDwnBindData->SubRelease();
        }
    }
}

HRESULT CDwnDoc::QueryService(BOOL fBindOnApt, REFGUID rguid, REFIID riid, void** ppvObj)
{
    HRESULT hr;

    if((rguid==IID_IAuthenticate || rguid==IID_IWindowForBindingUI) && (rguid==riid))
    {
        hr = QueryInterface(rguid, ppvObj);
    }
    else if(fBindOnApt && IsDocThread() && _pDoc)
    {
        hr = _pDoc->QueryService(rguid, riid, ppvObj);
    }
    else
    {
        *ppvObj = NULL;
        hr = E_NOINTERFACE;
    }

    return hr;
}

HRESULT CDwnDoc::Load(IStream* pstm, CDwnDoc** ppDwnDoc)
{
    CDwnDoc*    pDwnDoc = NULL;
    BYTE        bIsNull;
    HRESULT     hr;

    hr = pstm->Read(&bIsNull, sizeof(BYTE), NULL);

    if(hr || bIsNull)
    {
        goto Cleanup;
    }

    pDwnDoc = new CDwnDoc;

    if(pDwnDoc == NULL)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    hr = pstm->Read(&pDwnDoc->_dwBindf, 5*sizeof(DWORD), NULL);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pDwnDoc->_cstrAcceptLang.Load(pstm);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pDwnDoc->_cstrUserAgent.Load(pstm);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    if(hr && pDwnDoc)
    {
        pDwnDoc->Release();
        *ppDwnDoc = NULL;
    }
    else
    {
        *ppDwnDoc = pDwnDoc;
    }

    RRETURN(hr);
}

HRESULT CDwnDoc::Save(CDwnDoc* pDwnDoc, IStream* pstm)
{
    BOOL    bIsNull = pDwnDoc == NULL;
    HRESULT hr;

    hr = pstm->Write(&bIsNull, sizeof(BYTE), NULL);
    if(hr)
    {
        goto Cleanup;
    }

    if(pDwnDoc)
    {
        hr = pstm->Write(&pDwnDoc->_dwBindf, 5*sizeof(DWORD), NULL);
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDwnDoc->_cstrAcceptLang.Save(pstm);
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDwnDoc->_cstrUserAgent.Save(pstm);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

ULONG CDwnDoc::GetSaveSize(CDwnDoc* pDwnDoc)
{
    ULONG cb = sizeof(BYTE);

    if(pDwnDoc)
    {
        cb += 5 * sizeof(DWORD);
        cb += pDwnDoc->_cstrAcceptLang.GetSaveSize();
        cb += pDwnDoc->_cstrUserAgent.GetSaveSize();
    }

    return cb;
}

HRESULT CDwnDoc::SetString(CString* pcstr, LPCTSTR psz)
{
    if(psz && *psz)
    {
        RRETURN(pcstr->Set(psz));
    }
    else
    {
        pcstr->Free();
        return(S_OK);
    }
}

void CDwnDoc::TakeRequestHeaders(BYTE** ppb, ULONG* pcb)
{
    Assert(!_pbRequestHeaders);

    _pbRequestHeaders = *ppb;
    _cbRequestHeaders = *pcb;

    *ppb = NULL;
    *pcb = 0;
}

void CDwnDoc::SetDownloadNotify(IDownloadNotify* pDownloadNotify)
{
    Assert(_dwThreadId == GetCurrentThreadId());
    Assert(!_pDownloadNotify);

    if(pDownloadNotify)
    {
        pDownloadNotify->AddRef();
    }
    _pDownloadNotify = pDownloadNotify;
}

HRESULT CDwnDoc::SetAuthorColors(LPCTSTR pchColors, int cchColors)
{
    if(_fGotAuthorPalette)
    {
        Trace0("Ignoring author palette\n");
        RRETURN(S_OK);
    }

    Trace0("Setting author colors\n");

    HRESULT hr = S_OK;
    if(cchColors == -1)
    {
        cchColors = _tcslen(pchColors);
    }

    LPCTSTR pch = pchColors;
    LPCTSTR pchTok = pchColors;
    LPCTSTR pchEnd = pchColors + cchColors;

    PALETTEENTRY ape[256];

    unsigned cpe = 0;
    CColorValue cv;

    while((pch<pchEnd) && (cpe<256))
    {
        while(*pch && isspace(*pch))
        {
            pch++;
        }

        pchTok = pch;
        BOOL fParen = FALSE;

        while(pch<pchEnd && (fParen || !isspace(*pch)))
        {
            if(*pch == _T('('))
            {
                if(fParen)
                {
                    hr = E_INVALIDARG;
                    goto Cleanup;
                }
                else
                {
                    fParen = TRUE;
                }
            }
            else if(*pch == _T(')'))
            {
                if(fParen)
                {
                    fParen = FALSE;
                }
                else
                {
                    hr = E_INVALIDARG;
                    goto Cleanup;
                }
            }
            pch++;
        }

        int iStrLen = pch - pchTok;

        if(iStrLen > 0)
        {
            hr = cv.FromString(pchTok, FALSE, iStrLen);

            if(FAILED(hr))
            {
                goto Cleanup;
            }

            COLORREF cr = cv.GetColorRef();
            ape[cpe].peRed = GetRValue(cr);
            ape[cpe].peGreen = GetGValue(cr);
            ape[cpe].peBlue = GetBValue(cr);
            ape[cpe].peFlags = 0;

            cpe++;
        }
    }

    if(cpe)
    {
        Assert(!_ape);

        _ape = (PALETTEENTRY*)MemAlloc(cpe*sizeof(PALETTEENTRY));

        if(!_ape)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        memcpy(_ape, ape, cpe*sizeof(PALETTEENTRY));

        _cpe = cpe;
    }

Cleanup:
    Trace((SUCCEEDED(hr) ? _T("Author palette: %d colors\n") : _T("Author palette failed\n"), _cpe));
    // No matter what happened, we are no longer interested in getting the palette
    PreventAuthorPalette();

    RRETURN(hr);
}

HRESULT CDwnDoc::GetColors(CColorInfo* pCI)
{
    HRESULT hr = S_FALSE;

    // This is thread safe because by the time _cpe is set, _ape has already been allocated.
    // If for some reason we arrive a bit too soon, we'll get it next time.
    if(_cpe)
    {
        Assert(_ape);

        hr = pCI->AddColors(_cpe, _ape);
    }

    RRETURN1(hr, S_FALSE);
}

STDMETHODIMP CDwnDoc::Authenticate(HWND* phwnd, LPWSTR* ppszUsername, LPWSTR* ppszPassword)
{
    HRESULT hr;

    *phwnd = NULL;
    *ppszUsername = NULL;
    *ppszPassword = NULL;

    if(IsDocThread() && _pDoc)
    {
        IAuthenticate* pAuth;

        hr = _pDoc->QueryService(IID_IAuthenticate, IID_IAuthenticate, (void**)&pAuth);

        if(hr == S_OK)
        {
            hr = pAuth->Authenticate(phwnd, ppszUsername, ppszPassword);
            pAuth->Release();
        }
    }
    else
    {
        // Either we are on the wrong thread or the document has disconnected.  In either case,
        // we can no longer provide this service.
        hr = E_FAIL;
    }

    RRETURN(hr);
}

STDMETHODIMP CDwnDoc::GetWindow(REFGUID rguidReason, HWND* phwnd)
{
    HRESULT hr;

    *phwnd = NULL;

    if(IsDocThread() && _pDoc)
    {
        IWindowForBindingUI* pwfbu;

        hr = _pDoc->QueryService(IID_IWindowForBindingUI, IID_IWindowForBindingUI, (void**)&pwfbu);

        if(hr == S_OK)
        {
            hr = pwfbu->GetWindow(rguidReason, phwnd);
            pwfbu->Release();
        }
    }
    else
    {
        // Either we are on the wrong thread or the document has disconnected.  In either case,
        // we can no longer provide this service.
        hr = E_FAIL;
    }

    RRETURN(hr);
}

STDMETHODIMP CDwnDoc::QueryInterface(REFIID iid, LPVOID* ppv)
{
    if(iid==IID_IUnknown || iid==IID_IAuthenticate)
    {
        *ppv = (IAuthenticate*)this;
    }
    else if(iid == IID_IWindowForBindingUI)
    {
        *ppv = (IWindowForBindingUI*)this;
    }
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    ((IUnknown*)*ppv)->AddRef();
    return S_OK;
}
