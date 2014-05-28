
#include "stdafx.h"
#include "Download.h"

#include "Dwn.h"

static const GUID IID_IDwnBindInfo = 
{ 0x3050f3c3, 0x98b5, 0x11cf, { 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b } };
static const GUID CLSID_CDwnBindInfo = 
{ 0x3050f3c2, 0x98b5, 0x11cf, { 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b } };

// Globals --------------------------------------------------------------------
CDwnBindInfo* g_pDwnBindTrace = NULL;

// Definitions ----------------------------------------------------------------
#define Align64(n)      (((n)+63)&~63)

// Utilities ------------------------------------------------------------------
BOOL GetFileLastModTime(TCHAR* pchFile, FILETIME* pftLastMod)
{
    WIN32_FIND_DATA fd;
    HANDLE hFF = FindFirstFile(pchFile, &fd);

    if(hFF != INVALID_HANDLE_VALUE)
    {
        *pftLastMod = fd.ftLastWriteTime;
        FindClose(hFF);
        return(TRUE);
    }

    return FALSE;
}

BOOL GetUrlLastModTime(TCHAR* pchUrl, UINT uScheme, DWORD dwBindf, FILETIME* pftLastMod)
{
    BOOL    fRet = FALSE;
    HRESULT hr;

    Assert(uScheme == GetUrlScheme(pchUrl));

    if(uScheme == URL_SCHEME_FILE)
    {
        TCHAR achPath[MAX_PATH];
        DWORD cchPath;

        hr = CoInternetParseUrl(pchUrl, PARSE_PATH_FROM_URL, 0,
            achPath, ARRAYSIZE(achPath), &cchPath, 0);

        if(hr == S_OK)
        {
            fRet = GetFileLastModTime(achPath, pftLastMod);
        }
    }
    else if(uScheme==URL_SCHEME_HTTP || uScheme==URL_SCHEME_HTTPS)
    {
        fRet = !IsUrlCacheEntryExpired(pchUrl, dwBindf&BINDF_FWD_BACK, pftLastMod)
            && pftLastMod->dwLowDateTime
            && pftLastMod->dwHighDateTime;
    }

    return fRet;
}

// CDwnBindInfo ---------------------------------------------------------------
CDwnBindInfo::CDwnBindInfo()
{
#ifdef _DEBUG
    if(g_pDwnBindTrace == NULL)
    {
        g_pDwnBindTrace = this;
        Trace2("DwnBindInfo [%lX] Construct %d\n", this, GetRefs());
    }
#endif
}

CDwnBindInfo::~CDwnBindInfo()
{
#ifdef _DEBUG
    if(g_pDwnBindTrace==this)
    {
        g_pDwnBindTrace = NULL;
        Trace1("DwnBindInfo [%lX] Destruct\n", this);
    }
#endif

    if(_pDwnDoc)
    {
        _pDwnDoc->Release();
    }
}

// CDwnBindInfo (IUnknown) --------------------------------------------------------
STDMETHODIMP CDwnBindInfo::QueryInterface(REFIID iid, void** ppv)
{
    Assert(CheckThread());

    if(iid==IID_IUnknown || iid==IID_IBindStatusCallback)
    {
        *ppv = (IBindStatusCallback *)this;
    }
    else if(iid == IID_IServiceProvider)
    {
        *ppv = (IServiceProvider *)this;
    }
    else if(iid == IID_IHttpNegotiate)
    {
        *ppv = (IHttpNegotiate *)this;
    }
    else if(iid == IID_IMarshal)
    {
        *ppv = (IMarshal *)this;
    }
    else if(iid == IID_IInternetBindInfo)
    {
        *ppv = (IInternetBindInfo *)this;
    }
    else if(iid == IID_IDwnBindInfo)
    {
        *ppv = this;
        AddRef();
        return S_OK;
    }
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    ((IUnknown*)*ppv)->AddRef();
    return S_OK;
}

STDMETHODIMP_(ULONG) CDwnBindInfo::AddRef()
{
    ULONG ulRefs = super::AddRef();

#ifdef _DEBUG
    if(this == g_pDwnBindTrace)
    {
        Trace3("[%lX] DwnBindInfo %lX AR  %ld\n", GetCurrentThreadId(), this, ulRefs);
    }
#endif

    return ulRefs;
}

STDMETHODIMP_(ULONG) CDwnBindInfo::Release()
{
    ULONG ulRefs = super::Release();

#ifdef _DEBUG
    if(this == g_pDwnBindTrace)
    {
        Trace3("[%lX] DwnBindInfo %lX Rel %ld", GetCurrentThreadId(), this, ulRefs);
    }
#endif

    return ulRefs;
}

void CDwnBindInfo::SetDwnDoc(CDwnDoc* pDwnDoc)
{
    if(_pDwnDoc)
    {
        _pDwnDoc->Release();
    }

    _pDwnDoc = pDwnDoc;

    if(_pDwnDoc)
    {
        _pDwnDoc->AddRef();
    }
}

UINT CDwnBindInfo::GetScheme()
{
    return URL_SCHEME_UNKNOWN;
}

// CDwnBindInfo (IBindStatusCallback) -----------------------------------------
STDMETHODIMP CDwnBindInfo::OnStartBinding(DWORD grfBSCOption, IBinding* pbinding)
{
    Assert(CheckThread());
    return S_OK;
}

STDMETHODIMP CDwnBindInfo::GetPriority(LONG* pnPriority)
{
    Assert(CheckThread());
    *pnPriority = NORMAL_PRIORITY_CLASS;
    return S_OK;
}

STDMETHODIMP CDwnBindInfo::OnLowResource(DWORD dwReserved)
{
    Assert(CheckThread());
    return S_OK;
}

STDMETHODIMP CDwnBindInfo::OnProgress(ULONG ulPos, ULONG ulMax, ULONG ulCode, LPCWSTR pszText)
{
    Assert(CheckThread());

    if(pszText && (ulCode==BINDSTATUS_MIMETYPEAVAILABLE || ulCode==BINDSTATUS_RAWMIMETYPE))
    {
        _cstrContentType.Set(pszText);
    }

    if(pszText && (ulCode==BINDSTATUS_CACHEFILENAMEAVAILABLE))
    {
        _cstrCacheFilename.Set(pszText);
    }

    return S_OK;
}

STDMETHODIMP CDwnBindInfo::OnStopBinding(HRESULT hrReason, LPCWSTR szReason)
{
    Assert(CheckThread());
    return S_OK;
}

STDMETHODIMP CDwnBindInfo::GetBindInfo(DWORD* pdwBindf, BINDINFO* pbindinfo)
{
    Assert(CheckThread());

    HRESULT hr;

    if(pbindinfo->cbSize != sizeof(BINDINFO))
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    memset(pbindinfo, 0, sizeof(BINDINFO));

    pbindinfo->cbSize = sizeof(BINDINFO);

    *pdwBindf = BINDF_ASYNCHRONOUS|BINDF_ASYNCSTORAGE|BINDF_PULLDATA;

    if(_pDwnDoc)
    {
        if(_fIsDocBind)
        {
            *pdwBindf |= _pDwnDoc->GetDocBindf();
            pbindinfo->dwCodePage = _pDwnDoc->GetURLCodePage();
        }
        else
        {
            *pdwBindf |= _pDwnDoc->GetBindf();
            pbindinfo->dwCodePage = _pDwnDoc->GetDocCodePage();
        }

        if(_pDwnDoc->GetLoadf() & DLCTL_URL_ENCODING_DISABLE_UTF8)
        {
            pbindinfo->dwOptions = BINDINFO_OPTIONS_DISABLE_UTF8;
        }
        else if(_pDwnDoc->GetLoadf() & DLCTL_URL_ENCODING_ENABLE_UTF8)
        {
            pbindinfo->dwOptions = BINDINFO_OPTIONS_ENABLE_UTF8;
        }
    }

#ifdef _DEBUG
    *pdwBindf |= BINDF_NOWRITECACHE;
    *pdwBindf |= BINDF_GETNEWESTVERSION;
    *pdwBindf &= ~BINDF_PULLDATA;
#endif

    {
        // If a GET method for a form, always hit the server
        if(!(*pdwBindf&BINDF_OFFLINEOPERATION) && (*pdwBindf&BINDF_FORMS_SUBMIT))
        {
            *pdwBindf &= ~(BINDF_GETNEWESTVERSION | BINDF_PRAGMA_NO_CACHE);
            *pdwBindf |= BINDF_RESYNCHRONIZE;
        }
        pbindinfo->dwBindVerb = BINDVERB_GET;
    }

    hr = S_OK;

Cleanup:
    RRETURN(hr);
}

STDMETHODIMP CDwnBindInfo::OnDataAvailable(DWORD grfBSCF, DWORD dwSize, FORMATETC* pformatetc, STGMEDIUM* pstgmed)
{
    Assert(CheckThread());
    return S_OK;
}

STDMETHODIMP CDwnBindInfo::OnObjectAvailable(REFIID riid, IUnknown* punk)
{
    Assert(CheckThread());
    return S_OK;
}

// CDwnBindInfo (IInternetBindInfo) -------------------------------------------
STDMETHODIMP CDwnBindInfo::GetBindString(ULONG ulStringType, LPOLESTR* ppwzStr, ULONG cEl, ULONG* pcElFetched)
{
    HRESULT hr = S_OK;

    *pcElFetched = 0;

    switch(ulStringType)
    {
    case BINDSTRING_ACCEPT_MIMES:
        {
            if(cEl >= 1)
            {
                ppwzStr[0] = (LPOLESTR)CoTaskMemAlloc(4*sizeof(TCHAR));

                if(ppwzStr[0] == 0)
                {
                    hr = E_OUTOFMEMORY;
                    goto Cleanup;
                }

                memcpy(ppwzStr[0], _T("*/*"), 4*sizeof(TCHAR));
                *pcElFetched = 1;
            }
        }
        break;

    case BINDSTRING_POST_COOKIE:
        {
        }
        break;

    case BINDSTRING_POST_DATA_MIME:
        {
        }
        break;
    }

Cleanup:
    RRETURN(hr);
}

// CDwnBindInfo (IServiceProvider) --------------------------------------------
STDMETHODIMP CDwnBindInfo::QueryService(REFGUID rguidService, REFIID riid, void** ppvObj)
{
    Assert(CheckThread());

    HRESULT hr;

    if(rguidService == IID_IHttpNegotiate)
    {
        hr = QueryInterface(riid, ppvObj);
    }
    else if(rguidService == IID_IInternetBindInfo)
    {
        hr = QueryInterface(riid, ppvObj);
    }
    else if(_pDwnDoc)
    {
        hr = _pDwnDoc->QueryService(IsBindOnApt(), rguidService, riid, ppvObj);
    }
    else
    {
        *ppvObj = NULL;
        hr = E_NOINTERFACE;
    }

    RRETURN(hr);
}

// CDwnBindInfo (IHttpNegotiate) ----------------------------------------------
STDMETHODIMP CDwnBindInfo::BeginningTransaction(LPCWSTR pszUrl, LPCWSTR pszHeaders, DWORD dwReserved, LPWSTR* ppszHeaders)
{
    Assert(CheckThread());

    LPCTSTR     apch[16];
    UINT        acch[16];
    LPCTSTR*    ppch = apch;
    UINT*       pcch = acch;
    HRESULT     hr   = S_OK;

    // If we have been told the exact http headers to use, use them now
    if(_fIsDocBind && _pDwnDoc && _pDwnDoc->GetRequestHeaders())
    {
        TCHAR*  pch;
        UINT    cch = _pDwnDoc->GetRequestHeadersLength();

        pch = (TCHAR*)CoTaskMemAlloc(cch*sizeof(TCHAR));

        if(pch == NULL)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        // note: we don't need to convert an extra zero terminator, so cch-1
        AnsiToWideTrivial((char*)_pDwnDoc->GetRequestHeaders(), pch, cch-1);

        *ppszHeaders = pch;

        goto Cleanup;
    }

    // Otherwise, assemble the http headers
    *ppszHeaders = NULL;

    if(ppch > apch)
    {
        LPCTSTR*    ppchEnd = ppch;
        TCHAR*      pch;
        UINT        cch = 0;

        for(; ppch>apch; --ppch)
        {
            cch += *--pcch;
        }

        pch = (TCHAR*)CoTaskMemAlloc((cch+1) * sizeof(TCHAR));

        if(pch == NULL)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        *ppszHeaders = pch;

        for(; ppch<ppchEnd; pch+=*pcch++,ppch++)
        {
            memcpy(pch, *ppch, *pcch*sizeof(TCHAR));
        }

        *pch = 0;

        Assert((UINT)(pch-*ppszHeaders) == cch);
    }

Cleanup:
    RRETURN(hr);
}

STDMETHODIMP CDwnBindInfo::OnResponse(
          DWORD dwResponseCode,
          LPCWSTR szResponseHeaders,
          LPCWSTR szRequestHeaders,
          LPWSTR* ppszAdditionalRequestHeaders)
{
    Assert(CheckThread());
    return S_OK;
}

// CDwnBindInfo (Internal) ----------------------------------------------------
HRESULT CreateDwnBindInfo(IUnknown* pUnkOuter, IUnknown** ppUnk)
{
    *ppUnk = NULL;

    if(pUnkOuter != NULL)
    {
        RRETURN(CLASS_E_NOAGGREGATION);
    }

    CDwnBindInfo* pDwnBindInfo = new CDwnBindInfo;

    if(pDwnBindInfo == NULL)
    {
        RRETURN(E_OUTOFMEMORY);
    }

    *ppUnk = (IBindStatusCallback*)pDwnBindInfo;

    return S_OK;
}

// CDwnBindData (IUnknown) ----------------------------------------------------
STDMETHODIMP CDwnBindData::QueryInterface(REFIID iid, void **ppv)
{
    Assert(CheckThread());

    *ppv = NULL;

    if(iid == IID_IInternetBindInfo)
    {
        *ppv = (IInternetBindInfo *)this;
    }
    else if(iid == IID_IInternetProtocolSink)
    {
        *ppv = (IInternetProtocolSink *)this;
    }
    else
    {
        return super::QueryInterface(iid, ppv);
    }

    ((IUnknown*)*ppv)->AddRef();
    return S_OK;
}

STDMETHODIMP_(ULONG) CDwnBindData::AddRef()
{
    return super::AddRef();
}

STDMETHODIMP_(ULONG) CDwnBindData::Release()
{
    return super::Release();
}

// CDwnBindData (Internal) ----------------------------------------------------
CDwnBindData::~CDwnBindData()
{
    Assert(!_fBindOnApt || _u.pts==NULL);

    if(_pDwnStm)
    {
        _pDwnStm->Release();
    }

    MemFree(_pbPeek);

    if(!_fBindOnApt && _o.pInetProt)
    {
        _o.pInetProt->Release();
    }

    if(_hLock)
    {
        InternetUnlockRequestFile(_hLock);
    }
}

void CDwnBindData::Passivate()
{
    if(!_fBindTerm)
    {
        Terminate(E_ABORT);
    }

    super::Passivate();
}

void CDwnBindData::Terminate(HRESULT hrErr)
{
    if(_fBindTerm)
    {
        return;
    }

    BOOL fTerminate = FALSE;

    g_csDwnBindTerm.Enter();

    if(_fBindTerm)
    {
        goto Cleanup;
    }

    if(_hrErr == S_OK)
    {
        _hrErr = hrErr;
    }

    if(_fBindOnApt)
    {
        if(!_u.pstm
            && !_u.punkForRel
            && !_u.pbc
            && !_u.pbinding
            && !_u.pts)
        {
            goto Cleanup;
        }

        if(_u.dwTid != GetCurrentThreadId())
        {
            if(!_u.pts || _u.fTermPosted)
            {
                goto Cleanup;
            }

            // We're not on the apartment thread, so we can't access
            // the objects we're binding right now.  Post a callback
            // on the apartment thread.

            // SubAddRef and set flags before posting the message to avoid race
            SubAddRef();
            _u.fTermPosted = TRUE;
            _u.fTermReceived = FALSE;

            HRESULT hr = GWPostMethodCallEx(_u.pts, this,
                ONCALL_METHOD(CDwnBindData, TerminateOnApt, terminateonapt), 0, FALSE, "CDwnBindData::TerminateOnApt");

            if(hr)
            {
                _u.fTermReceived = TRUE;
                SubRelease();
            }

            goto Cleanup;
        }
    }

    fTerminate = TRUE;

Cleanup:
    g_csDwnBindTerm.Leave();

    if(fTerminate)
    {
        TerminateBind();
    }
}

void STDMETHODCALLTYPE CDwnBindData::TerminateOnApt(DWORD_PTR dwContext)
{
    Assert(!_u.fTermReceived);

    _u.fTermReceived = TRUE;
    TerminateBind();
    SubRelease();
}

void CDwnBindData::TerminateBind()
{
    Assert(CheckThread());

    SubAddRef();

    g_csDwnBindTerm.Enter();

    if(_pDwnStm && !_pDwnStm->IsEofWritten())
    {
        _pDwnStm->WriteEof(_hrErr);
    }

    SetEof();

    if(_fBindOnApt)
    {
        IUnknown*   punkRel1 = _u.pstm;
        IUnknown*   punkRel2 = _u.punkForRel;
        IBindCtx*   pbc      = NULL;
        IBinding*   pbinding = NULL;
        BOOL        fDoAbort = FALSE;
        BOOL        fRelTls  = FALSE;

        _u.pstm              = NULL;
        _u.punkForRel        = NULL;

        if(_u.pts && _u.fTermPosted && !_u.fTermReceived)
        {
            // We've posted a method call to TerminateOnApt which hasn't been received
            // yet.  This happens if Terminate gets called on the apartment thread
            // before messages are pumped.  To keep reference counts happy, we need
            // to simulate the receipt of the method call here by first killing any
            // posted method call and then undoing the SubAddRef.
            GWKillMethodCallEx(_u.pts, this,
                ONCALL_METHOD(CDwnBindData, TerminateOnApt, terminateonapt), 0);

            // note: no danger of post/kill/set-flag race because
            // we're protected by g_csDwnBindTerm
            _u.fTermReceived = TRUE;
            SubRelease();
        }

        if(_fBindDone)
        {
            pbc         = _u.pbc;
            _u.pbc      = NULL;
            pbinding    = _u.pbinding;
            _u.pbinding = NULL;
            fRelTls     = !!_u.pts;
            _u.pts      = NULL;
            _fBindTerm  = TRUE;
        }
        else if(!_fBindAbort && _u.pbinding)
        {
            pbinding    = _u.pbinding;
            fDoAbort    = TRUE;
            _fBindAbort = TRUE;
        }

        g_csDwnBindTerm.Leave();

        ReleaseInterface(punkRel1);
        ReleaseInterface(punkRel2);

        if(fDoAbort)
        {
            pbinding->Abort();
        }
        else
        {
            ReleaseInterface(pbinding);
        }

        if(pbc)
        {
            RevokeBindStatusCallback(pbc, this);
            ReleaseInterface(pbc);
        }

        if(fRelTls)
        {
            ReleaseThreadState();
        }
    }
    else if(_o.pInetProt)
    {
        BOOL fDoTerm    = _fBindDone && !_fBindTerm;
        BOOL fDoAbort   = !_fBindDone && !_fBindAbort;

        if(_fBindDone)
        {
            _fBindTerm = TRUE;
        }
        else
        {
            _fBindAbort = TRUE;
        }

        g_csDwnBindTerm.Leave();

        if(fDoAbort)
        {
            _o.pInetProt->Abort(E_ABORT, 0);
        }
        else if(fDoTerm)
        {
            _o.pInetProt->Terminate(0);
        }
    }
    else
    {
        _fBindTerm = TRUE;
        g_csDwnBindTerm.Leave();
    }

    SubRelease();
}

HRESULT CDwnBindData::SetBindOnApt()
{
    HRESULT hr;

    if(!(_dwFlags & DWNF_NOAUTOBUFFER))
    {
        _pDwnStm = new CDwnStm;

        if(_pDwnStm == NULL)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }
    }

    hr = AddRefThreadState();
    if(hr)
    {
        goto Cleanup;
    }

    _u.pts      = GetThreadState();
    _u.dwTid    = GetCurrentThreadId();
    _fBindDone  = TRUE;
    _fBindOnApt = TRUE;

Cleanup:
    RRETURN(hr);
}

#ifdef _DEBUG
BOOL CDwnBindData::CheckThread()
{
    return (!_fBindOnApt || _u.dwTid==GetCurrentThreadId());
}
#endif

void CDwnBindData::OnDwnDocCallback(void* pvArg)
{
    Assert(!_fBindOnApt);

    if(_o.pInetProt)
    {
        HRESULT hr = _o.pInetProt->Continue((PROTOCOLDATA*)pvArg);

        if(hr)
        {
            SignalDone(hr);
        }
    }
}

void CDwnBindData::SetEof()
{
    g_csDwnBindPend.Enter();

    _fPending = FALSE;
    _fEof = TRUE;

    g_csDwnBindPend.Leave();
}

void CDwnBindData::SetPending(BOOL fPending)
{
    g_csDwnBindPend.Enter();

    if(!_fEof)
    {
        _fPending = fPending;
    }

    g_csDwnBindPend.Leave();
}

// CDwnBindData (Binding) -----------------------------------------------------
void CDwnBindData::Bind(DWNLOADINFO* pdli, DWORD dwFlagsExtra)
{
    LPTSTR      pchAlloc = NULL;
    IMoniker*   pmkAlloc = NULL;
    IBindCtx*   pbcAlloc = NULL;
    IStream*    pstm     = NULL;
    CDwnDoc*    pDwnDoc  = pdli->pDwnDoc;
    LPCTSTR     pchUrl;
    IMoniker*   pmk;
    IBindCtx*   pbc;
    HRESULT     hr;

    _dwFlags = dwFlagsExtra;

    if(pDwnDoc)
    {
        _dwFlags |= pDwnDoc->GetDownf();
    }

    if(_dwFlags & DWNF_ISDOCBIND)
    {
        SetIsDocBind();
    }

    // Case 1: Binding to an user-supplied IStream
    if(pdli->pstm)
    {
        Assert(!pdli->fForceInet);

        _dwSecFlags = (!pdli->fUnsecureSource && IsUrlSecure(pdli->pchUrl)) ? SECURITY_FLAG_SECURE : 0;

        hr = SetBindOnApt();
        if(hr)
        {
            goto Cleanup;
        }

        ReplaceInterface(&_u.pstm, pdli->pstm);
        BufferData();
        SignalData();
        goto Cleanup;
    }

    // Case 2: Not binding to a moniker or URL.  This is a manual binding
    // where data will be provided externally and shunted to the consumer.
    // Actual configuration of buffering will occur outside this function.
    if(pdli->fClientData || (!pdli->pmk && !pdli->pchUrl))
    {
        _dwSecFlags = (!pdli->fUnsecureSource && IsUrlSecure(pdli->pchUrl)) ? SECURITY_FLAG_SECURE : 0;
        hr = S_OK;
        goto Cleanup;
    }

    // Case 3: Binding asynchronously with IInternetSession
    pchUrl = pdli->pchUrl;

    if(!pchUrl && IsAsyncMoniker(pdli->pmk)==S_OK)
    {
        hr = pdli->pmk->GetDisplayName(NULL, NULL, &pchAlloc);
        if(hr)
        {
            goto Cleanup;
        }

        pchUrl = pchAlloc;
    }

    // Since INTERNET_OPTION_SECURITY_FLAGS doesn't work for non-wininet URLs
    _uScheme = pchUrl ? GetUrlScheme(pchUrl) : URL_SCHEME_UNKNOWN;

    // Check _uScheme first before calling IsUrlSecure to avoid second (slow) GetUrlScheme call
    _dwSecFlags = (_uScheme!=URL_SCHEME_HTTP &&
        _uScheme!=URL_SCHEME_FILE &&
        !pdli->fUnsecureSource &&
        IsUrlSecure(pchUrl)) ? SECURITY_FLAG_SECURE : 0;

    if(_uScheme==URL_SCHEME_FILE
        && (_dwFlags&(DWNF_GETFILELOCK|DWNF_GETMODTIME)))
    {
        TCHAR achPath[MAX_PATH];
        DWORD cchPath;

        hr = CoInternetParseUrl(pchUrl, PARSE_PATH_FROM_URL, 0,
            achPath, ARRAYSIZE(achPath), &cchPath, 0);
        if(hr)
        {
            goto Cleanup;
        }

        hr = _cstrFile.Set(achPath);
        if(hr)
        {
            goto Cleanup;
        }

        if(_cstrFile && (_dwFlags&DWNF_GETMODTIME))
        {
            GetFileLastModTime(_cstrFile, &_ftLastMod);
        }
    }

    if(pchUrl && pdli->pInetSess && !pdli->pbc
        && (_uScheme==URL_SCHEME_FILE
        || _uScheme==URL_SCHEME_HTTP
        || _uScheme==URL_SCHEME_HTTPS))
    {
        hr = pdli->pInetSess->CreateBinding(NULL, pchUrl, NULL,
            NULL, &_o.pInetProt, OIBDG_DATAONLY);
        if(hr)
        {
            goto Cleanup;
        }

        if(!_fIsDocBind)
        {
            IOInetPriority* pOInetPrio = NULL;

            _o.pInetProt->QueryInterface(IID_IInternetPriority, (void**)&pOInetPrio);

            if(pOInetPrio)
            {
                pOInetPrio->SetPriority(THREAD_PRIORITY_BELOW_NORMAL);
                pOInetPrio->Release();
            }
        }

        hr = _o.pInetProt->Start(pchUrl, this, this, PI_MIMEVERIFICATION, 0);

        // That's it.  We're off ...
        goto Cleanup;
    }

    if(pdli->fForceInet)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    // Case 4: Binding through URL moniker on apartment thread
    pmk = pdli->pmk;

    if(pmk == NULL)
    {
        hr = CreateURLMoniker(NULL, pchUrl, &pmkAlloc);
        if(hr)
        {
            goto Cleanup;
        }

        pmk = pmkAlloc;
    }

    hr = SetBindOnApt();
    if(hr)
    {
        goto Cleanup;
    }

    pbc = pdli->pbc;

    if(pbc == NULL)
    {
        hr = CreateBindCtx(0, &pbcAlloc);
        if(hr)
        {
            goto Cleanup;
        }

        pbc = pbcAlloc;
    }

    ReplaceInterface(&_u.pbc, pbc);

    {
        hr = RegisterBindStatusCallback(pbc, this, 0, 0);
        if(FAILED(hr))
        {
            goto Cleanup;
        }

        hr = pmk->BindToStorage(pbc, NULL, IID_IStream, (void**)&pstm);

        if(FAILED(hr))
        {
            RevokeBindStatusCallback(pbc, this);
            goto Cleanup;
        }
    }

    if(pstm)
    {
        ReplaceInterface(&_u.pstm, pstm);
        BufferData();
        SignalData();
    }

    hr = S_OK;

Cleanup:
    // If failed to start binding, signal done
    if(hr)
    {
        SignalDone(hr);
    }

    ReleaseInterface(pbcAlloc);
    ReleaseInterface(pmkAlloc);
    ReleaseInterface(pstm);
    CoTaskMemFree(pchAlloc);
}

// CDwnBindData (Reading) ---------------------------------------------------------
HRESULT CDwnBindData::Peek(void* pv, ULONG cb, ULONG* pcb)
{
    ULONG   cbRead;
    ULONG   cbPeek = _pbPeek ? *(ULONG*)_pbPeek : 0;
    HRESULT hr = S_OK;

    *pcb = 0;

    if(cb > cbPeek)
    {
        if(Align64(cb) > Align64(cbPeek))
        {
            hr = MemRealloc((void**)&_pbPeek, sizeof(ULONG)+Align64(cb));
            if(hr)
            {
                goto Cleanup;
            }
        }

        cbRead = 0;

        hr = ReadFromData(_pbPeek+sizeof(ULONG)+cbPeek, cb-cbPeek, &cbRead);
        if(hr)
        {
            goto Cleanup;
        }

        cbPeek += cbRead;
        *(ULONG*)_pbPeek = cbPeek;

        if(cbPeek == 0)
        {
            // We don't want the state where _pbPeek exists but has no peek
            // data.  The IsEof and IsPending functions assume this won't
            // happen.
            MemFree(_pbPeek);
            _pbPeek = NULL;
        }

    }

    if(cb > cbPeek)
    {
        cb = cbPeek;
    }

    if(cb > 0)
    {
        memcpy(pv, _pbPeek+sizeof(ULONG), cb);
    }

    *pcb = cb;

Cleanup:
    RRETURN(hr);
}

HRESULT CDwnBindData::Read(void* pv, ULONG cb, ULONG* pcb)
{
    ULONG   cbRead  = 0;
    ULONG   cbPeek  = _pbPeek ? *(ULONG*)_pbPeek : 0;
    HRESULT hr      = S_OK;

    if(cbPeek)
    {
        cbRead = (cb>cbPeek) ? cbPeek : cb;

        memcpy(pv, _pbPeek+sizeof(ULONG), cbRead);

        if(cbRead == cbPeek)
        {
            MemFree(_pbPeek);
            _pbPeek = NULL;
        }
        else
        {
            memmove(_pbPeek+sizeof(ULONG), _pbPeek+sizeof(ULONG)+cbRead, 
                cbPeek-cbRead);
            *(ULONG*)_pbPeek -= cbRead;
        }

        cb -= cbRead;
        pv = (BYTE*)pv + cbRead;
    }

    *pcb = cbRead;

    if(cb)
    {
        hr = ReadFromData(pv, cb, pcb);

        if(hr)
        {
            *pcb  = cbRead;
        }
        else
        {
            *pcb += cbRead;
        }
    }

    if(_pDwnDoc)
    {
        _pDwnDoc->AddBytesRead(*pcb);
    }

    RRETURN(hr);
}

HRESULT CDwnBindData::ReadFromData(void* pv, ULONG cb, ULONG* pcb)
{
    BOOL fBindDone = _fBindDone;
    HRESULT hr;

    if(_pDwnStm)
    {
        hr = _pDwnStm->Read(pv, cb, pcb);
    }
    else
    {
        hr = ReadFromBind(pv, cb, pcb);
    }

    if(hr || (fBindDone && IsEof()))
    {
        SignalDone(hr);
    }

    RRETURN(hr);
}

HRESULT CDwnBindData::ReadFromBind(void* pv, ULONG cb, ULONG* pcb)
{
    Assert(CheckThread());

    HRESULT hr;

#ifdef _DEBUG
    BOOL fBindDone = _fBindDone;
#endif

    *pcb = 0;

    if(_fEof)
    {
        hr = S_FALSE;
    }
    else if(!_fBindOnApt)
    {
        if(_o.pInetProt)
        {
            SetPending(TRUE);

            hr = _o.pInetProt->Read(pv, cb, pcb);
        }
        else
        {
            hr = S_FALSE;
        }
    }
    else if(_u.pstm)
    {
        SetPending(TRUE);

        hr = _u.pstm->Read(pv, cb, pcb);
    }
    else
    {
        hr = S_FALSE;
    }

    AssertSz(hr!=E_PENDING||!fBindDone, 
        "URLMON reports data pending after binding done!");

    if(!hr && cb && *pcb==0)
    {
        hr = S_FALSE;
    }

    if(hr == E_PENDING)
    {
        hr = S_OK;
    }
    else if(hr == S_FALSE)
    {
        SetEof();

        if(_fBindOnApt)
        {
            ClearInterface(&_u.pstm);
        }

        hr = S_OK;
    }
    else if(hr == S_OK)
    {
        SetPending(FALSE);
    }

#ifdef _DEBUG
    _cbBind += *pcb;
#endif

    RRETURN(hr);
}

void CDwnBindData::BufferData()
{
    if(_pDwnStm)
    {
        void*   pv;
        ULONG   cbW, cbR;
        HRESULT hr = S_OK;

        for(;;)
        {
            hr = _pDwnStm->WriteBeg(&pv, &cbW);
            if(hr)
            {
                break;
            }

            Assert(cbW > 0);

            hr = ReadFromBind(pv, cbW, &cbR);
            if(hr)
            {
                break;
            }

            Assert(cbR <= cbW);

            _pDwnStm->WriteEnd(cbR);

            if(cbR == 0)
            {
                break;
            }
        }

        if(hr || _fEof)
        {
            _pDwnStm->WriteEof(hr);
        }

        if(hr)
        {
            SignalDone(hr);
        }
    }
}

BOOL CDwnBindData::IsPending()
{
    return (!_pbPeek && (_pDwnStm?_pDwnStm->IsPending():_fPending));
}

BOOL CDwnBindData::IsEof()
{
    return (!_pbPeek && (_pDwnStm?_pDwnStm->IsEof():_fEof));
}

// CDwnBindData (Callback) ----------------------------------------------------
void CDwnBindData::SetCallback(CDwnLoad* pDwnLoad)
{
    g_csDwnBindSig.Enter();

    _wSig     = _wSigAll;
    _pDwnLoad = pDwnLoad;
    _pDwnLoad->SubAddRef();

    g_csDwnBindSig.Leave();

    if(_wSig)
    {
        Signal(0);
    }
}

void CDwnBindData::Disconnect()
{
    g_csDwnBindSig.Enter();

    CDwnLoad* pDwnLoad = _pDwnLoad;

    _pDwnLoad = NULL;
    _wSig     = 0;

    g_csDwnBindSig.Leave();

    if(pDwnLoad)
    {
        pDwnLoad->SubRelease();
    }
}

void CDwnBindData::Signal(WORD wSig)
{
    SubAddRef();

    for(;;)
    {
        CDwnLoad*   pDwnLoad = NULL;
        WORD        wSigCur  = 0;

        g_csDwnBindSig.Enter();

        // DBF_DONE is sent exactly once, even though it may be signalled
        // more than once.  If we are trying to signal it and have already
        // done so, turn off the flag.
        wSig &= ~(_wSigAll & DBF_DONE);

        // Remember all flags signalled since we started in case we disconnect
        // and reconnect to a new client.  That way we can replay all the
        // missed flags.
        _wSigAll |= wSig;

        if(_pDwnLoad)
        {
            // Someone is listening, so lets figure out what to tell the
            // callback.  Notice that if _dwSigTid is not zero, then a
            // different thread is currently in a callback.  In that case,
            // we just let it do the callback again when it returns.
            _wSig |= wSig;

            if(_wSig && !_dwSigTid)
            {
                wSigCur   = _wSig;
                _wSig     = 0;
                _dwSigTid = GetCurrentThreadId();
                pDwnLoad  = _pDwnLoad;
                pDwnLoad->SubAddRef();
            }
        }

        g_csDwnBindSig.Leave();

        if(!pDwnLoad)
        {
            break;
        }

        if(_pDwnStm && (wSigCur&DBF_DATA))
        {
            BufferData();
        }

        pDwnLoad->OnBindCallback(wSigCur);
        pDwnLoad->SubRelease();

        _dwSigTid = 0;
        wSig      = 0;
    }

    SubRelease();
}

void CDwnBindData::SignalRedirect(LPCTSTR pszText, IUnknown* punkBinding)
{
    IWinInetHttpInfo* pwhi = NULL;
    char  achA[16];
    TCHAR achW[16];
    ULONG cch = ARRAYSIZE(achA);
    HRESULT hr;

    //$ Need protection here.  Redirection can happen more than once.
    //$ This means that GetRedirect() returns a potentially dangerous
    //$ string.

    // Discover the last HTTP request method
    _cstrMethod.Free();
    hr = punkBinding->QueryInterface(IID_IWinInetHttpInfo, (void**)&pwhi);
    if(!hr && pwhi)
    {
        hr = pwhi->QueryInfo(HTTP_QUERY_REQUEST_METHOD, &achA, &cch, NULL, 0);
        if(!hr && cch)
        {
            AnsiToWideTrivial(achA, achW, cch);
            hr = _cstrMethod.Set(achW);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

    // In case the new URL isn't covered by wininet: clear security flags
    _dwSecFlags = IsUrlSecure(pszText) ? SECURITY_FLAG_SECURE : 0;

    // Set the redirect url
    hr = _cstrRedirect.Set(pszText);

    // Redirection means that our POST request, if any, may have become a GET
    if(!_cstrMethod || !_tcsequal(_T("POST"), _cstrMethod))
    {
        AssertSz(FALSE, "must improve");
    }

Cleanup:
    ReleaseInterface(pwhi);
    if(hr)
    {
        SignalDone(hr);
    }
    else
    {
        Signal(DBF_REDIRECT);
    }
}

void CDwnBindData::SignalProgress(DWORD dwPos, DWORD dwMax, DWORD dwStatus)
{
    _DwnProg.dwMax    = dwMax;
    _DwnProg.dwPos    = dwPos;
    _DwnProg.dwStatus = dwStatus;

    Signal(DBF_PROGRESS);
}

void CDwnBindData::SignalFile(LPCTSTR pszText)
{
    if(pszText && *pszText
        && (_dwFlags&DWNF_GETFILELOCK)
        && (_uScheme==URL_SCHEME_HTTP
        || _uScheme==URL_SCHEME_HTTPS
        || _uScheme==URL_SCHEME_FTP
        || _uScheme==URL_SCHEME_GOPHER))
    {
        HRESULT hr = _cstrFile.Set(pszText);

        if(hr)
        {
            SignalDone(hr);
        }
    }
}

void CDwnBindData::SignalHeaders(IUnknown* punkBinding)
{
    IWinInetHttpInfo* pwhi = NULL;
    IWinInetInfo* pwi = NULL;
    BOOL fSignal = FALSE;
    ULONG cch;
    HRESULT hr = S_OK;

    punkBinding->QueryInterface(IID_IWinInetHttpInfo, (void**)&pwhi);

    if(pwhi)
    {
        CHAR    achA[256];
        TCHAR   achW[256];

        if(_dwFlags & DWNF_GETCONTENTTYPE)
        {
            cch = sizeof(achA);

            hr = pwhi->QueryInfo(HTTP_QUERY_CONTENT_TYPE, achA, &cch, NULL, 0);

            if(hr==S_OK && cch>0)
            {
                AnsiToWideTrivial(achA, achW, cch);

                hr = _cstrContentType.Set(achW);
                if(hr)
                {
                    goto Cleanup;
                }

                fSignal = TRUE;
            }

            hr = S_OK;
        }

        if(_dwFlags & DWNF_GETREFRESH)
        {
            cch = sizeof(achA);

            hr = pwhi->QueryInfo(HTTP_QUERY_REFRESH, achA, &cch, NULL, 0);

            if(hr==S_OK && cch>0)
            {
                AnsiToWideTrivial(achA, achW, cch);

                hr = _cstrRefresh.Set(achW);
                if(hr)
                {
                    goto Cleanup;
                }

                fSignal = TRUE;
            }

            hr = S_OK;
        }

        if(_dwFlags & DWNF_GETMODTIME)
        {
            SYSTEMTIME st;
            cch = sizeof(SYSTEMTIME);

            hr = pwhi->QueryInfo(HTTP_QUERY_LAST_MODIFIED|HTTP_QUERY_FLAG_SYSTEMTIME,
                &st, &cch, NULL, 0);

            if(hr==S_OK && cch==sizeof(SYSTEMTIME))
            {
                if(SystemTimeToFileTime(&st, &_ftLastMod))
                {
                    fSignal = TRUE;
                }
            }

            hr = S_OK;
        }

        Assert(!(_dwFlags&DWNF_HANDLEECHO) || (_dwFlags&DWNF_GETSTATUSCODE));

        if(_dwFlags & DWNF_GETSTATUSCODE)
        {
            cch = sizeof(_dwStatusCode);

            hr = pwhi->QueryInfo(HTTP_QUERY_STATUS_CODE|HTTP_QUERY_FLAG_NUMBER,
                &_dwStatusCode, &cch, NULL, 0);

            if(!hr && (_dwFlags&DWNF_HANDLEECHO) && _dwStatusCode==449)
            {
                ULONG cb = 0;

                hr = pwhi->QueryInfo(HTTP_QUERY_FLAG_REQUEST_HEADERS|HTTP_QUERY_ECHO_HEADERS_CRLF, NULL, &cb, NULL, 0);
                if((hr && hr!=S_FALSE) || !cb)
                {
                    goto NoHeaders;
                }

                Assert(!_pbRawEcho);
                MemFree(_pbRawEcho);

                _pbRawEcho = (BYTE*)MemAlloc(cb);
                if(!_pbRawEcho)
                {
                    hr = E_OUTOFMEMORY;
                    goto Cleanup;
                }

                _cbRawEcho = cb;

                hr = pwhi->QueryInfo(HTTP_QUERY_FLAG_REQUEST_HEADERS|HTTP_QUERY_ECHO_HEADERS_CRLF, _pbRawEcho, &cb, NULL, 0);
                Assert(!hr && cb+1==_cbRawEcho);

NoHeaders:
                ;
            }

            fSignal = TRUE;
            hr = S_OK;
        }
    }

    if(!pwhi)
    {
        punkBinding->QueryInterface(IID_IWinInetInfo, (void**)&pwi);
    }
    else
    {
        pwi = pwhi;
        pwhi = NULL;
    }

    if(_dwFlags & DWNF_GETFLAGS)
    {
        if(pwi)
        {
            DWORD dwFlags;
            cch = sizeof(DWORD);

            hr = pwi->QueryOption(INTERNET_OPTION_REQUEST_FLAGS, &dwFlags, &cch);
            if(hr==S_OK && cch==sizeof(DWORD))
            {
                _dwReqFlags = dwFlags;
            }

            // BUGBUG: wininet does not remember security for cached files,
            // if it's from the cache, don't ask wininet; just use security-based-on-url.
            cch = sizeof(DWORD);

            hr = pwi->QueryOption(INTERNET_OPTION_SECURITY_FLAGS, &dwFlags, &cch);
            if(hr==S_OK && cch==sizeof(DWORD))
            {
                if((dwFlags&SECURITY_FLAG_SECURE) || !(_dwReqFlags&INTERNET_REQFLAG_FROM_CACHE))
                {
                    _dwSecFlags = dwFlags;
                }
            }
        }

        fSignal = TRUE; // always pick up security flags

        hr = S_OK;
    }

    if((_dwFlags&DWNF_GETSECCONINFO) && pwi)
    {
        ULONG cb = sizeof(INTERNET_SECURITY_CONNECTION_INFO);
        INTERNET_SECURITY_CONNECTION_INFO isci;

        Assert(!_pSecConInfo);
        MemFree(_pSecConInfo);

        isci.dwSize = cb;

        hr = pwi->QueryOption(INTERNET_OPTION_SECURITY_CONNECTION_INFO, &isci, &cb);
        if(!hr && cb==sizeof(INTERNET_SECURITY_CONNECTION_INFO))
        {
            _pSecConInfo = (INTERNET_SECURITY_CONNECTION_INFO*)MemAlloc(cb);
            if(!_pSecConInfo)
            {
                hr = E_OUTOFMEMORY;
                goto Cleanup;
            }

            memcpy(_pSecConInfo, &isci, cb);
        }

        hr = S_OK;
    }


    if(_cstrFile)
    {
        fSignal = TRUE;

        if(pwi && _uScheme!=URL_SCHEME_FILE)
        {
            cch = sizeof(HANDLE);
            pwi->QueryOption(WININETINFO_OPTION_LOCK_HANDLE, &_hLock, &cch);
        }
    }

Cleanup:
    if(pwhi)
    {
        pwhi->Release();
    }

    if(pwi)
    {
        pwi->Release();
    }

    if(hr)
    {
        SignalDone(hr);
    }
    else if (fSignal)
    {
        Signal(DBF_HEADERS);
    }
}

void CDwnBindData::SignalData()
{
    Signal(DBF_DATA);
}

void CDwnBindData::SignalDone(HRESULT hrErr)
{
    Terminate(hrErr);

    Signal(DBF_DONE);
}

// CDwnBindData (Misc) --------------------------------------------------------
LPCTSTR CDwnBindData::GetFileLock(HANDLE* phLock, BOOL* pfPretransform)
{
    *pfPretransform = _fMimeFilter;
    if(_cstrFile) // note that the file will always be returned
    {
        *phLock = _hLock;
        _hLock = NULL;

        return _cstrFile;
    }

    *phLock = NULL;
    return NULL;
}

void CDwnBindData::GiveRawEcho(BYTE** ppb, ULONG* pcb)
{
    Assert(!*ppb);

    *ppb = _pbRawEcho;
    *pcb = _cbRawEcho;
    _pbRawEcho = NULL;
    _cbRawEcho = 0;
}

void CDwnBindData::GiveSecConInfo(INTERNET_SECURITY_CONNECTION_INFO** ppsci)
{
    Assert(!*ppsci);

    *ppsci = _pSecConInfo;
    _pSecConInfo = NULL;
}

// CDwnBindData (IBindStatusCallback) -----------------------------------------
STDMETHODIMP CDwnBindData::OnStartBinding(DWORD grfBSCOption, IBinding* pbinding)
{
    Assert(_fBindOnApt && CheckThread());

    _fBindDone = FALSE;

    ReplaceInterface(&_u.pbinding, pbinding);

    if(!_fIsDocBind)
    {
        pbinding->SetPriority(THREAD_PRIORITY_BELOW_NORMAL);
    }

    return S_OK;
}

STDMETHODIMP CDwnBindData::OnProgress(ULONG ulPos, ULONG ulMax, ULONG ulCode, LPCWSTR pszText)
{
    Assert(_fBindOnApt && CheckThread());

    switch(ulCode)
    {
    case BINDSTATUS_REDIRECTING:
        SignalRedirect(pszText, _u.pbinding);
        break;

    case BINDSTATUS_CACHEFILENAMEAVAILABLE:
        SignalFile(pszText);
        break;

    case BINDSTATUS_RAWMIMETYPE:
        _pRawMimeInfo = GetMimeInfoFromMimeType(pszText);
        break;

    case BINDSTATUS_MIMETYPEAVAILABLE:
        _pmi = GetMimeInfoFromMimeType(pszText);
        break;

    case BINDSTATUS_LOADINGMIMEHANDLER:
        _fMimeFilter = TRUE;
        break;

    case BINDSTATUS_FINDINGRESOURCE:
    case BINDSTATUS_CONNECTING:
    case BINDSTATUS_BEGINDOWNLOADDATA:
    case BINDSTATUS_DOWNLOADINGDATA:
    case BINDSTATUS_ENDDOWNLOADDATA:
        SignalProgress(ulPos, ulMax, ulCode);
        break;
    }
    return S_OK;
}

STDMETHODIMP CDwnBindData::OnStopBinding(HRESULT hrReason, LPCWSTR szReason)
{
    Assert(_fBindOnApt && CheckThread());

    LPWSTR pchError = NULL;

    SubAddRef();

    _fBindDone = TRUE;

    if(hrReason || _fBindAbort)
    {
        CLSID clsid;
        HRESULT hrUrlmon = S_OK;

        _u.pbinding->GetBindResult(&clsid, (DWORD*)&hrUrlmon, &pchError, NULL);

        // BUGBUG: URLMON returns a native Win32 error.
        if(SUCCEEDED(hrUrlmon))
        {
            hrUrlmon = HRESULT_FROM_WIN32(hrUrlmon);
        }

        SignalDone(hrUrlmon);
    }
    else
    {
        SetPending(FALSE);
        SignalData();
    }

    SubRelease();
    CoTaskMemFree(pchError);

    return S_OK;
}

STDMETHODIMP CDwnBindData::OnDataAvailable(DWORD grfBSCF, DWORD dwSize, FORMATETC* pformatetc, STGMEDIUM* pstgmed)
{
    Assert(_fBindOnApt && CheckThread());

    HRESULT hr = S_OK;

    if(pstgmed->tymed == TYMED_ISTREAM)
    {
        ReplaceInterface(&_u.pstm, pstgmed->pstm);
        ReplaceInterface(&_u.punkForRel, pstgmed->pUnkForRelease);
    }

    if(grfBSCF & (BSCF_DATAFULLYAVAILABLE|BSCF_LASTDATANOTIFICATION))
    {
        _fFullyAvail = TRUE;

        // Clients assume that they can find out how many bytes are fully
        // available in the callback to SignalHeaders.  Fill it in here if
        // we haven't already.
        if(_DwnProg.dwMax == 0)
        {
            _DwnProg.dwMax = dwSize;
        }
    }

    if(!_fSigHeaders)
    {
        _fSigHeaders = TRUE;
        SignalHeaders(_u.pbinding);
    }

    if(!_fSigMime)
    {
        _fSigMime = TRUE;

        if(_pmi == NULL)
        {
            _pmi = GetMimeInfoFromClipFormat(pformatetc->cfFormat);
        }

        Signal(DBF_MIME);
    }

    SetPending(FALSE);
    SignalData();

    RRETURN(hr);
}

// CDwnBindData (IHttpNegotiate) ----------------------------------------------
STDMETHODIMP CDwnBindData::OnResponse(
          DWORD dwResponseCode,
          LPCWSTR szResponseHeaders,
          LPCWSTR szRequestHeaders,
          LPWSTR* ppszAdditionalRequestHeaders)
{
    Assert(CheckThread());

    return S_OK;
}

// CDwnBindData (IInternetBindInfo) -------------------------------------------
STDMETHODIMP CDwnBindData::GetBindInfo(DWORD* pdwBindf, BINDINFO* pbindinfo)
{
    Assert(CheckThread());

    HRESULT hr;

    hr = super::GetBindInfo(pdwBindf, pbindinfo);

    if(hr == S_OK)
    {
        if(!_fBindOnApt)
        {
            *pdwBindf |= BINDF_DIRECT_READ;
        }

        if(_dwFlags & DWNF_IGNORESECURITY)
        {
            *pdwBindf |= BINDF_IGNORESECURITYPROBLEM;
        }
    }

    RRETURN(hr);
}

// CDwnBindData (IInternetProtocolSink) ---------------------------------------
STDMETHODIMP CDwnBindData::Switch(PROTOCOLDATA* ppd)
{
    HRESULT hr;

    if(!_pDwnDoc || _pDwnDoc->IsDocThread())
    {
        hr = _o.pInetProt->Continue(ppd);
    }
    else
    {
        hr = _pDwnDoc->AddDocThreadCallback(this, ppd);
    }

    RRETURN(hr);
}

STDMETHODIMP CDwnBindData::ReportProgress(ULONG ulCode, LPCWSTR pszText)
{
    SubAddRef();

    switch(ulCode)
    {
    case BINDSTATUS_REDIRECTING:
        SignalRedirect(pszText, _o.pInetProt);
        break;

    case BINDSTATUS_CACHEFILENAMEAVAILABLE:
        SignalFile(pszText);
        break;

    case BINDSTATUS_RAWMIMETYPE:
        _pRawMimeInfo = GetMimeInfoFromMimeType(pszText);
        break;

    case BINDSTATUS_MIMETYPEAVAILABLE:
        _pmi = GetMimeInfoFromMimeType(pszText);
        break;

    case BINDSTATUS_LOADINGMIMEHANDLER:
        _fMimeFilter = TRUE;
        break;

    case BINDSTATUS_FINDINGRESOURCE:
    case BINDSTATUS_CONNECTING:
        SignalProgress(0, 0, ulCode);
        break;
    }

    SubRelease();

    return S_OK;
}

STDMETHODIMP CDwnBindData::ReportData(DWORD grfBSCF, ULONG ulPos, ULONG ulMax)
{
    SubAddRef();

    if(grfBSCF & (BSCF_DATAFULLYAVAILABLE|BSCF_LASTDATANOTIFICATION))
    {
        _fFullyAvail = TRUE;

        // Clients assume that they can find out how many bytes are fully
        // available in the callback to SignalHeaders.  Fill it in here if
        // we haven't already.
        if(_DwnProg.dwMax == 0)
        {
            _DwnProg.dwMax = ulMax;
        }
    }

    if(!_fSigHeaders)
    {
        _fSigHeaders = TRUE;
        SignalHeaders(_o.pInetProt);
    }

    if(!_fSigData)
    {
        _fSigData = TRUE;

        Assert(_pDwnStm == NULL);

        // If the data is coming from the network, then read it immediately
        // into a buffers in order to release the socket connection as soon
        // as possible.
        if(!_fFullyAvail
            && (_uScheme==URL_SCHEME_HTTP || _uScheme==URL_SCHEME_HTTPS)
            && !(_dwFlags&(DWNF_DOWNLOADONLY|DWNF_NOAUTOBUFFER)))
        {
            // No problem if this fails.  We just end up not buffering.

            _pDwnStm = new CDwnStm;
        }
    }

    if(!_fSigMime)
    {
        _fSigMime = TRUE;
        Signal(DBF_MIME);
    }

    SignalProgress(ulPos, ulMax, BINDSTATUS_DOWNLOADINGDATA);

    SetPending(FALSE);
    SignalData();

    SubRelease();

    return S_OK;
}

STDMETHODIMP CDwnBindData::ReportResult(HRESULT hrResult, DWORD dwError, LPCWSTR szResult)
{
    SubAddRef();

    _fBindDone = TRUE;

    if(hrResult || _fBindAbort)
    {
        // Mimics urlmon's GetBindResult
        if(dwError)
        {
            hrResult = HRESULT_FROM_WIN32(dwError);
        }

        SignalDone(hrResult);
    }
    else
    {
        SetPending(FALSE);
        SignalData();
    }

    SubRelease();

    return S_OK;
}

// Public Functions -----------------------------------------------------------
HRESULT NewDwnBindData(DWNLOADINFO* pdli, CDwnBindData** ppDwnBindData, DWORD dwFlagsExtra)
{
    HRESULT hr = S_OK;

    if(pdli->pDwnBindData)
    {
        *ppDwnBindData = pdli->pDwnBindData;
        pdli->pDwnBindData->AddRef();
        return(S_OK);
    }

    CDwnBindData* pDwnBindData = new CDwnBindData;

    if(pDwnBindData == NULL)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pDwnBindData->SetDwnDoc(pdli->pDwnDoc);
    pDwnBindData->Bind(pdli, dwFlagsExtra);

    *ppDwnBindData = pDwnBindData;
    pDwnBindData = NULL;

Cleanup:
    if(pDwnBindData)
    {
        pDwnBindData->Release();
    }
    RRETURN(hr);
}

// TlsGetInternetSession ------------------------------------------------------
IInternetSession* TlsGetInternetSession()
{
    IInternetSession* pInetSess = TLS(pInetSess);

    if(pInetSess == NULL)
    {
        CoInternetGetSession(0, &pInetSess, 0);
        TLS(pInetSess) = pInetSess;
    }

    return pInetSess;
}