
#include "stdafx.h"
#include "Dwn.h"

#include "Image.h"
#include "Bits.h"

BOOL WINAPI UnlockUrlCacheEntryFileBugW(LPCWSTR lpszUrl, DWORD dwReserved)
{
    CStrIn strInUrl(lpszUrl);

    return UnlockUrlCacheEntryFileA(strInUrl, dwReserved);
}

UINT GetUrlScheme(const TCHAR * pchUrlIn)
{
    PARSEDURL puw = { 0 };

    if(!pchUrlIn)
    {
        return (UINT)URL_SCHEME_INVALID;
    }

    puw.cbSize = sizeof(PARSEDURL);

    return (SUCCEEDED(ParseURL(pchUrlIn, &puw)))?puw.nScheme:URL_SCHEME_INVALID;
}

//+---------------------------------------------------------------------------
//
//  Function:   IsSecureUrl
//
//  Synopsis:   Returns TRUE for https, FALSE for anything else
//
//              Starting with an IE 4.01 QFE in 1/98, IsUrlSecure returns TRUE 
//              for javascript:, vbscript:, and about: if the source 
//              document in the wrapped URL is secure. 
// 
//----------------------------------------------------------------------------
BOOL IsUrlSecure(const TCHAR* pchUrl)
{
    BOOL fSecure;
    ULONG cb;

    if(!pchUrl)
    {
        return FALSE;
    }

    switch(GetUrlScheme(pchUrl))
    {
    case URL_SCHEME_HTTPS:
        return TRUE;

    case URL_SCHEME_HTTP:
    case URL_SCHEME_FILE:
    case URL_SCHEME_RES:
    case URL_SCHEME_FTP:
        return FALSE;

    default:
        if(!CoInternetQueryInfo(pchUrl, QUERY_IS_SECURE, 0, &fSecure, sizeof(fSecure), &cb, 0) && cb==sizeof(fSecure))
        {
            return fSecure;
        }
        return FALSE;
    }
}

// Types ---------------------------------------------------------------
struct DWNINFOENTRY
{
    DWORD dwKey;
    CDwnInfo* pDwnInfo;
};

struct FINDINFO
{
    LPCTSTR         pchUrl;
    UINT            dt;
    DWORD           dwKey;
    UINT            cbUrl;
    UINT            iEnt;
    UINT            cEnt;
    CDwnInfo*       pDwnInfo;
    DWNINFOENTRY*   pde;
};

typedef CDataAry<DWNINFOENTRY> CDwnInfoAry;

// Globals --------------------------------------------------------------------
CDwnInfoAry g_aryDwnInfo;           // Active references
CDwnInfoAry g_aryDwnInfoCache;      // Cached references
ULONG       g_ulDwnInfoSize         = 0;
ULONG       g_ulDwnInfoItemThresh   = 128 * 1024;
ULONG       g_ulDwnInfoThreshBytes  = 1024 * 1024;
LONG        g_ulDwnInfoThreshCount  = 128;
DWORD       g_dwDwnInfoLru          = 0;

// Internal -------------------------------------------------------------------
BOOL DwnCacheFind(CDwnInfoAry* pary, FINDINFO* pfi)
{
    DWNINFOENTRY*   pde;
    LPCTSTR         pchUrl;
    CDwnInfo*       pDwnInfo;
    DWORD           dwKey;
    UINT            cb1, cb2;

    if(pfi->pde == NULL)
    {
        DWNINFOENTRY*   pdeBase = *pary;
        LONG            iEntLo = 0, iEntHi = (LONG)pary->Size() - 1, iEnt;

        cb1 = pfi->cbUrl = _tcslen(pfi->pchUrl) * sizeof(TCHAR);

        HashData((BYTE*)pfi->pchUrl, pfi->cbUrl, (BYTE*)&dwKey, sizeof(DWORD));

        pfi->dwKey = dwKey;

        while(iEntLo <= iEntHi)
        {
            iEnt = (iEntLo + iEntHi) / 2;
            pde  = &pdeBase[iEnt];

            if(pde->dwKey == dwKey)
            {
                iEntLo = iEnt;

                while(iEnt>0 && pde[-1].dwKey==dwKey)
                {
                    --iEnt, --pde;
                }

                pfi->iEnt = iEnt;
                pfi->cEnt = pary->Size() - iEnt - 1;
                pfi->pde  = pde;

                goto validate;
            }
            else if(pde->dwKey < dwKey)
            {
                iEntLo = iEnt + 1;
            }
            else
            {
                iEntHi = iEnt - 1;
            }
        }

        pfi->iEnt = iEntLo;

        return FALSE;
    }

advance:
    if(pfi->cEnt == 0)
    {
        return FALSE;
    }

    pfi->iEnt += 1;
    pfi->cEnt -= 1;
    pfi->pde  += 1;
    pde        = pfi->pde;
    cb1        = pfi->cbUrl;
    dwKey      = pfi->dwKey;

    if(pde->dwKey != dwKey)
    {
        return FALSE;
    }

validate:
    pDwnInfo = pde->pDwnInfo;

    if(pDwnInfo->GetType() != pfi->dt)
    {
        goto advance;
    }

    pchUrl = pDwnInfo->GetUrl();
    cb2    = _tcslen(pchUrl) * sizeof(TCHAR);

    if(cb1!=cb2 || memcmp(pchUrl, pfi->pchUrl, cb1)!=0)
    {
        goto advance;
    }

    pfi->pDwnInfo = pDwnInfo;

    return TRUE;
}

void DwnCachePurge()
{
    LONG            cb, cbItem;
    UINT            iEnt, cEnt, iLru = 0;
    DWORD           dwLru, dwLruItem;
    DWNINFOENTRY*   pde;
    CDwnInfo*       pDwnInfo;

    cb   = (LONG)(g_ulDwnInfoSize - g_ulDwnInfoThreshBytes);
    cEnt = g_aryDwnInfoCache.Size();

    if(cb<=0 && cEnt>(UINT)g_ulDwnInfoThreshCount)
    {
        Assert(cEnt-g_ulDwnInfoThreshCount == 1);
        cb = 1;
    }

    while(cb>0 && cEnt>0)
    {
        dwLru = 0xFFFFFFFF;
        pde   = g_aryDwnInfoCache;

        for(iEnt=0; iEnt<cEnt; ++iEnt,++pde)
        {
            pDwnInfo  = pde->pDwnInfo;
            dwLruItem = pDwnInfo->GetLru();

            if(dwLru > dwLruItem)
            {
                dwLru = dwLruItem;
                iLru  = iEnt;
            }
        }

        if(dwLru != 0xFFFFFFFF)
        {
            pDwnInfo = g_aryDwnInfoCache[iLru].pDwnInfo;
            cbItem   = pDwnInfo->GetCacheSize();

            cb -= cbItem;
            g_ulDwnInfoSize -= cbItem;

            g_aryDwnInfoCache.Delete(iLru);
            cEnt -= 1;

            pDwnInfo->SubRelease();
        }
    }

    Assert(g_ulDwnInfoSize <= g_ulDwnInfoThreshBytes);
    Assert(g_aryDwnInfoCache.Size() <= g_ulDwnInfoThreshCount);
}

void DwnCacheDeinit()
{
    DWNINFOENTRY* pde = g_aryDwnInfoCache;
    UINT cEnt = g_aryDwnInfoCache.Size();

    for(; cEnt>0; --cEnt,++pde)
    {
#ifdef _DEBUG
        g_ulDwnInfoSize -= pde->pDwnInfo->GetCacheSize();
#endif
        pde->pDwnInfo->SubRelease();
    }

    Assert(g_ulDwnInfoSize == 0);
    g_aryDwnInfoCache.DeleteAll();

    AssertSz(g_aryDwnInfo.Size()==0,
        "One or more CDwnInfo objects were leaked.  Most likely caused by "
        "leaking a CDoc or CImgElement.");

    g_aryDwnInfo.DeleteAll();
}

int FtCompare(FILETIME* pft1, FILETIME* pft2)
{
    return (pft1->dwHighDateTime>pft2->dwHighDateTime ? 1 :
        (pft1->dwHighDateTime<pft2->dwHighDateTime ? -1 :
        (pft1->dwLowDateTime>pft2->dwLowDateTime ? 1 :
        (pft1->dwLowDateTime<pft2->dwLowDateTime ? -1 : 0))));
}

// AnsiToWideTrivial ----------------------------------------------------------
void AnsiToWideTrivial(const CHAR* pchA, WCHAR* pchW, LONG cch)
{
    for(; cch>=0; --cch)
    {
        *pchW++ = *pchA++;
    }
}

// CDwnStmStm -----------------------------------------------------------------
class CDwnStmStm : public CBaseFT, public IStream
{
    typedef CBaseFT super;

public:
    DECLARE_MEMCLEAR_NEW_DELETE()

    CDwnStmStm(CDwnStm* pDwnStm);
    virtual void Passivate();

    // IUnknown methods
    STDMETHOD (QueryInterface)(REFIID riid, void** ppv);
    STDMETHOD_(ULONG,AddRef)();
    STDMETHOD_(ULONG,Release)();

    // IStream
    STDMETHOD(Clone)(IStream** ppStream);
    STDMETHOD(Commit)(DWORD dwFlags);
    STDMETHOD(CopyTo)(IStream* pStream, ULARGE_INTEGER cb, ULARGE_INTEGER* pcbRead, ULARGE_INTEGER* pcbWrite);
    STDMETHOD(LockRegion)(ULARGE_INTEGER ib, ULARGE_INTEGER cb, DWORD dwLockType);
    STDMETHOD(Read)(void HUGEP* pv, ULONG cb, ULONG* pcb);
    STDMETHOD(Revert)();
    STDMETHOD(Seek)(LARGE_INTEGER ib, DWORD dwOrigin, ULARGE_INTEGER* pib);
    STDMETHOD(SetSize)(ULARGE_INTEGER cb);
    STDMETHOD(Stat)(STATSTG* pstatstg, DWORD dwFlags);
    STDMETHOD(UnlockRegion)(ULARGE_INTEGER ib, ULARGE_INTEGER cb, DWORD dwLockType);
    STDMETHOD(Write)(const void HUGEP* pv, ULONG cb, ULONG* pcb);

protected:
    CDwnStm* _pDwnStm;
    ULONG _ib;
};

CDwnStmStm::CDwnStmStm(CDwnStm* pDwnStm)
{
    _pDwnStm = pDwnStm;
    _pDwnStm->AddRef();
}

void CDwnStmStm::Passivate(void)
{
    _pDwnStm->Release();

    super::Passivate();
}

STDMETHODIMP CDwnStmStm::QueryInterface(REFIID iid, void** ppv)
{
    HRESULT hr;

    if(iid==IID_IUnknown || iid==IID_IStream)
    {
        *ppv = (IStream*)this;
        AddRef();
        hr = S_OK;
    }
    else
    {
        *ppv = NULL;
        hr = E_NOINTERFACE;
    }

    return hr;
}

STDMETHODIMP_(ULONG) CDwnStmStm::AddRef()
{
    ULONG ulRefs = super::AddRef();
    return ulRefs;
}

STDMETHODIMP_(ULONG) CDwnStmStm::Release()
{
    ULONG ulRefs = super::Release();
    return ulRefs;
}

STDMETHODIMP CDwnStmStm::Clone(IStream ** ppStream)
{
    HRESULT hr = E_NOTIMPL;
    *ppStream = NULL;
    RRETURN(hr);
}

STDMETHODIMP CDwnStmStm::Commit(DWORD dwFlags)
{
    HRESULT hr = S_OK;
    RRETURN(hr);
}

STDMETHODIMP CDwnStmStm::CopyTo(
        IStream* pStream,
        ULARGE_INTEGER cb,
        ULARGE_INTEGER* pcbRead,
        ULARGE_INTEGER* pcbWrite)
{
    HRESULT hr = E_NOTIMPL;
    RRETURN(hr);
}

STDMETHODIMP CDwnStmStm::LockRegion(ULARGE_INTEGER ib, ULARGE_INTEGER cb, DWORD dwLockType)
{
    HRESULT hr = E_NOTIMPL;
    RRETURN(hr);
}

STDMETHODIMP CDwnStmStm::Read(void HUGEP* pv, ULONG cb, ULONG* pcb)
{
    ULONG   cbRead;
    HRESULT hr;

    if(pcb == NULL)
    {
        pcb = &cbRead;
    }

    *pcb = 0;

    hr = _pDwnStm->Seek(_ib);
    if(hr)
    {
        goto Cleanup;
    }

    hr = _pDwnStm->Read(pv, cb, pcb);
    if(hr)
    {
        goto Cleanup;
    }

    _ib += *pcb;

    if(*pcb == 0)
    {
        hr = S_FALSE;
        goto Cleanup;
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

STDMETHODIMP CDwnStmStm::Revert()
{
    HRESULT hr = E_NOTIMPL;
    RRETURN(hr);
}

STDMETHODIMP CDwnStmStm::Seek(LARGE_INTEGER ib, DWORD dwOrigin, ULARGE_INTEGER* pib)
{
    HRESULT hr = E_NOTIMPL;

    if(dwOrigin == STREAM_SEEK_SET)
    {
        _ib = ib.LowPart;
        hr = S_OK;
    }

    RRETURN(hr);
}

STDMETHODIMP CDwnStmStm::SetSize(ULARGE_INTEGER cb)
{
    HRESULT hr = E_NOTIMPL;
    RRETURN(hr);
}

STDMETHODIMP CDwnStmStm::Stat(STATSTG* pstatstg, DWORD dwFlags)
{
    memset(pstatstg, 0, sizeof(STATSTG));

    pstatstg->type           = STGTY_STREAM;
    pstatstg->grfMode        = STGM_READ;
    pstatstg->cbSize.LowPart = _pDwnStm->Size();

    return S_OK;
}

STDMETHODIMP CDwnStmStm::UnlockRegion(ULARGE_INTEGER ib, ULARGE_INTEGER cb, DWORD dwLockType)
{
    HRESULT hr = E_NOTIMPL;
    RRETURN(hr);
}

STDMETHODIMP CDwnStmStm::Write(const void HUGEP* pv, ULONG cb, ULONG* pcb)
{
    HRESULT hr = E_NOTIMPL;
    RRETURN(hr);
}

HRESULT CreateStreamOnDwnStm(CDwnStm* pDwnStm, IStream** ppStream)
{
    *ppStream = new CDwnStmStm(pDwnStm);
    RRETURN(*ppStream ? S_OK : E_OUTOFMEMORY);
}

#ifdef _DEBUG
CDwnCrit* g_pDwnCritHead;

CDwnCrit::CDwnCrit(LPCSTR pszName, UINT cLevel)
{
    ::InitializeCriticalSection(GetPcs());

    _pszName        = pszName;
    _cLevel         = cLevel;
    _dwThread       = 0;
    _cEnter         = 0;
    _pDwnCritNext   = g_pDwnCritHead;
    g_pDwnCritHead  = this;
}

void CDwnCrit::Enter()
{
    ::EnterCriticalSection(GetPcs());

    Assert(_dwThread==0 || _dwThread==GetCurrentThreadId());

    if(_dwThread == 0)
    {
        _dwThread = GetCurrentThreadId();

        CDwnCrit* pDwnCrit = g_pDwnCritHead;

        for(; pDwnCrit; pDwnCrit=pDwnCrit->_pDwnCritNext)
        {
            if(pDwnCrit == this)
            {
                continue;
            }

            if(pDwnCrit->_dwThread != GetCurrentThreadId())
            {
                continue;
            }

            if(pDwnCrit->_cLevel > _cLevel)
            {
                continue;
            }

            char ach[256];

            wsprintfA(ach, "CDwnCrit: %s (%d) -> %s (%d) deadlock potential",
                pDwnCrit->_pszName, pDwnCrit->_cLevel, _pszName?_pszName:"", _cLevel);

            AssertSz(0, ach);
        }
    }

    _cEnter += 1;
}

void CDwnCrit::Leave()
{
    Assert(_dwThread == GetCurrentThreadId());
    Assert(_cEnter > 0);

    if(--_cEnter == 0)
    {
        _dwThread = 0;
    }

    ::LeaveCriticalSection(GetPcs());
}
#endif

CDwnCrit::CDwnCrit()
{
    ::InitializeCriticalSection(GetPcs());

#ifdef _DEBUG
    _pszName        = NULL;
    _cLevel         = (UINT)-1;
    _dwThread       = 0;
    _cEnter         = 0;
    _pDwnCritNext   = NULL;
#endif
}

CDwnCrit::~CDwnCrit()
{
    Assert(_dwThread == 0);
    Assert(_cEnter == 0);
    ::DeleteCriticalSection(GetPcs());
}



CDwnStm::CDwnStm(UINT cbBuf) : CDwnChan(g_csDwnStm.GetPcs())
{
    _cbBuf = cbBuf;
}

CDwnStm::~CDwnStm()
{
    BUF* pbuf;
    BUF* pbufNext;

    for(pbuf=_pbufHead; pbuf; pbuf=pbufNext)
    {
        pbufNext = pbuf->pbufNext;
        MemFree(pbuf);
    }
}

HRESULT CDwnStm::SetSeekable()
{
    _fSeekable = TRUE;
    return S_OK;
}

HRESULT CDwnStm::Write(void* pv, ULONG cb)
{
    void* pvW;
    ULONG cbW;
    HRESULT hr = S_OK;

    while(cb > 0)
    {
        hr = WriteBeg(&pvW, &cbW);
        if(hr)
        {
            goto Cleanup;
        }

        if(cbW > cb)
        {
            cbW = cb;
        }

        memcpy(pvW, pv, cbW);

        WriteEnd(cbW);

        pv = (BYTE*)pv + cbW;
        cb = cb - cbW;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CDwnStm::WriteBeg(void** ppv, ULONG* pcb)
{
    BUF* pbuf = _pbufWrite;
    HRESULT hr = S_OK;

    if(pbuf == NULL)
    {
        pbuf = (BUF*)MemAlloc(offsetof(BUF, ab) + _cbBuf);

        if(pbuf == NULL)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        pbuf->pbufNext = NULL;
        pbuf->ib = 0;
        pbuf->cb = _cbBuf;

        g_csDwnStm.Enter();

        if(_pbufTail == NULL)
        {
            _pbufHead = pbuf;
            _pbufTail = pbuf;
        }
        else
        {
            _pbufTail->pbufNext = pbuf;
            _pbufTail = pbuf;
        }

        if(_pbufRead == NULL)
        {
            Assert(_ibRead == 0);
            _pbufRead = pbuf;
        }

        g_csDwnStm.Leave();

        _pbufWrite = pbuf;
    }

    Assert(pbuf->cb > pbuf->ib);

    *ppv = &pbuf->ab[pbuf->ib];
    *pcb = pbuf->cb - pbuf->ib;

Cleanup:
    RRETURN(hr);
}

void CDwnStm::WriteEnd(ULONG cb)
{
    if(cb > 0)
    {
        BUF* pbuf = _pbufWrite;
        ULONG ib   = pbuf->ib + cb;

        Assert(ib <= pbuf->cb);

        if(ib >= pbuf->cb)
        {
            _pbufWrite = NULL;
        }

        // As soon as pbuf->ib is written and matches pbuf->cb, the reader
        // can asynchronously read and free the buffer.  Therefore, pbuf
        // cannot be accessed after this next line.
        pbuf->ib  = ib;
        _cbWrite += cb;

        Signal();
    }
}

void CDwnStm::WriteEof(HRESULT hrEof)
{
    if(!_fEof || hrEof)
    {
        _hrEof = hrEof;
        _fEof  = TRUE;
        Signal();
    }
}

HRESULT CDwnStm::Read(void* pv, ULONG cb, ULONG* pcbRead)
{
    void*   pvR;
    ULONG   cbR;
    ULONG   cbRead  = 0;
    HRESULT hr      = S_OK;

    while(cb > 0)
    {
        hr = ReadBeg(&pvR, &cbR);
        if(hr)
        {
            break;
        }

        if(cbR == 0)
        {
            break;
        }

        if(cbR > cb)
        {
            cbR = cb;
        }

        memcpy(pv, pvR, cbR);

        pv = (BYTE*)pv + cbR;
        cb = cb - cbR;
        cbRead += cbR;

        ReadEnd(cbR);
    }

    *pcbRead = cbRead;

    RRETURN(hr);
}

HRESULT CDwnStm::ReadBeg(void** ppv, ULONG* pcb)
{
    BUF* pbuf = _pbufRead;
    HRESULT hr = S_OK;

    if(_fEof && _hrEof!=S_OK)
    {
        *ppv = NULL;
        *pcb = 0;
        hr   = _hrEof;
        goto Cleanup;
    }

    if(pbuf)
    {
        Assert(_ibRead <= pbuf->ib);

        *ppv = &pbuf->ab[_ibRead];
        *pcb = pbuf->ib - _ibRead;
    }
    else
    {
        *ppv = NULL;
        *pcb = 0;
    }

Cleanup:
    return S_OK;
}

void CDwnStm::ReadEnd(ULONG cb)
{
    BUF* pbuf = _pbufRead;

    Assert(pbuf);
    Assert(_ibRead+cb <= pbuf->ib);

    _ibRead += cb;

    if(_ibRead >= pbuf->cb)
    {
        _ibRead = 0;

        g_csDwnStm.Enter();

        _pbufRead = pbuf->pbufNext;

        if(!_fSeekable)
        {
            _pbufHead = _pbufRead;

            if(_pbufHead == NULL)
            {
                _pbufTail = NULL;
            }
        }

        g_csDwnStm.Leave();

        if(!_fSeekable)
        {
            MemFree(pbuf);
        }
    }

    _cbRead += cb;
}

BOOL CDwnStm::ReadEof(HRESULT* phrEof)
{
    if(_fEof && (_hrEof || _cbRead==_cbWrite))
    {
        *phrEof = _hrEof;
        return TRUE;
    }
    else
    {
        *phrEof = S_OK;
        return FALSE;
    }
}

HRESULT CDwnStm::Seek(ULONG ib)
{
    BUF*    pbuf;
    ULONG   cb;
    HRESULT hr = S_OK;

    if(!_fSeekable || ib > _cbWrite)
    {
        Assert(FALSE);
        hr = E_FAIL;
        goto Cleanup;
    }

    pbuf = _pbufHead;
    cb   = ib;

    if(pbuf)
    {
        while(cb > pbuf->cb)
        {
            cb  -= pbuf->cb;
            pbuf = pbuf->pbufNext;
        }
    }

    g_csDwnStm.Enter();

    if(!pbuf || cb<pbuf->cb)
    {
        _pbufRead = pbuf;
        _ibRead   = cb;
    }
    else
    {
        _pbufRead = pbuf->pbufNext;
        _ibRead   = 0;
    }

    g_csDwnStm.Leave();

    _cbRead = ib;

Cleanup:
    RRETURN(hr);
}

HRESULT CDwnStm::CopyStream(IStream* pstm, ULONG* pcbCopy)
{
    void*   pv;
    ULONG   cbW;
    ULONG   cbR;
    ULONG   cbRTotal;
    HRESULT hr = S_OK;

    cbRTotal = 0;

    for(;;)
    {
        hr = WriteBeg(&pv, &cbW);
        if(hr)
        {
            goto Cleanup;
        }

        Assert(cbW > 0);

        hr = pstm->Read(pv, cbW, &cbR);
        if(FAILED(hr))
        {
            goto Cleanup;
        }

        hr = S_OK;

        Assert(cbR <= cbW);

        WriteEnd(cbR);

        cbRTotal += cbR;

        if(cbR == 0)
        {
            break;
        }
    }

Cleanup:
    if(pcbCopy)
    {
        *pcbCopy = cbRTotal;
    }

    RRETURN(hr);
}


CDwnInfo::CDwnInfo() : CBaseFT(&_cs) {}

CDwnInfo::~CDwnInfo()
{
    if(_pDwnInfoLock)
    {
        _pDwnInfoLock->SubRelease();
    }
}

void CDwnInfo::Passivate()
{
    EnterCriticalSection();

#ifdef _DEBUG
    _fPassive = TRUE;
#endif

    if(_arySink.Size() > 0)
    {
        SINKENTRY* pSinkEntry = _arySink;
        UINT cSink = _arySink.Size();

        for(; cSink>0; --cSink,++pSinkEntry)
        {
            if(pSinkEntry->dwCookie)
            {
                pSinkEntry->pProgSink->DelProgress(pSinkEntry->dwCookie);
            }
            pSinkEntry->pProgSink->Release();
        }

        _arySink.SetSize(0);
    }

    LeaveCriticalSection();
}

HRESULT CDwnInfo::Init(DWNLOADINFO* pdli)
{
    CDwnDoc* pDwnDoc = pdli->pDwnDoc;
    HRESULT hr;

    hr = _cstrUrl.Set(pdli->pchUrl);
    if(hr)
    {
        goto Cleanup;
    }

    _dwBindf   = pDwnDoc->GetBindf();
    _dwRefresh = pdli->fResynchronize ? IncrementLcl() : pDwnDoc->GetRefresh();
    _dwFlags   = DWNLOAD_NOTLOADED | (pDwnDoc->GetDownf() & ~DWNF_STATE);

Cleanup:
    RRETURN(hr);
}

void CDwnInfo::DelProgSinks()
{
    EnterCriticalSection();

    for(CDwnCtx* pDwnCtx=_pDwnCtxHead; pDwnCtx; pDwnCtx=pDwnCtx->_pDwnCtxNext)
    {
        Assert(pDwnCtx->_pDwnInfo == this);
        pDwnCtx->SetProgSink(NULL);
    }

    LeaveCriticalSection();
}

void CDwnInfo::AddDwnCtx(CDwnCtx* pDwnCtx)
{
    AddRef();

    EnterCriticalSection();

    pDwnCtx->_pDwnInfo    = this;
    pDwnCtx->_pDwnCtxNext = _pDwnCtxHead;
    _pDwnCtxHead          = pDwnCtx;
    pDwnCtx->SetPcs(GetPcs());

    LeaveCriticalSection();
}

void CDwnInfo::DelDwnCtx(CDwnCtx* pDwnCtx)
{
    EnterCriticalSection();

    CDwnCtx** ppDwnCtx = &_pDwnCtxHead;

    for(; *ppDwnCtx; ppDwnCtx=&(*ppDwnCtx)->_pDwnCtxNext)
    {
        if(*ppDwnCtx == pDwnCtx)
        {
            *ppDwnCtx = pDwnCtx->_pDwnCtxNext;

            DelProgSink(pDwnCtx->_pProgSink);

            Assert(pDwnCtx->_pDwnInfo == this);
            pDwnCtx->_pDwnInfo = NULL;
            goto found;
        }
    }

    AssertSz(FALSE, "Couldn't find CDwnCtx in CDwnInfo list");

found:
    LeaveCriticalSection();
    Release();
}

HRESULT CDwnInfo::GetFile(LPTSTR* ppch)
{
    HRESULT hr;
    LPCTSTR pchUrl = GetUrl();

    *ppch = NULL;

    if(_tcsnipre(_T("file"), 4, pchUrl, -1))
    {
        TCHAR achPath[MAX_PATH];
        DWORD cchPath;

        hr = CoInternetParseUrl(pchUrl, PARSE_PATH_FROM_URL, 0,
            achPath, ARRAYSIZE(achPath), &cchPath, 0);
        if(hr)
        {
            goto Cleanup;
        }

        hr = MemAllocString(achPath, ppch);
    }
    else
    {
        BYTE                        buf[MAX_CACHE_ENTRY_INFO_SIZE];
        INTERNET_CACHE_ENTRY_INFO*  pInfo = (INTERNET_CACHE_ENTRY_INFO*)buf;
        DWORD                       cInfo = sizeof(buf);

        if(RetrieveUrlCacheEntryFile(pchUrl, pInfo, &cInfo, 0))
        {
            DoUnlockUrlCacheEntryFile(pchUrl, 0);
            hr = MemAllocString(pInfo->lpszLocalFileName, ppch);
        }
        else
        {
            hr = E_FAIL;
        }
    }

Cleanup:
    RRETURN(hr);
}

void CDwnInfo::Signal(WORD wChg)
{
    if(_pDwnCtxHead)
    {
        EnterCriticalSection();

        for(CDwnCtx* pDwnCtx=_pDwnCtxHead; pDwnCtx; pDwnCtx=pDwnCtx->_pDwnCtxNext)
        {
            pDwnCtx->Signal(wChg);
        }

        LeaveCriticalSection();
    }
}

void CDwnInfo::SetLoad(CDwnCtx* pDwnCtx, BOOL fLoad, BOOL fReload, DWNLOADINFO* pdli)
{
    CDwnLoad* pDwnLoadOld = NULL;
    CDwnLoad* pDwnLoadNew = NULL;
    HRESULT hr = S_OK;

    Assert(!EnteredCriticalSection());

    EnterCriticalSection();

    int cLoad = fReload ? (pDwnCtx->_fLoad?0:1) : (fLoad?1:-1);

    if(cLoad==0 && !(TstFlags(DWNLOAD_ERROR|DWNLOAD_STOPPED)))
    {
        goto Cleanup;
    }

    Assert(!(_cLoad==0 && cLoad==-1));

    _cLoad += cLoad;
    pDwnCtx->_fLoad = fLoad;

#ifdef _DEBUG
    {
        UINT cLoad = 0;
        CDwnCtx* pDwnCtxT = _pDwnCtxHead;

        for(; pDwnCtxT; pDwnCtxT=pDwnCtxT->_pDwnCtxNext)
        {
            cLoad += !!pDwnCtxT->_fLoad;
        }

        AssertSz(cLoad==_cLoad, "CDwnInfo _cLoad is inconistent with "
            "sum of CDwnCtx _fLoad");
    }
#endif

    if((cLoad>0 && _cLoad==1 && !TstFlags(DWNLOAD_COMPLETE))
        || (cLoad==0 && _cLoad>0))
    {
        if(!TstFlags(DWNLOAD_NOTLOADED))
        {
            Abort(E_ABORT, &pDwnLoadOld);
            Reset();
        }
        else
        {
            StartProgress();
        }

        Assert(_pDwnLoad == NULL);

        UpdFlags(DWNLOAD_MASK, DWNLOAD_LOADING);

        hr = NewDwnLoad(&_pDwnLoad);

        if(hr == S_OK)
        {
            hr = _pDwnLoad->Init(pdli, this);
        }

        if(hr == S_OK)
        {
            pDwnLoadNew = _pDwnLoad;
            pDwnLoadNew->AddRef();
        }
        else
        {
            Abort(hr, &pDwnLoadOld);
        }
    }
    else if(cLoad<0 && _cLoad==0)
    {
        Abort(S_OK, &pDwnLoadOld);
    }

Cleanup:
    LeaveCriticalSection();

    Assert(!EnteredCriticalSection());

    if(pDwnLoadOld)
    {
        pDwnLoadOld->Release();
    }

    if(pDwnLoadNew)
    {
        pDwnLoadNew->SetCallback();
        pDwnLoadNew->Release();
    }
}

void CDwnInfo::OnLoadDone(CDwnLoad* pDwnLoad, HRESULT hrErr)
{
    Assert(!EnteredCriticalSection());

    EnterCriticalSection();

    if(pDwnLoad == _pDwnLoad)
    {
        OnLoadDone(hrErr);
        _pDwnLoad = NULL;
    }
    else
    {
        pDwnLoad = NULL;
    }

    LeaveCriticalSection();

    if(pDwnLoad)
    {
        pDwnLoad->Release();
    }
}

void CDwnInfo::Abort(HRESULT hrErr, CDwnLoad** ppDwnLoad)
{
    Assert(EnteredCriticalSection());

    if(TstFlags(DWNLOAD_LOADING))
    {
        UpdFlags(DWNLOAD_MASK, hrErr?DWNLOAD_ERROR:DWNLOAD_STOPPED);
        Signal(IMGCHG_COMPLETE);
    }

    *ppDwnLoad = _pDwnLoad;
    _pDwnLoad  = NULL;
}

HRESULT CDwnInfo::AddProgSink(IProgSink* pProgSink)
{
    Assert(EnteredCriticalSection());
    Assert(!_fPassive);

    SINKENTRY*  pSinkEntry  = _arySink;
    UINT        cSink       = _arySink.Size();
    DWORD       dwCookie;
    HRESULT     hr = S_OK;

    for(; cSink>0; --cSink,++pSinkEntry)
    {
        if(pSinkEntry->pProgSink == pProgSink)
        {
            pSinkEntry->ulRefs += 1;
            goto Cleanup;
        }
    }

    hr = _arySink.AppendIndirect(NULL, &pSinkEntry);
    if(hr)
    {
        goto Cleanup;
    }

    // Don't add the progress if we're still at the (nonpending) NOTLOADED state
    if(!TstFlags(DWNLOAD_NOTLOADED))
    {
        hr = pProgSink->AddProgress(GetProgSinkClass(), &dwCookie);

        if(hr==S_OK && _pDwnLoad)
        {
            hr = _pDwnLoad->RequestProgress(pProgSink, dwCookie);

            if(hr)
            {
                pProgSink->DelProgress(dwCookie);
                dwCookie = 0;
            }
        }
    }
    else
    {
        dwCookie = 0; // not a valid cookie
    }

    if(hr)
    {
        _arySink.Delete(_arySink.Size()-1);
        goto Cleanup;
    }

    pSinkEntry->pProgSink = pProgSink;
    pSinkEntry->ulRefs    = 1;
    pSinkEntry->dwCookie  = dwCookie;
    pProgSink->AddRef();

Cleanup:
    RRETURN(hr);
}

void CDwnInfo::StartProgress()
{
    SINKENTRY*  pSinkEntry  = _arySink;
    UINT        cSink       = _arySink.Size();
    HRESULT     hr;

    Assert(TstFlags(DWNLOAD_NOTLOADED));

    for(; cSink; --cSink,++pSinkEntry)
    {
        Assert(pSinkEntry->ulRefs && !pSinkEntry->dwCookie);

        hr = pSinkEntry->pProgSink->AddProgress(GetProgSinkClass(), &(pSinkEntry->dwCookie));

        if(hr==S_OK && _pDwnLoad)
        {
            hr = _pDwnLoad->RequestProgress(pSinkEntry->pProgSink, pSinkEntry->dwCookie);

            if(hr)
            {
                pSinkEntry->pProgSink->DelProgress(pSinkEntry->dwCookie);
                pSinkEntry->dwCookie = 0;
            }
        }
    }
}

void CDwnInfo::DelProgSink(IProgSink* pProgSink)
{
    Assert(EnteredCriticalSection());

    SINKENTRY* pSinkEntry = _arySink;
    UINT cSink = _arySink.Size();

    for(; cSink>0; --cSink,++pSinkEntry)
    {
        if(pSinkEntry->pProgSink == pProgSink)
        {
            if(--pSinkEntry->ulRefs == 0)
            {
                if(pSinkEntry->dwCookie)
                {
                    pProgSink->DelProgress(pSinkEntry->dwCookie);
                }
                pProgSink->Release();
                _arySink.Delete(_arySink.Size()-cSink);
            }
            break;
        }
    }
}

HRESULT CDwnInfo::SetProgress(DWORD dwFlags, DWORD dwState,
        LPCTSTR pch, DWORD dwIds, DWORD dwPos, DWORD dwMax)
{
    if(dwFlags && _arySink.Size())
    {
        EnterCriticalSection();

        SINKENTRY* pSinkEntry = _arySink;
        UINT cSink = _arySink.Size();

        for(; cSink > 0; --cSink,++pSinkEntry)
        {
            pSinkEntry->pProgSink->SetProgress(pSinkEntry->dwCookie,
                dwFlags, dwState, pch, dwIds, dwPos, dwMax);
        }

        LeaveCriticalSection();
    }

    return S_OK;
}



CDwnLoad::~CDwnLoad()
{
    if(_pDwnInfo)
    {
        _pDwnInfo->SubRelease();
    }

    if(_pDwnBindData)
    {
        _pDwnBindData->Release();
    }

    if(_pDownloadNotify)
    {
        _pDownloadNotify->Release();
    }
}

HRESULT CDwnLoad::Init(DWNLOADINFO* pdli, CDwnInfo* pDwnInfo, UINT idsLoad, DWORD dwFlagsExtra)
{
    HRESULT hr;
    TCHAR* pchAlloc = NULL;
    const TCHAR* pchUrl;

    Assert(!_pDwnInfo && !_pDwnBindData);
    Assert(!_pDownloadNotify);

    _cDone    = 1;
    _idsLoad  = idsLoad;
    _pDwnInfo = pDwnInfo;
    _pDwnInfo->SubAddRef();

    SetPcs(_pDwnInfo->GetPcs());

    // Only notify IDownloadNotify if load is _NOT_ from
    // an IStream, an existing bind context, or client-supplied data
    _pDownloadNotify = pdli->pDwnDoc && !pdli->pstm && !pdli->pbc && !pdli->fClientData?pdli->pDwnDoc->GetDownloadNotify():NULL;

    if(_pDownloadNotify)
    {
        _pDownloadNotify->AddRef();

        pchUrl = pdli->pchUrl;

        // If all we have is a moniker, we can try to extract the URL anyway
        if(!pchUrl && pdli->pmk)
        {
            hr = pdli->pmk->GetDisplayName(NULL, NULL, &pchAlloc);
            if(hr)
            {
                goto Cleanup;
            }

            pchUrl = pchAlloc;
        }

        //$ WIN64: IDownloadNotify::DownloadStart needs a DWORD_PTR as second argument
        hr = _pDownloadNotify->DownloadStart(pchUrl, (DWORD)(DWORD_PTR)this, pDwnInfo->GetType(), 0);
        if(hr)
        {
            goto Cleanup;
        }
    }

    hr = NewDwnBindData(pdli, &_pDwnBindData, dwFlagsExtra);

    // initial guess for security based on url (may be updated in onbindheaders)
    _pDwnInfo->SetSecFlags(_pDwnBindData->GetSecFlags());

Cleanup:
    CoTaskMemFree(pchAlloc);
    RRETURN(hr);
}

void CDwnLoad::SetCallback()
{
    _pDwnBindData->SetCallback(this);
}

void CDwnLoad::OnBindCallback(DWORD dwFlags)
{
    HRESULT hr = S_OK;

    Assert(!EnteredCriticalSection());

    if(dwFlags & DBF_PROGRESS)
    {
        DWNPROG DwnProg;

        _pDwnBindData->GetProgress(&DwnProg);

        hr = OnBindProgress(&DwnProg);
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(dwFlags & DBF_REDIRECT)
    {
        hr = OnBindRedirect(_pDwnBindData->GetRedirect(), _pDwnBindData->GetMethod());
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(dwFlags & DBF_HEADERS)
    {
        hr = OnBindHeaders();
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(dwFlags & DBF_MIME)
    {
        hr = OnBindMime(_pDwnBindData->GetMimeInfo());
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(dwFlags & DBF_DATA)
    {
        hr = OnBindData();
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    if(FAILED(hr))
    {
        if(!_fDwnBindTerm)
        {
            _pDwnBindData->Disconnect();
            _pDwnBindData->Terminate(hr);
        }

        dwFlags |= DBF_DONE;
    }

    if(dwFlags & DBF_DONE)
    {
        OnBindDone(_pDwnBindData->GetBindResult());
    }
}

void CDwnLoad::Passivate()
{
    _fPassive = TRUE;

    if(_pDwnBindData && !_fDwnBindTerm)
    {
        _pDwnBindData->Disconnect();
        _pDwnBindData->Terminate(E_ABORT);
    }

    super::Passivate();
}

HRESULT CDwnLoad::RequestProgress(IProgSink* pProgSink, DWORD dwCookie)
{
    HRESULT hr;

    hr = pProgSink->SetProgress(dwCookie,
        PROGSINK_SET_STATE|PROGSINK_SET_TEXT|PROGSINK_SET_IDS|
        PROGSINK_SET_POS|PROGSINK_SET_MAX,
        _dwState, GetProgText(), _dwIds, _dwPos, _dwMax);

    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CDwnLoad::OnBindProgress(DWNPROG* pDwnProg)
{
    DWORD   dwFlags = 0;
    DWORD   dwState = _dwState;
    DWORD   dwPos   = _dwPos;
    DWORD   dwMax   = _dwMax;
    UINT    dwIds   = _dwIds;

    if(_fPassive)
    {
        goto Cleanup;
    }

    switch(pDwnProg->dwStatus)
    {
    case BINDSTATUS_FINDINGRESOURCE:
    case BINDSTATUS_CONNECTING:
        dwState  = PROGSINK_STATE_CONNECTING;
        dwIds    = _idsLoad;
        break;

    case BINDSTATUS_BEGINDOWNLOADDATA:
    case BINDSTATUS_DOWNLOADINGDATA:
    case BINDSTATUS_ENDDOWNLOADDATA:
        dwState  = PROGSINK_STATE_LOADING;
        dwIds    = _idsLoad;
        dwPos    = pDwnProg->dwPos;
        dwMax    = pDwnProg->dwMax;

        if(_dwState!=PROGSINK_STATE_LOADING && _pDwnBindData->GetRedirect())
        {
            // Looks like we got redirected somewhere else.  Force the
            // progress text to get recalculated.
            _dwIds = 0;
        }

        break;

    default:
        goto Cleanup;
    }

    if(_dwState != dwState)
    {
        _dwState = dwState;
        dwFlags |= PROGSINK_SET_STATE;
    }

    if(_dwPos != dwPos)
    {
        _dwPos = dwPos;
        dwFlags |= PROGSINK_SET_POS;
    }

    if(_dwMax != dwMax)
    {
        _dwMax = dwMax;
        dwFlags |= PROGSINK_SET_MAX;
    }

    if(_dwIds != dwIds)
    {
        _dwIds = dwIds;
        dwFlags |= PROGSINK_SET_TEXT | PROGSINK_SET_IDS;
    }

    if(dwFlags)
    {
        const TCHAR* pch = (dwFlags&PROGSINK_SET_TEXT) ? GetProgText() : NULL;

        _pDwnInfo->SetProgress(dwFlags, _dwState, pch, _dwIds, _dwPos, _dwMax);
    }

Cleanup:
    return S_OK;
}

LPCTSTR CDwnLoad::GetProgText()
{
    if(_dwState == PROGSINK_STATE_LOADING)
    {
        LPCTSTR pch = _pDwnBindData->GetRedirect();
        if(pch)
        {
            return pch;
        }
    }

    return GetUrl();
}

void CDwnLoad::OnDone(HRESULT hrErr)
{
    if(InterlockedDecrement(&_cDone) == 0)
    {
        if(_pDownloadNotify)
        {
            //$ WIN64: IDownloadNotify::DownloadComplete needs a DWORD_PTR as first argument
            _pDownloadNotify->DownloadComplete((DWORD)(DWORD_PTR)this, hrErr, 0);
        }

        _hrErr = hrErr;
        _pDwnInfo->OnLoadDone(this, hrErr);
    }
}




HRESULT CDwnInfo::Create(UINT dt, DWNLOADINFO* pdli, CDwnInfo** ppDwnInfo)
{
    CDwnInfo*   pDwnInfo    = NULL;
    FINDINFO    fi;
    BOOL        fScanActive = FALSE;
    HRESULT     hr          = S_OK;

    g_csDwnCache.Enter();
    memset(&fi, 0, sizeof(FINDINFO));

    if(dt!=DWNCTX_HTM && pdli->pchUrl && *pdli->pchUrl)
    {
        CDwnDoc* pDwnDoc = pdli->pDwnDoc;

        fi.pchUrl = pdli->pchUrl;
        fi.dt     = dt;

        if(fi.dt == DWNCTX_FILE)
        {
            fi.dt = DWNCTX_BITS;
        }

        while(DwnCacheFind(&g_aryDwnInfo, &fi))
        {
            fScanActive = TRUE;

            // If we have a bind-in-progress, don't attach early to an image in the cache.  We end
            // up abandoning the bind-in-progress and it never terminates properly.  We probably
            // could just hit it with an ABORT, but this case is not common enough to add yet
            // another state transition at this time (this case is when you hyperlink to a page
            // which turns out to actually be an image).
            if(pdli->fResynchronize || pdli->pDwnBindData)
            {
                break;
            }

            if(fi.pDwnInfo->AttachEarly(dt, pDwnDoc->GetRefresh(), pDwnDoc->GetDownf(), pDwnDoc->GetBindf()))
            {
                pDwnInfo = fi.pDwnInfo;
                pDwnInfo->AddRef();

                break;
            }
        }
    }

    if(pDwnInfo == NULL)
    {
        switch(dt)
        {
        case DWNCTX_IMG:
            pDwnInfo = new CImgInfo();
            break;

        case DWNCTX_BITS:
        case DWNCTX_FILE:
            pDwnInfo = new CBitsInfo(dt);
            break;

        default:
            AssertSz(FALSE, "Unknown DWNCTX type");
            break;
        }

        if(pDwnInfo == NULL)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        hr = pDwnInfo->Init(pdli);

        if(hr==S_OK && dt!=DWNCTX_HTM && *pDwnInfo->GetUrl())
        {
            DWNINFOENTRY de = { fi.dwKey, pDwnInfo };
            hr = g_aryDwnInfo.InsertIndirect(fi.iEnt, &de);

            if(hr == S_OK)
            {
                if(dt==DWNCTX_IMG
                    && (pDwnInfo->GetBindf()&(BINDF_GETNEWESTVERSION|BINDF_NOWRITECACHE|BINDF_RESYNCHRONIZE|BINDF_PRAGMA_NO_CACHE))==0
                    && !pdli->pDwnBindData)
                {
                    UINT uScheme = GetUrlScheme(pDwnInfo->GetUrl());

                    if(uScheme==URL_SCHEME_FILE || uScheme==URL_SCHEME_HTTP || uScheme==URL_SCHEME_HTTPS)
                    {
                        pDwnInfo->AttachByLastModEx(fScanActive, uScheme);
                    }
                }
            }
        }

        if(hr)
        {
            pDwnInfo->Release();
            goto Cleanup;
        }
    }

    *ppDwnInfo = pDwnInfo;

Cleanup:
    g_csDwnCache.Leave();
    RRETURN(hr);
}

BOOL CDwnInfo::AttachByLastMod(CDwnLoad* pDwnLoad, FILETIME* pft, BOOL fDoAttach)
{
    if(!_cstrUrl || !*_cstrUrl)
    {
        return(FALSE);
    }

    EnterCriticalSection();

    BOOL fRet = FALSE;

    if(pDwnLoad == _pDwnLoad)
    {
        _ftLastMod = *pft;

        if(fDoAttach)
        {
            g_csDwnCache.Enter();

            fRet = AttachByLastModEx(TRUE, URL_SCHEME_UNKNOWN);

            g_csDwnCache.Leave();
        }
    }

    LeaveCriticalSection();

    return fRet;
}

BOOL CDwnInfo::AttachByLastModEx(BOOL fScanActive, UINT uScheme)
{
    UINT        dt          = GetType();
    FINDINFO    fi          = { _cstrUrl, dt };
    CDwnInfo*   pDwnInfo    = NULL;
    BOOL        fGotLastMod = _ftLastMod.dwLowDateTime || _ftLastMod.dwHighDateTime;

    if(fScanActive)
    {
        while(DwnCacheFind(&g_aryDwnInfo, &fi))
        {
            if(fi.pDwnInfo!=this && fi.pDwnInfo->TstFlags(DWNLOAD_COMPLETE))
            {
                if(!fGotLastMod)
                {
                    if(!GetUrlLastModTime(_cstrUrl, uScheme, _dwBindf, &_ftLastMod))
                    {
                        goto Cleanup;
                    }

                    fGotLastMod = TRUE;
                }

                if(memcmp(&_ftLastMod, &fi.pDwnInfo->_ftLastMod, sizeof(FILETIME)) == 0)
                {
                    // If the active object is locking another object, then attach
                    // to the locked object in order to avoid chains of locked objects.
                    pDwnInfo = fi.pDwnInfo->_pDwnInfoLock;

                    if(pDwnInfo == NULL)
                    {
                        pDwnInfo = fi.pDwnInfo;
                    }

                    if(CanAttachLate(pDwnInfo))
                    {
                        pDwnInfo->SubAddRef();
                    }
                    else
                    {
                        pDwnInfo = NULL;
                    }

                    break;
                }
            }
        }
    }

    if(pDwnInfo == NULL)
    {
        memset(&fi, 0, sizeof(FINDINFO));
        fi.pchUrl = _cstrUrl;
        fi.dt     = dt;

        while(DwnCacheFind(&g_aryDwnInfoCache, &fi))
        {
            if(fi.pDwnInfo!=this && CanAttachLate(fi.pDwnInfo))
            {
                if(!fGotLastMod)
                {
                    if(!GetUrlLastModTime(_cstrUrl, uScheme, _dwBindf, &_ftLastMod))
                    {
                        goto Cleanup;
                    }

                    fGotLastMod = TRUE;
                }

                if(memcmp(&_ftLastMod, &fi.pDwnInfo->_ftLastMod, sizeof(FILETIME)) == 0)
                {
                    Assert(!fi.pDwnInfo->_pDwnInfoLock);

                    pDwnInfo = fi.pDwnInfo;

                    g_ulDwnInfoSize -= pDwnInfo->GetCacheSize();
                    g_aryDwnInfoCache.Delete(fi.iEnt);

                    break;
                }
            }
        }
    }

    if(pDwnInfo)
    {
        AttachLate(pDwnInfo);
        pDwnInfo->SubRelease();
    }

Cleanup:
    return !!pDwnInfo;
}

ULONG CDwnInfo::Release()
{
    g_csDwnCache.Enter();

    ULONG ulRefs = (ULONG)InterlockedRelease();

    if(ulRefs==0 && _cstrUrl && *_cstrUrl)
    {
        FINDINFO    fi       = { _cstrUrl, GetType() };
        CDwnInfo*   pDwnInfo = NULL;
        DWORD       cbSize;
        UINT        iEntDel  = 0;
        BOOL        fDeleted = FALSE;

        // We want to cache this object if it is completely loaded and
        // we either own the bits (don't have a lock on another object)
        // or we are locking another object, but it has no active references.
        // In the latter case, we actually want to cache the locked object,
        // not this object, because we don't allow entries in the cache which
        // have locks on other objects.
        if(!TstFlags(DWNLOAD_COMPLETE))
        {
        }
        else if(TstFlags(DWNF_DOWNLOADONLY))
        {
        }
        else if(_ftLastMod.dwLowDateTime==0 && _ftLastMod.dwHighDateTime==0)
        {
        }
        else if(_pDwnInfoLock)
        {
            if(_pDwnInfoLock->GetRefs() > 0)
            {
            }
            else
            {
                pDwnInfo = _pDwnInfoLock;
                pDwnInfo->SubAddRef();
            }
        }
        else
        {
            pDwnInfo = this;
            SubAddRef();
        }

        while(DwnCacheFind(&g_aryDwnInfo, &fi))
        {
            if(fi.pDwnInfo == this)
            {
                iEntDel = fi.iEnt;
                fDeleted = TRUE;

                if(pDwnInfo == NULL)
                {
                    break;
                }
            }
            else if(pDwnInfo && fi.pDwnInfo->_pDwnInfoLock==pDwnInfo)
            {
                // Some other active object has the target locked.  We don't
                // put it into the cache because it is still accessible
                // through that active object.
                pDwnInfo->SubRelease();
                pDwnInfo = NULL;

                if(fDeleted)
                {
                    break;
                }
            }
        }

        if(fDeleted)
        {
            g_aryDwnInfo.Delete(iEntDel);
        }

        AssertSz(fDeleted, "Can't find reference to CDwnInfo in active array");

        if(pDwnInfo)
        {
            cbSize = pDwnInfo->ComputeCacheSize();

            if(cbSize == 0)
            {
            }
            else if(cbSize > g_ulDwnInfoItemThresh)
            {
            }
            else
            {
                pDwnInfo->SetCacheSize(cbSize);
                pDwnInfo->SetLru(++g_dwDwnInfoLru);

                memset(&fi, 0, sizeof(FINDINFO));
                fi.pchUrl = pDwnInfo->GetUrl();
                fi.dt     = pDwnInfo->GetType();

                while(DwnCacheFind(&g_aryDwnInfoCache, &fi))
                {
                    if(FtCompare(&fi.pDwnInfo->_ftLastMod, &pDwnInfo->_ftLastMod) > 0)
                    {
                        // The current entry in the cache has a newer mod date
                        // then the one we're trying to add.  Forget about
                        // ours.
                        goto endcache;
                    }

                    // Replace the older entry with this newer one.
                    g_ulDwnInfoSize -= fi.pDwnInfo->GetCacheSize();


                    fi.pDwnInfo->SubRelease();
                    fi.pDwnInfo = pDwnInfo;
                    g_aryDwnInfoCache[fi.iEnt].pDwnInfo = pDwnInfo;
                    pDwnInfo = NULL;

                    g_ulDwnInfoSize += cbSize;

                    if(g_ulDwnInfoSize > g_ulDwnInfoThreshBytes)
                    {
                        DwnCachePurge();
                    }

                    goto endcache;
                }

                // No matching entry found, but now we know where to insert.
                DWNINFOENTRY de = { fi.dwKey, pDwnInfo };

                if(g_aryDwnInfoCache.InsertIndirect(fi.iEnt, &de) == S_OK)
                {
                    // Each CDwnInfo stored in the cache maintains a secondary
                    // reference count (inheriting from CBaseFT).  Since this
                    // object is now globally cached without active references,
                    // it should no longer maintain its secondary reference
                    // because it will be destroyed when the DLL is unloaded.
                    pDwnInfo = NULL;

                    g_ulDwnInfoSize += cbSize;

                    if(g_ulDwnInfoSize>g_ulDwnInfoThreshBytes
                        || g_aryDwnInfoCache.Size()>g_ulDwnInfoThreshCount)
                    {
                        DwnCachePurge();
                    }
                }
            }

endcache:
            if(pDwnInfo)
            {
                pDwnInfo->SubRelease();
                pDwnInfo = NULL;
            }
        }
    }

    g_csDwnCache.Leave();

    if(ulRefs == 0)
    {
        Passivate();
        SubRelease();
    }

    return ulRefs;
}