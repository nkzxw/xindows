
#include "stdafx.h"
#include "Connect.h"

DECLARE_CPtrAry(CCPCEnumCPAry, IConnectionPoint*)

#ifdef _WIN64
DWORD g_dwCookieForWin64 = 0;
#endif

//+------------------------------------------------------------------------
//
//  CEnumConnections Implementation
//
//-------------------------------------------------------------------------

//+---------------------------------------------------------------------------
//
//  Member:     CEnumConnections::Create
//
//  Synopsis:   Creates a new CEnumGeneric object.
//
//  Arguments:  [pary]    -- Array to enumerate.
//              [ppenum]  -- Resulting CEnumGeneric object.
//
//  Returns:    HRESULT.
//
//  History:    5-15-94   adams   Created
//
//----------------------------------------------------------------------------
HRESULT CEnumConnections::Create(int ccd, CONNECTDATA* pcd, CEnumConnections** ppenum)
{
    HRESULT                 hr      = S_OK;
    CEnumConnections*       penum   = NULL;
    CDataAry<CONNECTDATA>*  pary    = NULL;
    CONNECTDATA*            pcdT;
    int                     c;
    int                     cNonNull;

    Assert(ppenum);
    *ppenum = NULL;

    penum = new CEnumConnections;
    if(!penum)
    {
        goto MemError;
    }

    pary = new CDataAry<CONNECTDATA>;
    if(!pary)
    {
        goto MemError;
    }

    Assert(!penum->_pary);

    // Copy the AddRef'd array of pointers into a CFormsAry
    hr = pary->EnsureSize(ccd);
    if(hr)
    {
        goto Error;
    }

    cNonNull = 0;
    for(c=ccd,pcdT=pcd; c>0; c--,pcdT++)
    {
        if(pcdT->pUnk)
        {
            (*pary)[cNonNull] = *pcdT;
            pcdT->pUnk->AddRef();
            cNonNull++;
        }
    }

    pary->SetSize(cNonNull);
    penum->_pary = pary;

    *ppenum = penum;

Cleanup:
    RRETURN(hr);

MemError:
    hr = E_OUTOFMEMORY;
    // fall through

Error:
    ReleaseInterface(penum);
    delete pary;

    goto Cleanup;
}

//+---------------------------------------------------------------------------
//
//  Member:     CEnumConnections::CEnumConnections
//
//  Synopsis:   ctor.
//
//  History:    5-15-94   adams   Created
//
//----------------------------------------------------------------------------
CEnumConnections::CEnumConnections(void) :
CBaseEnum(sizeof(IUnknown*), IID_IEnumConnections, TRUE, TRUE)
{
}

//+---------------------------------------------------------------------------
//
//  Member:     CEnumConnections::CEnumConnections
//
//  Synopsis:   ctor.
//
//  History:    5-15-94   adams   Created
//
//----------------------------------------------------------------------------
CEnumConnections::CEnumConnections(const CEnumConnections& enumc)
: CBaseEnum(enumc)
{
}

//+---------------------------------------------------------------------------
//
//  Member:     CEnumConnections::~CEnumConnections
//
//  Synopsis:   dtor.
//
//  History:    10-12-99   alanau   Created
//
//----------------------------------------------------------------------------
CEnumConnections::~CEnumConnections()
{
    if(_pary && _fDelete)
    {
        if(_fAddRef)
        {
            CONNECTDATA*    pcdT;
            int             i;
            int             cSize = _pary->Size();

            for(i=0,pcdT=(CONNECTDATA*)Deref(0);
                i<cSize;
                i++,pcdT++)
            {
                pcdT->pUnk->Release();
            }
        }

        delete _pary;
        _pary = NULL;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CEnumConnections::Next
//
//  Synopsis:   Per IEnumConnections::Next.
//
//  History:    12-28-93   adams   Created
//
//----------------------------------------------------------------------------
STDMETHODIMP CEnumConnections::Next(ULONG ulConnections, void* reelt, ULONG* lpcFetched)
{
    HRESULT         hr;
    CONNECTDATA*    pCD;        // CONNECTDATA's to enumerate
    CONNECTDATA*    pCDEnd;     // end of CONNECTDATA's
    CONNECTDATA*    pCDSrc;
    int             iActual;    // elems returned

    // Determine number of elements to enumerate.
    if(_i+(int)ulConnections <= _pary->Size())
    {
        hr = S_OK;
        iActual = (int)ulConnections;
    }
    else
    {
        hr = S_FALSE;
        iActual = _pary->Size() - _i;
    }

    if(iActual>0 && !reelt)
    {
        RRETURN(E_INVALIDARG);
    }

    if(lpcFetched)
    {
        *lpcFetched = (ULONG)iActual;
    }

    // Return elements to enumerate.
    pCDEnd = (CONNECTDATA*)reelt + iActual;
    for(pCD=(CONNECTDATA*)reelt,pCDSrc=(CONNECTDATA*)Deref(_i);
        pCD<pCDEnd;
        pCD++,pCDSrc++)
    {
        *pCD = *pCDSrc;
        pCD->pUnk->AddRef();
    }

    _i += iActual;

    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CEnumConnections::Clone
//
//  Synopsis:   Per IEnumConnections::Clone.
//
//  History:    12-28-93   adams   Created
//
//----------------------------------------------------------------------------
STDMETHODIMP CEnumConnections::Clone(CBaseEnum** ppenum)
{
    HRESULT hr;

    if(!ppenum)
    {
        RRETURN(E_INVALIDARG);
    }

    *ppenum = NULL;
    hr = Create(_pary->Size(), (CONNECTDATA*)(LPVOID)*_pary, (CEnumConnections**)ppenum);
    if(hr)
    {
        RRETURN(hr);
    }

    (*(CEnumConnections**)ppenum)->_i = _i;
    return S_OK;
}



//+------------------------------------------------------------------------
//
//  CConnectionPointContainer implementation
//
//-------------------------------------------------------------------------

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPointContainer::CConnectionPointContainer
//
//  Synopsis:   ctor.
//
//  Arguments:  [pBase]     -- CBase object using this container.
//
//----------------------------------------------------------------------------
CConnectionPointContainer::CConnectionPointContainer(CBase* pBase, IUnknown* pUnkOuter)
{
    CONNECTION_POINT_INFO* pcpi;
    long i;

    MemSetName((this, "CPC pBase=%08x", pBase));

    Assert(pBase);
    Assert(pBase->BaseDesc()->_pcpi);

    for(i=0; i<CONNECTION_POINTS_MAX; i++)
    {
        _aCP[i]._index = -1;
    }

    // Note: we set the refcount to 0 because the caller of this function will
    // AddRef it as part of its implementation of QueryInterface.
    _ulRef      = 0;
    _pBase      = pBase;
    _pUnkOuter  = pUnkOuter;

    // _pUnkOuter is specified explicitly when creating a new
    // COM identity on top of the CBase object. In this case,
    // we only need to SubAddRef the CBase object.
    if(_pUnkOuter)
    {
        _pBase->SubAddRef();
        _pUnkOuter->AddRef();
    }
    else
    {
        _pBase->PrivateAddRef();
        _pBase->PunkOuter()->AddRef();
    }

    for(i=0,pcpi=GetCPI(); pcpi&&pcpi->piid; pcpi++,i++)
    {
        _aCP[i]._index = i;
    }

    Assert(i < CONNECTION_POINTS_MAX-1);
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPointContainer::~CConnectionPointContainer
//
//  Synopsis:   dtor.
//
//----------------------------------------------------------------------------
CConnectionPointContainer::~CConnectionPointContainer()
{
    // _pUnkOuter is specified explicitly when creating a new
    // COM identity on top of the CBase object. In this case,
    // we only need to SubRelease the CBase object.
    if(_pUnkOuter)
    {
        _pBase->SubRelease();
        _pUnkOuter->Release();
    }
    else
    {
        _pBase->PrivateRelease();
        _pBase->PunkOuter()->Release();
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CConnectionPointContainer::QueryInterface, IUnknown
//
//  Synopsis:   As per IUnknown.
//
//-------------------------------------------------------------------------
STDMETHODIMP CConnectionPointContainer::QueryInterface(REFIID iid, void** ppvObj)
{
    IUnknown* pUnk = _pUnkOuter ? _pUnkOuter : _pBase->PunkOuter();
    RRETURN_NOTRACE(pUnk->QueryInterface(iid, ppvObj));
}

//+------------------------------------------------------------------------
//
//  Member:     CConnectionPointContainer::AddRef, IUnknown
//
//  Synopsis:   As per IUnknown.
//
//-------------------------------------------------------------------------
STDMETHODIMP_(ULONG) CConnectionPointContainer::AddRef()
{
    return _ulRef++;
}

//+------------------------------------------------------------------------
//
//  Member:     CConnectionPointContainer::Release, IUnknown
//
//  Synopsis:   As per IUnknown.
//
//-------------------------------------------------------------------------
STDMETHODIMP_(ULONG) CConnectionPointContainer::Release()
{
    Assert(_ulRef > 0);

    if(--_ulRef == 0)
    {
        _ulRef = ULREF_IN_DESTRUCTOR;
        delete this;
        return 0;
    }

    return _ulRef;
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPointContainer::EnumConnectionPoints
//
//  Synopsis:   Enumerates the container's connection points.
//
//----------------------------------------------------------------------------
STDMETHODIMP CConnectionPointContainer::EnumConnectionPoints(IEnumConnectionPoints** ppEnum)
{
    HRESULT         hr;
    CCPCEnumCPAry*  paCP;
    CConnectionPt*  pCP;

    if(!ppEnum)
    {
        RRETURN(E_POINTER);
    }

    // Copy pointers to connection points to new array.
    *ppEnum = NULL;
    paCP = new CCPCEnumCPAry;
    if(!paCP)
    {
        RRETURN(E_OUTOFMEMORY);
    }

    for(pCP=_aCP; pCP->_index!=-1; pCP++)
    {
        hr = paCP->Append(pCP);
        if(hr)
        {
            goto Error;
        }

        pCP->AddRef();
    }

    // Create enumerator which references array.
    hr = paCP->EnumElements(IID_IEnumConnectionPoints, (LPVOID*)ppEnum,
        TRUE, FALSE, TRUE);
    if(hr)
    {
        goto Error;
    }

Cleanup:
    RRETURN(hr);

Error:
    paCP->ReleaseAll();
    delete paCP;
    goto Cleanup;
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPointContainer::FindConnectionPoint
//
//  Synopsis:   Finds a connection point with a particular IID.
//
//----------------------------------------------------------------------------
STDMETHODIMP CConnectionPointContainer::FindConnectionPoint(REFIID iid, IConnectionPoint** ppCP)
{
    CConnectionPt* pCP;

    if(!ppCP)
    {
        RRETURN(E_POINTER);
    }

    *ppCP = NULL;

    for(pCP=_aCP; pCP->_index!=-1; pCP++)
    {
        if(*(pCP->MyCPI()->piid) == iid)
        {
            *ppCP = pCP;
            pCP->AddRef();
            return S_OK;
        }
    }

    RRETURN(E_NOINTERFACE);
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPointContainer::GetCPI
//
//  Synopsis:   Get a hold of the array of connection pt information.
//
//----------------------------------------------------------------------------
CONNECTION_POINT_INFO* CConnectionPointContainer::GetCPI()
{
    return (CONNECTION_POINT_INFO*)(_pBase->BaseDesc()->_pcpi);
}



//+------------------------------------------------------------------------
//
//  CConnectionPt implementation
//
//-------------------------------------------------------------------------

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPt::QueryInterface
//
//  Synopsis:   Per IUnknown.
//
//----------------------------------------------------------------------------
STDMETHODIMP CConnectionPt::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
    if(!ppvObj)
    {
        RRETURN(E_INVALIDARG);
    }

    if(iid==IID_IUnknown || iid==IID_IConnectionPoint)
    {
        *ppvObj = (IConnectionPoint*)this;
    }
    else
    {
        *ppvObj = NULL;
        RRETURN_NOTRACE(E_NOINTERFACE);
    }

    ((IUnknown*)*ppvObj)->AddRef();
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPt::AddRef
//
//  Synopsis:   Per IUnknown.
//
//----------------------------------------------------------------------------
STDMETHODIMP_(ULONG) CConnectionPt::AddRef()
{
    MyCPC()->_ulRef++;
    return MyCPC()->_pBase->SubAddRef();
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPt::Release
//
//  Synopsis:   Per IUnknown.
//
//----------------------------------------------------------------------------
STDMETHODIMP_(ULONG) CConnectionPt::Release()
{
    CBase* pBase;

    Assert(MyCPC()->_ulRef > 0);

    pBase = MyCPC()->_pBase;
    if(--MyCPC()->_ulRef == 0)
    {
        MyCPC()->_ulRef = ULREF_IN_DESTRUCTOR;
        delete MyCPC();
    }

    return pBase->SubRelease();
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPt::GetConnectionInterface
//
//  Synopsis:   Returns the IID of this connection point.
//
//----------------------------------------------------------------------------
STDMETHODIMP CConnectionPt::GetConnectionInterface(IID* pIID)
{
    if(!pIID)
    {
        RRETURN(E_POINTER);
    }

    Assert(_index != -1);
    *pIID = *MyCPI()->piid;
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPt::GetConnectionPointContainer
//
//  Synopsis:   Returns the container of this connection point.
//
//----------------------------------------------------------------------------
STDMETHODIMP CConnectionPt::GetConnectionPointContainer(IConnectionPointContainer** ppCPC)
{
    Assert(_index != -1);

    if(!ppCPC)
    {
        RRETURN(E_POINTER);
    }

    if(MyCPC()->_pBase->PunkOuter() != (IUnknown*)MyCPC()->_pBase)
    {
        RRETURN(MyCPC()->_pBase->PunkOuter()->QueryInterface(
            IID_IConnectionPointContainer, (void**)ppCPC));
    }

    *ppCPC = MyCPC();
    (*ppCPC)->AddRef();

    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPt::Advise
//
//  Synopsis:   Adds a connection.
//
//----------------------------------------------------------------------------
STDMETHODIMP CConnectionPt::Advise(LPUNKNOWN pUnkSink, DWORD* pdwCookie)
{
    Assert(_index != -1);

    const CONNECTION_POINT_INFO* pcpi = MyCPI();

    RRETURN(MyCPC()->_pBase->DoAdvise(
        *pcpi->piid,
        pcpi->dispid,
        pUnkSink,
        NULL,
        pdwCookie));
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPt::Unadvise
//
//  Synopsis:   Removes a connection.
//
//----------------------------------------------------------------------------
STDMETHODIMP CConnectionPt::Unadvise(DWORD dwCookie)
{
    Assert(_index != -1);

    RRETURN(MyCPC()->_pBase->DoUnadvise(dwCookie, MyCPI()->dispid));
}

//+---------------------------------------------------------------------------
//
//  Member:     CConnectionPt::EnumConnections
//
//  Synopsis:   Enumerates the connections.
//
//----------------------------------------------------------------------------
STDMETHODIMP CConnectionPt::EnumConnections(LPENUMCONNECTIONS* ppEnum)
{
    Assert(_index != -1);

    HRESULT                 hr;
    DISPID                  dispidBase;
    CEnumConnections*       pEnum   = NULL;
    CDataAry<CONNECTDATA>   aryCD;
    CONNECTDATA             cdNew;
    AAINDEX                 aaidx = AA_IDX_UNKNOWN;

    cdNew.pUnk = NULL;

    if(!ppEnum)
    {
        RRETURN(E_POINTER);
    }

    if(*(MyCPC()->_pBase->GetAttrArray()))
    {
        dispidBase = MyCPI()->dispid;

        // Iterate thru attr array values with this dispid, appending
        // the IUnknown into a local array.
        for(; ;)
        {
            aaidx = MyCPC()->_pBase->FindNextAAIndex(
                dispidBase, CAttrValue::AA_Internal, aaidx);
            if(aaidx == AA_IDX_UNKNOWN)
            {
                break;
            }

            ClearInterface(&cdNew.pUnk);
            hr = MyCPC()->_pBase->GetUnknownObjectAt(aaidx, &cdNew.pUnk);
            if(hr)
            {
                goto Cleanup;
            }

#ifdef _WIN64
            Verify(MyCPC()->_pBase->GetCookieAt(aaidx, &cdNew.dwCookie) == S_OK);
#else
            cdNew.dwCookie = (DWORD)cdNew.pUnk;
#endif

            hr = aryCD.AppendIndirect(&cdNew);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

    hr = CEnumConnections::Create(aryCD.Size(), (CONNECTDATA*)(LPVOID)aryCD, &pEnum);

    // Note that since our enumeration base class doesn't officially
    // support the IEnumConnections interface, we have to do
    // an explicit cast via void *
    *ppEnum = (IEnumConnections*)(void*)pEnum;

Cleanup:
    ReleaseInterface(cdNew.pUnk);
    RRETURN(hr);
}



//+---------------------------------------------------------------------------
//
//  Function:   ConnectSink
//
//  Synopsis:   Connects a sink to an object which fires to the sink through
//              a connection point.
//
//  Arguments:  [pUnkSource] -- The source object.
//              [iid]        -- The id of the connection point.
//              [pUnkSink]   -- The sink.
//              [pdwCookie]  -- Cookie that identifies the connection for
//                                  a later disconnect.  May be NULL.
//
//----------------------------------------------------------------------------
HRESULT ConnectSink(IUnknown* pUnkSource, REFIID iid, IUnknown* pUnkSink, DWORD* pdwCookie)
{
    HRESULT hr;
    IConnectionPointContainer* pCPC;

    Assert(pUnkSource);
    Assert(pUnkSink);

    hr = pUnkSource->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
    if(hr)
    {
        RRETURN(hr);
    }

    hr = ConnectSinkWithCPC(pCPC, iid, pUnkSink, pdwCookie);
    pCPC->Release();
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   ConnectSinkWithCPC
//
//  Synopsis:   Connects a sink to an object which fires to the sink through
//              a connection point.
//
//  Arguments:  [pCPC]       -- The source object as an
//                                  IConnectionPointContainer.
//
//              [iid]        -- The id of the connection point.
//              [pUnkSink]   -- The sink.
//              [pdwCookie]  -- Cookie that identifies the connection for
//                                  a later disconnect.  May be NULL.
//
//----------------------------------------------------------------------------
HRESULT ConnectSinkWithCPC(IConnectionPointContainer* pCPC, REFIID iid, IUnknown* pUnkSink, DWORD* pdwCookie)
{
    HRESULT             hr;
    IConnectionPoint*   pCP;
    DWORD               dwCookie;

    Assert(pCPC);
    Assert(pUnkSink);

    hr = ClampITFResult(pCPC->FindConnectionPoint(iid, &pCP));
    if(hr)
    {
        RRETURN1(hr, E_NOINTERFACE);
    }

    if(!pdwCookie)
    {
        // The CDK erroneously fails to handle a NULL cookie, so
        // we pass in a dummy one.
        pdwCookie = &dwCookie;
    }

    hr = ClampITFResult(pCP->Advise(pUnkSink, pdwCookie));
    pCP->Release();
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   DisconnectSink
//
//  Synopsis:   Disconnects a sink from a source.
//
//  Arguments:  [pUnkSource] -- The source object.
//              [iid]        -- The id of the connection point.
//              [pdwCookie]  -- Pointer to the cookie which identifies the
//                                  connection.
//
//  Modifies:   [*pdwCookie] - sets to 0 if disconnect is successful.
//
//----------------------------------------------------------------------------
HRESULT DisconnectSink(IUnknown* pUnkSource, REFIID iid, DWORD* pdwCookie)
{
    HRESULT                     hr;
    IConnectionPointContainer*  pCPC    = NULL;
    IConnectionPoint*           pCP     = NULL;
    DWORD                       dwCookie;

    Assert(pUnkSource);
    Assert(pdwCookie);

    if(!*pdwCookie)
    {
        return S_OK;
    }

    hr = pUnkSource->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);
    if(hr)
    {
        goto Cleanup;
    }

    hr = ClampITFResult(pCPC->FindConnectionPoint(iid, &pCP));
    if(hr)
    {
        goto Cleanup;
    }

    // Allow clients to use *pdwCookie as an indicator of whether
    // or not they are advised by setting it to zero before
    // calling Unadvise.  This prevents recursion in some
    // scenarios.
    dwCookie = *pdwCookie;
    *pdwCookie = NULL;
    hr = ClampITFResult(pCP->Unadvise(dwCookie));
    if(hr)
    {
        *pdwCookie = dwCookie;
    }

Cleanup:
    ReleaseInterface(pCP);
    ReleaseInterface(pCPC);

    RRETURN(hr);
}
