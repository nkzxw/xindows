
#include "stdafx.h"
#include "DrawInfo.h"

//+---------------------------------------------------------------------------
//
//  Member:     CParentInfo::Init
//
//  Synopsis:   Initialize a CParentInfo.
//
//----------------------------------------------------------------------------
void CParentInfo::Init()
{
    _sizeParent = _afxGlobalData._Zero.size;
}

void CParentInfo::Init(const CDocInfo* pdci)
{
    Assert(pdci);
    ::memcpy(this, pdci, sizeof(CDocInfo));
    Init();
}

void CParentInfo::Init(const CCalcInfo* pci)
{
    ::memcpy(this, pci, sizeof(CParentInfo));
}

void CParentInfo::Init(SIZE* psizeParent)
{
    SizeToParent(psizeParent?psizeParent:&_afxGlobalData._Zero.size);
}

void CParentInfo::Init(CLayout* pLayout)
{
    Assert(pLayout);

    CDocInfo::Init(pLayout->ElementOwner());

    SizeToParent(pLayout);
}

//+---------------------------------------------------------------------------
//
//  Member:     CParentInfo::SizeToParent
//
//  Synopsis:   Set the parent size to the client RECT of the passed CLayout
//
//----------------------------------------------------------------------------
void CParentInfo::SizeToParent(CLayout* pLayout)
{
    RECT rc;

    Assert(pLayout);

    pLayout->GetClientRect(&rc);
    SizeToParent(&rc);
}



//+---------------------------------------------------------------------------
//
//  Member:     CCalcInfo::Init
//
//  Synopsis:   Initialize a CCalcInfo.
//
//----------------------------------------------------------------------------
void CCalcInfo::Init()
{
    _smMode     = SIZEMODE_NATURAL;
    _grfLayout  = 0L;
    _hdc        = (_pDoc ? _pDoc->GetHDC() : TLS(hdcDesktop));
    _yBaseLine  = 0;
    _fUseOffset     = FALSE;
    _fTableCalcInfo = FALSE;
}

void CCalcInfo::Init(const CDocInfo* pdci, CLayout* pLayout)
{
    CView* pView;
    DWORD grfState;

    Assert(pdci);
    Assert(pLayout);

    ::memcpy(this, pdci, sizeof(CDocInfo));

    CParentInfo::Init();

    Init();

    pView = pLayout->GetView();
    Assert(pView);

    grfState = pView->GetState();

    if(!(grfState & (CView::VS_OPEN|CView::VS_INLAYOUT|CView::VS_INRENDER)))
    {
        Verify(pView->OpenView());
    }
}

void CCalcInfo::Init(CLayout* pLayout, SIZE* psizeParent, HDC hdc)
{
    CView* pView;
    DWORD grfState;

    CParentInfo::Init(pLayout);

    _smMode    = SIZEMODE_NATURAL;
    _grfLayout = 0L;

    // If a DC was passed in, use that one - else ask the doc which DC to use.
    // Only if there is no doc yet, use the desktop DC.
    if(hdc)
    {
        _hdc = hdc;
    }
    else if(_pDoc)
    {
        _hdc = _pDoc->GetHDC();
    }
    else
    {
        _hdc = TLS(hdcDesktop);
    }

    _yBaseLine = 0;
    _fUseOffset     = FALSE;
    _fTableCalcInfo = FALSE;

    pView = pLayout->GetView();
    Assert(pView);

    grfState = pView->GetState();

    if(!(grfState & (CView::VS_OPEN|CView::VS_INLAYOUT|CView::VS_INRENDER)))
    {
        Verify(pView->OpenView());
    }
}



//+---------------------------------------------------------------------------
//
//  Member:     CFormDrawInfo::DrawImageFlags
//
//  Synopsis:   Return DRAWIMAGE flags to be used when drawing with the DI
//
//----------------------------------------------------------------------------
DWORD CFormDrawInfo::DrawImageFlags()
{
    return _fHtPalette?0:DRAWIMAGE_NHPALETTE;
}

//+---------------------------------------------------------------------------
//
//  Member:     CFormDrawInfo::Init
//
//  Synopsis:   Initialize paint info for painting to form's hwnd
//
//----------------------------------------------------------------------------
void CFormDrawInfo::Init(HDC hdc, RECT* prcClip)
{
    _hdc = _hic = (hdc ? hdc : TLS(hdcDesktop));

    _fInplacePaint = TRUE;
    _dwDrawAspect  = DVASPECT_CONTENT;
    _lindex        = -1;

    _dvai.cb       = sizeof(_dvai);
    _dvai.dwFlags  = DVASPECTINFOFLAG_CANOPTIMIZE;
    _pvAspect      = (void*)&_dvai;

    _rcClip.top    = _rcClip.left   = LONG_MIN;
    _rcClip.bottom = _rcClip.right  = LONG_MAX;

    IntersectRect(&_rcClipSet, &_pDoc->_pInPlace->_rcClip, &_pDoc->_pInPlace->_rcPos);

    _fDeviceCoords = FALSE;
    _sizeDeviceOffset = _afxGlobalData._Zero.size;
}

//+---------------------------------------------------------------------------
//
//  Member:     CFormDrawInfo::Init
//
//  Synopsis:   Initialize paint info for painting to form's hwnd
//
//----------------------------------------------------------------------------
void CFormDrawInfo::Init(CElement* pElement, HDC hdc, RECT* prcClip)
{
    memset(this, 0, sizeof(*this));
    Assert(pElement);
    CDocInfo::Init(pElement);
    Init(hdc, prcClip);
    InitToSite(pElement->GetUpdatedLayout(), prcClip);
}

//+---------------------------------------------------------------------------
//
//  Member:     CFormDrawInfo::Init
//
//  Synopsis:   Initialize paint info for painting to form's hwnd
//
//----------------------------------------------------------------------------
void CFormDrawInfo::Init(CLayout* pLayout, HDC hdc, RECT* prcClip)
{
    Assert(pLayout);
    Init(pLayout->ElementOwner(), hdc, prcClip);
}

//+---------------------------------------------------------------------------
//
//  Member:     CFormDrawInfo::InitToSite
//
//  Synopsis:   Reduce/set the clipping RECT to the visible portion of
//              the passed CSite
//
//----------------------------------------------------------------------------
void CFormDrawInfo::InitToSite(CLayout* pLayout, RECT* prcClip)
{
    RECT rcClip;
    RECT rcUserClip;

    // For CRootElement (no layout) on normal documents,
    // initialize the clip RECT to that of the entire document
    if(!pLayout)
    {
        IntersectRect(&rcClip, &_pDoc->_pInPlace->_rcClip, &_pDoc->_pInPlace->_rcPos);

        rcUserClip = rcClip;
    }
    // For all sites other than CRootSite or CRootSite on print documents,
    // set _rcClip to the visible client RECT
    // NOTE: Do not pass _rcClip to prevent its modification during initialization
    else
    {
        // BUGBUG: This needs to be fixed! (brendand)
        rcClip = rcUserClip = _afxGlobalData._Zero.rc;
    }

    // Reduce the current clipping RECT
    IntersectRect(&_rcClip, &_rcClip, &rcClip);

    // If the clip RECT is empty, canonicalize it to all zeros
    if(IsRectEmpty(&_rcClip))
    {
        _rcClip = _afxGlobalData._Zero.rc;
    }

    // If passed a clipping RECT, use it to reduce the clip RECT
    if(prcClip)
    {
        IntersectRect(&_rcClip, &_rcClip, prcClip);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     GetDC, GetGlobalDC, GetDirectDrawSurface, GetSurface
//
//  Synopsis:   Each of these is a simple layer over the true CDispSurface call
//
//----------------------------------------------------------------------------
HDC CFormDrawInfo::GetDC(BOOL fPhysicallyClip)
{
    if(!_hdc && _pSurface)
    {
        if(!_fDeviceCoords)
        {
            _pSurface->GetDC(&_hdc, _rc, _rcClip, fPhysicallyClip);
        }

        else
        {
            // we have to do tricky stuff here because the client wants to deal
            // with device coordinates in order to circumvent GDI's 16-bit
            // coordinate limitations on Win 9x.  We can't use SetViewportOrgEx
            // as we can on Win NT because of potential overflow, especially
            // when printing large documents.
            _rc.OffsetRect(-_sizeDeviceOffset);
            _rcClip.OffsetRect(-_sizeDeviceOffset);
            _pSurface->GetDC(&_hdc, _rc, _rcClip, fPhysicallyClip);
            _rc.OffsetRect(_sizeDeviceOffset);
            _rcClip.OffsetRect(_sizeDeviceOffset);
            ::SetViewportOrgEx(_hdc, 0, 0, NULL);
        }
    }

    return _hdc;
}

HDC CFormDrawInfo::GetGlobalDC(BOOL fPhysicallyClip)
{
    // we don't anticipate using the device coordinate mode in this case
    Assert(!_fDeviceCoords);
    HDC hdc = NULL;
    if(_pSurface)
    {
        _pSurface->GetGlobalDC(&hdc, &_rc, &_rcClip, fPhysicallyClip);
    }

    return hdc;
}

HRESULT CFormDrawInfo::GetDirectDrawSurface(IDirectDrawSurface** ppSurface, SIZE* pOffset)
{
    return _pSurface?_pSurface->GetDirectDrawSurface(ppSurface, pOffset):E_FAIL;
}

HRESULT CFormDrawInfo::GetSurface(const IID& iid, void** ppv, SIZE* pOffset)
{
    return _pSurface?_pSurface->GetSurface(iid, ppv, pOffset):E_FAIL;
}

void CFormDrawInfo::SetDeviceCoordinateMode()
{
    _sizeDeviceOffset = _afxGlobalData._Zero.size;
    _fDeviceCoords = FALSE;
}



//+---------------------------------------------------------------------------
//
//  Member:     CSetDrawInfo::CSetDrawInfo
//
//  Synopsis:   Initialize a CSetDrawInfo
//
//----------------------------------------------------------------------------
CSetDrawSurface::CSetDrawSurface(CFormDrawInfo* pDI, const RECT* prcBounds,
                                 const RECT* prcRedraw, CDispSurface* pSurface)
{
    Assert(pDI);
    Assert(pSurface);

    _pDI      = pDI;
    _hdc      = pDI->_hdc;
    _pSurface = pDI->_pSurface;

    _pDI->_pSurface = pSurface;
    _pDI->_rc       = *prcBounds;
    _pDI->_rcClip   = *prcRedraw;
    _pDI->_fIsMemoryDC = _pDI->_pSurface->IsMemory();
    _pDI->_fIsMetafile = _pDI->_pSurface->IsMetafile();

    _pDI->_fDeviceCoords = FALSE;
    _pDI->_sizeDeviceOffset = _afxGlobalData._Zero.size;
}



//+---------------------------------------------------------------------------
//
//  Member:     CDocInfo::Init
//
//  Synopsis:   Initialize a CDocInfo.
//
//----------------------------------------------------------------------------
void CDocInfo::Init(CElement* pElement)
{
    Assert(pElement);
    _pDoc = pElement->Doc();
    Assert(_pDoc);
    Init(&_pDoc->_dci);
}