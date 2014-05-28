
#include "stdafx.h"
#include "Transform.h"

//+---------------------------------------------------------------------------
//
//  Member:     CTransform::WindowFromDocument
//
//  Synopsis:   Convert from document space to window space.
//
//----------------------------------------------------------------------------
long CTransform::WindowFromDocumentX(long xlDoc) const
{
    long dxt = DxtFromHim(xlDoc);
    long lRetValueNew = _ptDst.x + DxzFromDxt(dxt);
    return lRetValueNew;
}

long CTransform::WindowFromDocumentY(long ylDoc) const
{
    long dyt = DytFromHim(ylDoc);
    long lRetValueNew = _ptDst.y + DyzFromDyt(dyt);
    return lRetValueNew;
}

long CTransform::WindowFromDocumentCX(long xlDoc) const
{
    long dxt = DxtFromHim(xlDoc);
    long lRetValueNew = DxzFromDxt(dxt);
    return lRetValueNew;
}

long CTransform::WindowFromDocumentCY(long ylDoc) const
{
    long dyt = DytFromHim(ylDoc);
    long lRetValueNew = DyzFromDyt(dyt);
    return lRetValueNew;
}

void CTransform::WindowFromDocument(RECT* prcOut, const RECTL* prclIn) const
{
    long xtLeft = DxtFromHim(prclIn->left);
    long ytTop = DytFromHim(prclIn->top);
    long xtRight = DxtFromHim(prclIn->right);
    long ytBottom = DytFromHim(prclIn->bottom);

    prcOut->left = _ptDst.x + DxzFromDxt(xtLeft);
    prcOut->top = _ptDst.y + DyzFromDyt(ytTop);
    prcOut->right = _ptDst.x + DxzFromDxt(xtRight);
    prcOut->bottom = _ptDst.y + DyzFromDyt(ytBottom);
}

void CTransform::WindowFromDocument(POINT* pptOut, long xl, long yl) const
{
    long xt = DxtFromHim(xl);
    long yt = DytFromHim(yl);

    pptOut->x = _ptDst.x + DxzFromDxt(xt);
    pptOut->y = _ptDst.y + DyzFromDyt(yt);
}

void  CTransform::WindowFromDocument(SIZE* psizeOut, long cxl, long cyl) const
{
    long dxt = DxtFromHim(cxl);
    long dyt = DytFromHim(cyl);

    psizeOut->cx = DxzFromDxt(dxt);
    psizeOut->cy = DyzFromDyt(dyt);
}

// NB (cthrash) fRelative should be set to TRUE if lValue is relative to
// a pixel value (e.g. CUnitValue::UNIT_PERCENT) and shall be scaled only
// by the zoom factor.  Otherwise we should scale by the current DPI rather
// than the screen DPI.
//
// For example, let's say you have a 100x100 image.  On the screen you want
// it to be 100x100 scaled by any zooming factor.  When printing, you want
// it scaled by zooming factor (which should be unity) plus the relative
// DPI scaling between the screen and the printer.  Generally speaking, this
// would be, in pixel dimensions, larger on the printer than on the screen.
//
// If, on the other hand, we have an image whose width is designated to be
// a percentage of some other value, we don't want to scale by the DPI ratio
// when printing.
long CTransform::DocPixelsFromWindowX(long lValue, BOOL fRelative)
{
    // treat the incoming pixels as if they were in target units
    long lRetValueNew = fRelative ? lValue : DxzFromDxt(lValue, fRelative);
    return lRetValueNew;
}

long CTransform::DocPixelsFromWindowY(long lValue, BOOL fRelative)
{
    // treat the incoming pixels as if they were in target units
    long lRetValueNew = fRelative ? lValue : DyzFromDyt(lValue, fRelative);
    return lRetValueNew;
}

void CTransform::WindowFromDocPixels(POINT* ppt, POINT p, BOOL fRelative)
{
    ppt->x = _ptDst.x + WindowXFromDocPixels(p.x, fRelative);
    ppt->y = _ptDst.y + WindowYFromDocPixels(p.y, fRelative);
}

long CTransform::WindowXFromDocPixels(long lPixels, BOOL fRelative)
{
    long lRetValueNew = DxrFromDxz(lPixels, fRelative);
    return lRetValueNew;
}

long CTransform::WindowYFromDocPixels(long lPixels, BOOL fRelative)
{
    long lRetValueNew = DyrFromDyz(lPixels, fRelative);
    return lRetValueNew;
}

//+---------------------------------------------------------------------------
//
//  Member:     CTransform::DocumentFromWindow
//
//  Synopsis:   Convert from document space to window space.
//
//----------------------------------------------------------------------------
void CTransform::DocumentFromWindow(RECTL* prclOut, const RECT* prcIn) const
{
    long xtLeft = DxtFromDxz(prcIn->left - _ptDst.x);
    long xtRight = DxtFromDxz(prcIn->right - _ptDst.x);
    long ytTop = DytFromDyz(prcIn->top - _ptDst.y);
    long ytBottom = DytFromDyz(prcIn->bottom - _ptDst.y);

    prclOut->left = HimFromDxt(xtLeft);
    prclOut->right = HimFromDxt(xtRight);
    prclOut->top = HimFromDyt(ytTop);
    prclOut->bottom = HimFromDyt(ytBottom);
}

void CTransform::DocumentFromWindow(POINTL* pptlOut, long x, long y) const
{
    long dxt = DxtFromDxz(x - _ptDst.x);
    long dyt = DytFromDyz(y - _ptDst.y);

    pptlOut->x = HimFromDxt(dxt);
    pptlOut->y = HimFromDyt(dyt);
}

void CTransform::DocumentFromWindow(SIZEL* psizelOut, long cx, long cy) const
{
    long dxt = DxtFromDxz(cx);
    long dyt = DytFromDyz(cy);

    psizelOut->cx = HimFromDxt(dxt);
    psizelOut->cy = HimFromDyt(dyt);
}

//+---------------------------------------------------------------------------
//
//  Member:     CTransform::DeviceFromTwips
//              CTransform::TwipsFromDevice
//
//  Synopsis:   Convert twips to device coordinates
//
//----------------------------------------------------------------------------
long CTransform::DeviceFromTwipsCY(long twip) const
{
    long dyt = DytFromTwp(twip);
    long lRetValueNew = DyzFromDyt(dyt);
    return lRetValueNew;
}

long CTransform::DeviceFromTwipsCX(long twip) const
{
    long dxt = DxtFromTwp(twip);
    long lRetValueNew = DxzFromDxt(dxt);
    return lRetValueNew;
}

long CTransform::TwipsFromDeviceCY(long pix) const
{
    long dyt = DytFromDyz(pix);
    long lRetValueNew = TwpFromDyt(dyt);
    return lRetValueNew;
}

long CTransform::TwipsFromDeviceCX(long pix) const
{
    long dxt = DxtFromDxz(pix);
    long lRetValueNew = TwpFromDxt(dxt);
    return lRetValueNew;
}

X CTransform::GetScaledDxpInch()
{
    return MulDivQuick(_wNumerX, _dxrInch, _wDenom);
}

Y CTransform::GetScaledDypInch()
{
    return MulDivQuick(_wNumerY, _dyrInch, _wDenom);
}

//+---------------------------------------------------------------------------
//
//  Pixel scaling routines
//
//----------------------------------------------------------------------------
X CTransform::DxrFromDxt(X dxt) const
{
    // target units ==> unzoomed render units
    return (_fDiff?MulDivQuick(dxt, _dxrInch, _dxtInch):dxt);
}

Y CTransform::DyrFromDyt(Y dyt) const
{
    return (_fDiff?MulDivQuick(dyt, _dyrInch, _dytInch):dyt);
}

X CTransform::DxzFromDxr(X dxr, BOOL fRelative) const
{
    // unzoomed render units ==> zoomed render units
    return ((_fScaled&&!fRelative)?MulDivQuick(dxr, _wNumerX, _wDenom):dxr);
}

Y CTransform::DyzFromDyr(Y dyr, BOOL fRelative) const
{
    return ((_fScaled&&!fRelative)?MulDivQuick(dyr, _wNumerY, _wDenom):dyr);
}

X CTransform::DxzFromDxt(X dxt, BOOL fRelative) const
{
    // wrapper function
    X dxr = DxrFromDxt(dxt);
    return DxzFromDxr(dxr, fRelative);
}

Y CTransform::DyzFromDyt(Y dyt, BOOL fRelative) const
{
    Y dyr = DyrFromDyt(dyt);
    return DyzFromDyr(dyr, fRelative);
}

//+---------------------------------------------------------------------------
//
//  Function:   RectzFromRectr
//
//  Synopsis:   Scale a rectangle according to the current zoom factor.
//	
//  Arguments:  pRectZ      destination scaled rectangle
//              pRectR      source un-scaled rectangle
//
//----------------------------------------------------------------------------
void CTransform::RectzFromRectr(LPRECT pRectZ, LPRECT pRectR)
{
    if(!_fScaled)
    {
        *pRectZ = *pRectR;
        return;
    }

    pRectZ->left = DxzFromDxr(pRectR->left);
    pRectZ->right = pRectZ->left;
    if(pRectR->right > pRectR->left)
    {
        pRectZ->right += max((long)DxzFromDxr(pRectR->right-pRectR->left), 1L);
    }

    pRectZ->top = DyzFromDyr(pRectR->top);
    pRectZ->bottom = pRectZ->top;
    if(pRectR->bottom > pRectR->top)
    {
        pRectZ->bottom += max((long)DyzFromDyr(pRectR->bottom-pRectR->top), 1L);
    }
}

//+---------------------------------------------------------------------------
//
//  Pixel measurement routines
//
//----------------------------------------------------------------------------
X CTransform::DxtFromPts(float pts) const
{
    // points ==> target units
    return MulDivQuick(pts, _dxtInch, ptsInch);
}

Y CTransform::DytFromPts(float pts) const
{
    return MulDivQuick(pts, _dytInch, ptsInch);
}

X CTransform::DxtFromTwp(long twp) const
{
    // twips ==> target units
    return MulDivQuick(twp, _dxtInch, twpInch);
}

Y CTransform::DytFromTwp(long twp) const
{
    return MulDivQuick(twp, _dytInch, twpInch);
}

X CTransform::DxtFromHim(long him) const
{
    // himetric ==> target units
    return MulDivQuick(him, _dxtInch, himInch);
}

Y CTransform::DytFromHim(long him) const
{
    return MulDivQuick(him, _dytInch, himInch);
}

//+---------------------------------------------------------------------------
//
//  Reverse pixel scaling routines
//
//----------------------------------------------------------------------------
X CTransform::DxtFromDxr(X dxr) const
{
    // unzoomed render units ==> target units
    return (_fDiff?MulDivQuick(dxr, _dxtInch, _dxrInch):dxr);
}

Y CTransform::DytFromDyr(Y dyr) const
{
    return (_fDiff?MulDivQuick(dyr, _dytInch, _dyrInch):dyr);
}

X CTransform::DxrFromDxz(X dxz, BOOL fRelative) const
{
    // zoomed render units ==> unzoomed render units
    Assert(_wNumerX);
    return ((_fScaled&&!fRelative)?MulDivQuick(dxz, _wDenom, _wNumerX):dxz);
}

Y CTransform::DyrFromDyz(Y dyz, BOOL fRelative) const
{
    Assert(_wNumerY);
    return ((_fScaled&&!fRelative)?MulDivQuick(dyz, _wDenom, _wNumerY):dyz);
}

X CTransform::DxtFromDxz(X dxz, BOOL fRelative) const
{
    // wrapper function
    X dxr = DxrFromDxz(dxz, fRelative);
    return DxtFromDxr(dxr);
}

Y CTransform::DytFromDyz(Y dyz, BOOL fRelative) const
{
    Y dyr = DyrFromDyz(dyz, fRelative);
    return DytFromDyr(dyr);
}

X CTransform::HimFromDxt(long dxt) const
{
    // target units ==> himetric
    return MulDivQuick(dxt, himInch, _dxtInch);
}

Y CTransform::HimFromDyt(long dyt) const
{
    return MulDivQuick(dyt, himInch, _dytInch);
}

X CTransform::TwpFromDxt(long dxt) const
{
    // target units ==> twips
    return MulDivQuick(dxt, twpInch, _dxtInch);
}

Y CTransform::TwpFromDyt(long dyt) const
{
    return MulDivQuick(dyt, twpInch, _dytInch);
}

//+---------------------------------------------------------------------------
//
//  Member:     CTransform::Init
//
//  Synopsis:   Initialize
//
//----------------------------------------------------------------------------
void CTransform::Init(SIZEL sizelSrc)
{
    _sizeSrc = *(SIZE*)&sizelSrc;

    // initialize to 100% zoom factor by using actual screen resolution
    _sizeInch = _afxGlobalData._sizePixelsPerInch;

    // assume destination size is the same as the source
    _ptDst = _afxGlobalData._Zero.pt;
    _sizeDst.cx = MulDivQuick(_sizeSrc.cx, _sizeInch.cx, 2540L);
    _sizeDst.cy = MulDivQuick(_sizeSrc.cy, _sizeInch.cy, 2540L);

    // We never have zero size. These are used for transforms,
    // and sizeSrc will be used as a denominator in a MulDiv,
    // which will return -1 for a zero divide.
    Assert(_sizeSrc.cx && _sizeDst.cx);
    Assert(_sizeSrc.cy && _sizeDst.cy);

    SetScaleFraction(1, 1, 1);
    SetResolutions(&_afxGlobalData._sizePixelsPerInch, &_afxGlobalData._sizePixelsPerInch);
}

void CTransform::Init(const RECT* prcDst, SIZEL sizelSrc, const SIZE* psizeInch)
{
    _sizeSrc = *(SIZE*)&sizelSrc;
    _ptDst   = *(POINT*)prcDst;
    _sizeDst.cx = prcDst->right - prcDst->left;
    _sizeDst.cy = prcDst->bottom - prcDst->top;

    // if provided a DPI value, go with it, otherwise assume unchanged
    if(psizeInch)
    {
        // ASSUME: this happens only when printing
        _sizeInch = *psizeInch;

        // REVIEW sidda:    hack to force scaling from display pixels to printer pixels when printing
        //                  In future this should be done via explicit conversions.
        SetScaleFraction(1, 1, 1);
        SetResolutions(&_afxGlobalData._sizePixelsPerInch, psizeInch);
    }
    else if(_sizeInch.cx == 0)
    {
        _sizeInch = _afxGlobalData._sizePixelsPerInch; // reset to something reasonable if this became zero for some reason
        SetResolutions(&_afxGlobalData._sizePixelsPerInch, &_afxGlobalData._sizePixelsPerInch);
    }

    // We never have zero size. These are used for transforms,
    // and sizeSrc will be used as a denominator in a MulDiv,
    // which will return -1 for a zero divide.
    Assert(_sizeSrc.cx && _sizeDst.cx);
    Assert(_sizeSrc.cy && _sizeDst.cy);
}

void CTransform::Init(const CTransform* pTransform)
{
    _sizeSrc    = pTransform->_sizeSrc;
    _ptDst      = pTransform->_ptDst;
    _sizeDst    = pTransform->_sizeDst;
    _sizeInch   = pTransform->_sizeInch;

    _wNumerX    = pTransform->_wNumerX;
    _wNumerY    = pTransform->_wNumerY;
    _wDenom     = pTransform->_wDenom;
    _fScaled    = pTransform->_fScaled;

    _dresLayout = pTransform->_dresLayout;
    _dresRender = pTransform->_dresRender;
    _fDiff      = pTransform->_fDiff;
}

//+---------------------------------------------------------------------------
//
//  Function:     GCD
//
//  Synopsis:   Compute GCD
//
//----------------------------------------------------------------------------
int GCD(int w1, int w2)
{
    w1 = abs(w1);
    w2 = abs(w2);

    if(w2 > w1)
    {
        int wT = w1;
        w1 = w2;
        w2 = wT;
    }
    if(w2 == 0)
    {
        return w2;
    }

    for(;;)
    {
        if((w1%=w2) == 0)
        {
            return w2;
        }

        if((w2%=w1) == 0)
        {
            return w1;
        }
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CTransform::zoom
//
//  Synopsis:	Adjust internal data structures as necessary to apply zoom.
//
//  Arguments:  wNumerX     X fraction numerator
//              wNumerY     Y fraction numerator
//              wDenom      fraction denominator
//----------------------------------------------------------------------------
void CTransform::zoom(int wNumerX, int wNumerY, int wDenom)
{
    // update view scaling information
    SetScaleFraction(wNumerX, wNumerY, wDenom);
    _sizeInch.cx = GetScaledDxpInch();
    _sizeInch.cy = GetScaledDypInch();

    _sizeDst.cx = DxzFromDxt(DxtFromHim(_sizeSrc.cx));
    _sizeDst.cy = DyzFromDyt(DytFromHim(_sizeSrc.cy));
}

//+---------------------------------------------------------------------------
//
//  Member:     CTransform::SetScaleFraction
//
//  Synopsis:	Setup a scaling fraction. Only isotropic scaling is supported.
//              This method is meant to do away with the complex SetFrt() method.
//
//  Arguments:  wNumer      fraction numerator
//              wDenom      fraction denominator
//----------------------------------------------------------------------------
void CTransform::SetScaleFraction(int wNumerX, int wNumerY, int wDenom)
{
    Assert(wNumerX >= 0);
    Assert(wNumerY >= 0);
    Assert(wDenom > 0);

    _fScaled = (wNumerX!=wDenom||wNumerY!=wDenom) ? TRUE : FALSE;
    _wNumerX = wNumerX;
    _wNumerY = wNumerY;
    _wDenom = wDenom;
    SimplifyScaleFraction();
}

//+---------------------------------------------------------------------------
//
//  Member:     CTransform::SimplifyScaleFraction
//
//  Synopsis:	Simplify the scaling fraction stored internally.
//----------------------------------------------------------------------------
void CTransform::SimplifyScaleFraction()
{
    int gcd;

    if((gcd=GCD(_wNumerX, _wDenom))>1 && (gcd=GCD(_wNumerY, gcd))>1)
    {
        _wNumerX /= gcd;
        _wNumerY /= gcd;
        _wDenom /= gcd;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CTransform::SetResolutions
//
//  Synopsis:	Initialize the target and rendering device resolutions.
//
//----------------------------------------------------------------------------
void CTransform::SetResolutions(const SIZE* psizeInchTarget, const SIZE* psizeInchRender)
{
    _dresLayout.dxuInch = psizeInchTarget->cx;
    _dresLayout.dyuInch = psizeInchTarget->cy;
    _dresRender.dxuInch = psizeInchRender->cx;
    _dresRender.dyuInch = psizeInchRender->cy;

    _fDiff = (_dxtInch!=_dxrInch||_dytInch!=_dyrInch) ? TRUE : FALSE;
}