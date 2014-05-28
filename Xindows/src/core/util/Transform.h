
#ifndef __XINDOWS_CORE_UTIL_TRANSFORM_H__
#define __XINDOWS_CORE_UTIL_TRANSFORM_H__

//+-------------------------------------------------------------------------
//
//  Display units.
//
//--------------------------------------------------------------------------
typedef INT     Z;
typedef Z       X;
typedef Z       Y;

//+-------------------------------------------------------------------------
//
//  Device Resolution.
//
//--------------------------------------------------------------------------
typedef struct _DRES
{
    X dxuInch;
    Y dyuInch;
} DRES;

//+-------------------------------------------------------------------------
//
//  Unit constants
//
//--------------------------------------------------------------------------
const Z ptsInch = 72;   // number of points per inch
const Z twpInch = 1440; // number of twips per inch
const Z himInch = 2540; // number of himetric units per inch

class CTransform
{
public:
    void DocumentFromWindow(POINTL* pptlDoc, long xWin, long yWin) const;
    void DocumentFromWindow(SIZEL* psizelDoc, long cxWin, long cyWin) const;
    void DocumentFromWindow(RECTL* prclDoc, const RECT* prcWin) const;
    void DocumentFromWindow(POINTL* pptlDoc, POINT ptWin) const
    {
        DocumentFromWindow(pptlDoc, ptWin.x, ptWin.y);
    }
    void DocumentFromWindow(SIZEL* psizelDoc, SIZE sizeWin) const
    {
        DocumentFromWindow(psizelDoc, sizeWin.cx, sizeWin.cy);
    }

    long WindowFromDocumentX(long xlDoc) const;
    long WindowFromDocumentY(long ylDoc) const;
    long WindowFromDocumentCX(long xlDoc) const;
    long WindowFromDocumentCY(long ylDoc) const;

    void WindowFromDocument(POINT* pptWin,  long xlDoc, long ylDoc) const;
    void WindowFromDocument(RECT* prcWin, const RECTL* prclDoc) const;
    void WindowFromDocument(SIZE* psizeWin, long cxlDoc, long cylDoc) const;
    void WindowFromDocument(POINT* pptWin, POINTL ptlDoc) const
    {
        WindowFromDocument(pptWin, (int)ptlDoc.x, (int)ptlDoc.y);
    }
    void WindowFromDocument(SIZE* psizeWin, SIZEL sizelDoc) const
    {
        WindowFromDocument(psizeWin, (int)sizelDoc.cx, (int)sizelDoc.cy);
    }

    void HimetricFromDevice(POINTL* pptlDocOut, int cxWinIn, int cyWinIn)
    {
        DocumentFromWindow(pptlDocOut, cxWinIn, cyWinIn);
    }
    void HimetricFromDevice(POINTL* pptlDocOut, POINT ptWinIn)
    {
        DocumentFromWindow(pptlDocOut, ptWinIn);
    }
    void HimetricFromDevice(SIZEL* psizelDocOut, SIZE sizeWinIn) const
    {
        DocumentFromWindow(psizelDocOut, sizeWinIn);
    }
    void HimetricFromDevice(SIZEL*psizelDocOut, int cxWinin, int cyWinIn) const
    {
        DocumentFromWindow(psizelDocOut, cxWinin, cyWinIn);
    }

    long DeviceFromHimetricCY(long hi) const
    {
        return WindowFromDocumentCY(hi);
    }
    long DeviceFromHimetricCX(long hi) const
    {
        return (long)WindowFromDocumentCX((int)hi);
    }

    void DeviceFromHimetric(POINT* pptWinOut, int xlDocIn, int ylDocIn)
    {
        WindowFromDocument(pptWinOut, xlDocIn, ylDocIn);
    }
    void DeviceFromHimetric(POINT* pptWinOut, POINTL ptlDocIn)
    {
        WindowFromDocument(pptWinOut, ptlDocIn);
    }
    void DeviceFromHimetric(SIZE* psizeWinOut, SIZEL sizelDocIn) const
    {
        WindowFromDocument(psizeWinOut, sizelDocIn);
    }
    void DeviceFromHimetric(SIZE* psizeWinOut, int cxlDocIn, int cylDocIn) const
    {
        WindowFromDocument(psizeWinOut, cxlDocIn, cylDocIn);
    }

    void WindowFromDocPixels(POINT* ppt, POINT p, BOOL fRelative=FALSE);
    long WindowXFromDocPixels(long lPixels, BOOL fRelative=FALSE);
    long WindowYFromDocPixels(long lPixels, BOOL fRelative=FALSE);
    long DocPixelsFromWindowX(long lValue, BOOL fRelative=FALSE);
    long DocPixelsFromWindowY(long lValue, BOOL fRelative=FALSE);

    long DeviceFromTwipsCX(long twip) const;
    long DeviceFromTwipsCY(long twip) const;
    long TwipsFromDeviceCX(long pix) const;
    long TwipsFromDeviceCY(long pix) const;

    void Init(const CTransform* pTransform);
    void Init(const RECT* prcDst, SIZEL sizelSrc, const SIZE* psizeInch=NULL);
    void Init(SIZEL sizelSrc);


    // Effective scaled resolution
    void zoom(int wNumerX, int wNumerY, int wDenom);
    void SetScaleFraction(int wNumerX, int wNumerY, int wDenom);
    void SetResolutions(const SIZE* psizeInchTarget, const SIZE* psizeInchRender);
    X    GetScaledDxpInch();
    Y    GetScaledDypInch();
    inline X    GetNumerX()         { return _wNumerX; }
    inline Y    GetNumerY()         { return _wNumerY; }
    inline Z    GetDenom()          { return _wDenom; }
    inline BOOL IsZoomed() const    { return _fScaled; }
    inline BOOL TargetDiffers()     { return _fDiff; }

    // Pixel scaling routines
    //
    // Layout units:            dxt
    // Unzoomed render units:   dxr
    // Zoomed render units:     dxz
    X DxrFromDxt(X dxt) const;
    Y DyrFromDyt(Y dyt) const;
    X DxzFromDxr(X dxr, BOOL fRelative=FALSE) const;
    Y DyzFromDyr(Y dyr, BOOL fRelative=FALSE) const;
    X DxzFromDxt(X dxt, BOOL fRelative=FALSE) const;
    Y DyzFromDyt(Y dyt, BOOL fRelative=FALSE) const;
    void RectzFromRectr(LPRECT pRectZ, LPRECT pRectR);

    // Pixel measurement routines
    //
    // Points:      pts
    // Twips:       twp
    // Himetric:    him
    X DxtFromPts(float pts) const;
    Y DytFromPts(float pts) const;
    X DxtFromTwp(long twp) const;
    Y DytFromTwp(long twp) const;
    X DxtFromHim(long him) const;
    Y DytFromHim(long him) const;

    // Reverse scaling routines
    //
    // Note: going from low-resolution to high-resolution units is almost always
    // a bug when used in measuring code.
    X DxtFromDxr(X dxr) const;
    Y DytFromDyr(Y dyr) const;
    X DxrFromDxz(X dxz, BOOL fRelative=FALSE) const;
    Y DyrFromDyz(Y dyz, BOOL fRelative=FALSE) const;
    X DxtFromDxz(X dxz, BOOL fRelative=FALSE) const;
    Y DytFromDyz(Y dyz, BOOL fRelative=FALSE) const;

    X HimFromDxt(long dxt) const;
    Y HimFromDyt(long dyt) const;
    X TwpFromDxt(long dxt) const;
    Y TwpFromDyt(long dyt) const;

    CTransform() {}
    CTransform(const CTransform* ptransform)
    {
        Init(ptransform);
    }
    CTransform(const CTransform& transform)
    {
        Init(&transform);
    }
    CTransform(const RECT* prcDst, const SIZEL* psizelSrc, const SIZE* psizeInch=NULL)
    {
        Init(prcDst, *psizelSrc, psizeInch);
    }

    SIZE    _sizeSrc;
    POINT   _ptDst;
    SIZE    _sizeDst;
    SIZE    _sizeInch;  // device dots per inch

private:
    void SimplifyScaleFraction();

    X       _wNumerX;   // X numerator of scaling fraction
    Y       _wNumerY;   // Y numerator of scaling fraction
    Z       _wDenom;    // denominator of scaling fraction
    BOOL    _fScaled;   // TRUE if non-unity scaling fraction

    union
    {
        DRES _dresLayout;	// horizontal, vertical layout resolution (may be device-independent)
        struct
        {
            X _dxtInch;		// resolution of target device
            Y _dytInch;		// resolution of target device
        };
    };

    union
    {
        DRES _dresRender;	// horizontal, vertical rendering resolution (for example, display or printer)
        struct
        {
            X _dxrInch;		// resolution of display device
            Y _dyrInch;		// resolution of display device
        };
    };

    BOOL    _fDiff;			// if FALSE, use layout resolution only
};



class CSaveTransform : public CTransform
{
public:
    CSaveTransform(CTransform* ptransform) { Save(ptransform); }
    CSaveTransform(CTransform& transform)  { Save(&transform); }

    ~CSaveTransform() { Restore(); }

private:
    CTransform* _ptransform;

    void Save(CTransform* ptransform)
    {
        _ptransform = ptransform;
        memcpy(this, _ptransform, sizeof(CTransform));
    }
    void Restore()
    {
        memcpy(_ptransform, this, sizeof(CTransform));
    }
};

class CDocument;
class CElement;
//+---------------------------------------------------------------------------
//
//  Class:      CDocInfo
//
//  Purpose:    Document-associated transform
//
//----------------------------------------------------------------------------
class CDocInfo : public CTransform
{
public:
    CDocument* _pDoc;   // Associated CDoc

    CDocInfo()                      { Init(); }
    CDocInfo(const CDocInfo* pdci)  { Init(pdci); }
    CDocInfo(const CDocInfo& dci)   { Init(&dci); }
    CDocInfo(CElement* pElement)    { Init(pElement); }

    void Init(CElement* pElement);
    void Init(const CDocInfo* pdci)
    {
        memcpy(this, pdci, sizeof(CDocInfo));
    }

protected:
    void Init() { _pDoc = NULL; }
};



//+-------------------------------------------------------------------------
//
//  CDrawInfo  Tag = DI
//
//--------------------------------------------------------------------------
class CDrawInfo : public CDocInfo
{
public:
    BOOL            _fInplacePaint;      // True if painting to current inplace location.
    BOOL            _fIsMetafile;        // True if associated HDC is a meta-file
    BOOL            _fIsMemoryDC;        // True if associated HDC is a memory DC
    BOOL            _fHtPalette;         // True if selected palette is the halftone palette
    DWORD           _dwDrawAspect;       // per IViewObject::Draw
    LONG            _lindex;             // per IViewObject::Draw
    void*           _pvAspect;           // per IViewObject::Draw
    DVTARGETDEVICE* _ptd;                // per IViewObject::Draw
    HDC             _hic;                // per IViewObject::Draw
    HDC             _hdc;                // per IViewObject::Draw
    LPCRECTL        _prcWBounds;         // per IViewObject::Draw
    DWORD           _dwContinue;         // per IViewObject::Draw
    BOOL            (CALLBACK *_pfnContinue)(ULONG_PTR); // per IViewObject::Draw
};

#endif //__XINDOWS_CORE_UTIL_TRANSFORM_H__