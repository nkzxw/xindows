
#ifndef __XINDOWS_CORE_UTIL_BUTTONUTIL_H__
#define __XINDOWS_CORE_UTIL_BUTTONUTIL_H__

enum BORDER_FLAGS
{
    BRFLAGS_BUTTON	    = 0x01,
    BRFLAGS_ADJUSTRECT  = 0x02,
    BRFLAGS_DEFAULT	    = 0x04,
    BRFLAGS_MONO	    = 0x08  // Inner border only (for flat scrollbars)
};

enum BUTTON_GLYPH
{
    BG_UP, BG_DOWN, BG_LEFT, BG_RIGHT,
    BG_COMBO, BG_REFEDIT, BG_PLAIN, BG_REDUCE
};

typedef enum _fmBorderStyle
{
    fmBorderStyleNone   = 0,
    fmBorderStyleSingle = 1,
    fmBorderStyleRaised = 2,
    fmBorderStyleSunken = 3,
    fmBorderStyleEtched = 4,
    fmBorderStyleBump   = 5,
    fmBorderStyleRaisedMono = 6,
    fmBorderStyleSunkenMono = 7,
    fmBorderStyleDouble = 8,
    fmBorderStyleDotted = 9,
    fmBorderStyleDashed = 10
} fmBorderStyle;

//+------------------------------------------------------------------------
//
//  Contents:   Button Drawing routines common to scrollbar and dropbutton.
//
//-------------------------------------------------------------------------
HRESULT BRDrawBorder(CDrawInfo* pDI, RECT* prc, fmBorderStyle BorderStyle,
                     COLORREF colorBorder, ThreeDColors* peffectColor, DWORD dwFlags);
int     BRGetBorderWidth(fmBorderStyle BorderStyle);
HRESULT BRAdjustRectForBorder(CTransform* pTransform, RECT* prc, fmBorderStyle BorderStyle);
HRESULT BRAdjustRectForBorderRev(CTransform* pTransform, RECT* prc, fmBorderStyle BorderStyle);
HRESULT BRAdjustRectlForBorder(RECTL* prcl, fmBorderStyle BorderStyle);
HRESULT BRAdjustSizelForBorder(SIZEL* psizel, fmBorderStyle BorderStyle, BOOL fSubtractAdd=0); // 0 - subtract; 1 - add, default is subtract.

// Rect drawing helper.  It tests for an empty rect, and very large rects
// which exceed 32K units.
class CUtilityButton
{
public:    
    void Invalidate(HWND hWnd, const RECT& rc, DWORD dwFlags=0);
    void DrawButton(CDrawInfo*, HWND, BUTTON_GLYPH, BOOL pressed, BOOL enabled, BOOL focused,
        const RECT& rcBounds, const SIZEL& sizelExtent, unsigned long padding);
    virtual ThreeDColors& GetColors(void);

private:
    void DrawNull(HDC hdc, HBRUSH, BUTTON_GLYPH glyph, const RECT& rcBounds, const SIZEL& sizel);
    void DrawArrow(HDC hdc, HBRUSH, BUTTON_GLYPH glyph, const RECT& rcBounds, const SIZEL& sizel);
    void DrawDotDotDot(HDC hdc, HBRUSH, BUTTON_GLYPH glyph, const RECT& rcBounds, const SIZEL& sizel);
    void DrawReduce(HDC hdc, HBRUSH, BUTTON_GLYPH glyph, const RECT& rcBounds, const SIZEL& sizel);

public:
    BOOL _fFlat; // for flat scrollbars/buttons
};

#endif //__XINDOWS_CORE_UTIL_BUTTONUTIL_H__