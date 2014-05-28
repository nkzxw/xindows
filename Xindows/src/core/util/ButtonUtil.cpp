
#include "stdafx.h"
#include "ButtonUtil.h"

const int HAIRLINE_IN_HIMETRICS     = 26;
const int BORDEREFFECT_IN_HIMETRICS = 52;

void DrawEdge2(HDC hdc, LPRECT lprc, UINT edge, UINT flags, ThreeDColors& c3d,
               COLORREF colorBorder, UINT borderXWidth, UINT borderYWidth)
{
    COLORREF	colorTL;
    COLORREF	colorBR;
    RECT		rc;
    RECT		rc2;
    UINT		bdrMask;
    HBRUSH		hbrOld = NULL;
    COLORREF	crNow  = (COLORREF)0xFFFFFFFF;

    rc = *lprc;
    BOOL foutEffect = !(flags&BF_MONO); // No outer border if BF_MONO

    int XIn, XOut=0, YIn, YOut=0; // init XOut, YOut for retail build to pass

    if(foutEffect)
    {
        // border width can be odd number of pixels, so we divide it between
        // the inner and outer effect.
        XIn = borderXWidth / 2;
        XOut = borderXWidth - XIn;

        YIn = borderYWidth / 2;
        YOut = borderYWidth - YIn;
    }
    else
    {
        XIn = borderXWidth;
        YIn = borderYWidth;
    }

    Assert((BDR_OUTER==0x0003) && (BDR_INNER==0x000C));

    Assert(rc.left <= rc.right);
    Assert(rc.top <= rc.bottom);
    if(!(flags & BF_FLAT))
    {
        bdrMask = foutEffect ? BDR_OUTER : BDR_INNER;
        for(; bdrMask<=BDR_INNER; bdrMask<<=2)
        {
            switch(edge & bdrMask)
            {
            case BDR_RAISEDOUTER:
                colorTL = (flags&BF_SOFT) ? c3d.BtnHighLight() : c3d.BtnLight();
                colorBR = c3d.BtnDkShadow();
                break;

            case BDR_RAISEDINNER:
                colorTL = (flags&BF_SOFT) ? c3d.BtnLight() : c3d.BtnHighLight();
                colorBR = c3d.BtnShadow();
                break;

            case BDR_SUNKENOUTER:
                // fButton should be wndframe
                colorTL = (flags&BF_SOFT) ? c3d.BtnDkShadow() : c3d.BtnShadow();
                colorBR = c3d.BtnHighLight();
                break;

            case BDR_SUNKENINNER:
                if(flags & BF_MONO)
                {
                    // inversion of BDR_RAISEDINNER
                    colorTL = c3d.BtnShadow();
                    colorBR = (flags&BF_SOFT) ? c3d.BtnLight() : c3d.BtnHighLight();
                }
                else
                {
                    colorTL = (flags&BF_SOFT) ? c3d.BtnShadow() : c3d.BtnDkShadow();
                    colorBR = c3d.BtnLight();
                }
                break;

            default:
                return;
            }

            if(flags & (BF_RIGHT|BF_BOTTOM))
            {
                if(colorBR != crNow)
                {
                    HBRUSH hbrNew;
                    SelectCachedBrush(hdc, colorBR, &hbrNew, &hbrOld, &crNow);
                }

                if(flags & BF_RIGHT)
                {
                    rc2 = rc;
                    rc2.left = (rc.right -= foutEffect?XOut:XIn);
                    if(rc2.left < rc.left)
                    {
                        rc2.left = rc.left;
                    }
                    PatBlt(hdc, rc2.left, rc2.top, rc2.right-rc2.left, rc2.bottom-rc2.top, PATCOPY);
                }

                if(flags & BF_BOTTOM)
                {
                    rc2 = rc;
                    rc2.top = (rc.bottom -= foutEffect?YOut:YIn);
                    if(rc2.top < rc.top)
                    {
                        rc2.top = rc.top;
                    }
                    PatBlt(hdc, rc2.left, rc2.top, rc2.right-rc2.left, rc2.bottom-rc2.top, PATCOPY);
                }
            }

            if(flags & (BF_LEFT|BF_TOP))
            {
                if(colorTL != crNow)
                {
                    HBRUSH hbrNew;
                    SelectCachedBrush(hdc, colorTL, &hbrNew, &hbrOld, &crNow);
                }

                if(flags & BF_LEFT)
                {
                    rc2 = rc;
                    rc2.right = (rc.left += foutEffect?XOut:XIn);
                    if(rc2.right > rc.right)
                    {
                        rc2.right = rc.right;
                    }
                    PatBlt(hdc, rc2.left, rc2.top, rc2.right-rc2.left, rc2.bottom-rc2.top, PATCOPY);
                }

                if(flags & BF_TOP)
                {
                    rc2 = rc;
                    rc2.bottom = (rc.top += foutEffect?YOut:YIn);
                    if(rc2.bottom > rc.bottom)
                    {
                        rc2.bottom = rc.bottom;
                    }
                    PatBlt(hdc, rc2.left, rc2.top, rc2.right-rc2.left, rc2.bottom-rc2.top, PATCOPY);
                }
            }
            if(foutEffect)
            {
                foutEffect = FALSE;
            }
        }
    }
    else
    {
        if(colorBorder != crNow)
        {
            HBRUSH hbrNew;
            SelectCachedBrush(hdc, colorBorder, &hbrNew, &hbrOld, &crNow);
        }

        if(flags & BF_RIGHT)
        {
            rc2 = rc;
            rc2.left = (rc.right -= borderXWidth);
            if(rc2.left < rc.left)
            {
                rc2.left = rc.left;
            }
            PatBlt(hdc, rc2.left, rc2.top, rc2.right-rc2.left, rc2.bottom-rc2.top, PATCOPY);
        }

        if(flags & BF_BOTTOM)
        {
            rc2 = rc;
            rc2.top = (rc.bottom -= borderYWidth);
            if(rc2.top < rc.top)
            {
                rc2.top = rc.top;
            }
            PatBlt(hdc, rc2.left, rc2.top, rc2.right-rc2.left, rc2.bottom-rc2.top, PATCOPY);
        }

        if(flags & BF_LEFT)
        {
            rc2 = rc;
            rc2.right = (rc.left += borderXWidth);
            if(rc2.right > rc.right)
            {
                rc2.right = rc.right;
            }
            PatBlt(hdc, rc2.left, rc2.top, rc2.right-rc2.left, rc2.bottom-rc2.top, PATCOPY);
        }

        if(flags & BF_TOP)
        {
            rc2 = rc;
            rc2.bottom = (rc.top += borderYWidth);
            if(rc2.bottom > rc.bottom)
            {
                rc2.bottom = rc.bottom;
            }
            PatBlt(hdc, rc2.left, rc2.top, rc2.right-rc2.left, rc2.bottom-rc2.top, PATCOPY);
        }
    }

    if(flags & BF_MIDDLE)
    {
        if(c3d.BtnFace() != crNow)
        {
            HBRUSH hbrNew;
            SelectCachedBrush(hdc, c3d.BtnFace(), &hbrNew, &hbrOld, &crNow);
        }
        PatBlt(hdc, rc.left, rc.top, rc.right-rc.left, rc.bottom-rc.top, PATCOPY);
    }

    if(hbrOld)
    {
        ReleaseCachedBrush((HBRUSH)SelectObject(hdc, hbrOld));
    }
}

HRESULT BRDrawBorder(CDrawInfo* pDI, RECT* prc, fmBorderStyle BorderStyle,
                     COLORREF colorBorder, ThreeDColors* peffectColor, DWORD dwFlags)
{
    ThreeDColors*	ptdc;
    ThreeDColors	tdc;
    UINT			uBdrXWidth;
    UINT			uBdrYWidth;
    UINT			uEdgeStyle=0, uEdgeFlags=0;
    UINT			borderUnit;


    if(peffectColor == NULL)
    {
        ptdc = &tdc; // get the default colors
    }
    else
    {
        ptdc = peffectColor;
    }

    Assert(ptdc != NULL);
    Assert(pDI->_hdc);


    if(dwFlags & BRFLAGS_DEFAULT)
    {
        uBdrYWidth = pDI->WindowFromDocumentCY(HAIRLINE_IN_HIMETRICS);
        uBdrXWidth = pDI->WindowFromDocumentCX(HAIRLINE_IN_HIMETRICS);

        DrawEdge2( pDI->_hdc, (RECT*)prc, 0, BF_FLAT|BF_RECT,
            *ptdc, colorBorder, uBdrXWidth, uBdrYWidth);

        BRAdjustRectForBorder(pDI, (RECT*)prc, fmBorderStyleSingle);
    }

    // draw the beveled edge
    if(BorderStyle > fmBorderStyleSingle)
    {
        borderUnit = (dwFlags&BRFLAGS_MONO) ? HAIRLINE_IN_HIMETRICS : BORDEREFFECT_IN_HIMETRICS;

        switch(BorderStyle)
        {
        case fmBorderStyleRaised:
            uEdgeStyle = EDGE_RAISED;
            break;

        case fmBorderStyleSunken:
            uEdgeStyle = EDGE_SUNKEN;
            break;

        case fmBorderStyleEtched:
            uEdgeStyle = EDGE_ETCHED;
            break;

        case fmBorderStyleBump:
            uEdgeStyle = EDGE_BUMP;
            break;

        default:
            Assert(0 && "CBorderHelper::Render, illegal effect value");
        }

        // Draw the border of the control
        uEdgeFlags = BF_RECT | ((dwFlags&BRFLAGS_BUTTON)?BF_SOFT:0);
        if(dwFlags & BRFLAGS_MONO)
        {
            uEdgeFlags |= BF_MONO;
        }
    }
    else if(BorderStyle == fmBorderStyleSingle)
    {
        borderUnit = HAIRLINE_IN_HIMETRICS;

        uEdgeFlags = BF_FLAT|BF_RECT;
    }
    else
    {
        goto cleanup; // nothing to do
    }

    uBdrYWidth = pDI->WindowFromDocumentCY(borderUnit);
    uBdrXWidth = pDI->WindowFromDocumentCX(borderUnit);

    // border width can not go below 1 pixel
    if(uBdrYWidth == 0)
    {
        uBdrYWidth = 1;
    }

    if(uBdrXWidth == 0)
    {
        uBdrXWidth = 1;
    }

    // Draw the border
    // BUGBUG DrawEdge2  currently checks for underflow of prc
    DrawEdge2(pDI->_hdc, (RECT*)prc, uEdgeStyle, uEdgeFlags,
        *ptdc, colorBorder, uBdrXWidth, uBdrYWidth);

    if(dwFlags & BRFLAGS_ADJUSTRECT)
    {
        BRAdjustRectForBorder(pDI, prc, BorderStyle);
    }

cleanup:
    return S_OK;
}

int BRGetBorderWidth(fmBorderStyle BorderStyle)
{
    int uBorderWidth;

    // count for the beveled edge
    if(BorderStyle > fmBorderStyleSingle)
    {
        uBorderWidth = BORDEREFFECT_IN_HIMETRICS;
    } // if no effect count for a single (for now) border line
    else if(BorderStyle == fmBorderStyleSingle)
    {
        uBorderWidth = HAIRLINE_IN_HIMETRICS;
    }
    else if(BorderStyle == fmBorderStyleNone)
    {
        uBorderWidth = 0;
    } else
    {
        uBorderWidth = -1;
    }

    return uBorderWidth;
}

HRESULT BRAdjustRectForBorderActual(CTransform* pTransform, RECT* prc, fmBorderStyle BorderStyle, BOOL fInflateForBorder)
{
    int uInflateXBy, uInflateYBy;
    UINT inflateUnit;

    Assert(prc);

    // count for the beveled edge
    if(BorderStyle > fmBorderStyleSingle)
    {
        // inflateUnit is 52 HiMetrics
        inflateUnit =  BORDEREFFECT_IN_HIMETRICS;
    } // if no effect count for a single (for now) border line
    else if(BorderStyle == fmBorderStyleSingle)
    {
        // inflateUnit = 26 HiMetrics.
        inflateUnit = HAIRLINE_IN_HIMETRICS;
    }
    else
    {
        goto Cleanup;
    }
    // compute the actually border dimantions
    uInflateXBy = pTransform->WindowFromDocumentCX(inflateUnit);
    uInflateYBy = pTransform->WindowFromDocumentCY(inflateUnit);

    // Border width can't go below 1 pixel
    if(uInflateXBy == 0)
    {
        uInflateXBy = 1;
    }

    if(uInflateYBy == 0)
    {
        uInflateYBy = 1;
    }

    // Compute border size with zooming.
    if(fInflateForBorder)
    {
        InflateRect((RECT*)prc, uInflateXBy, uInflateYBy );
    }
    else
    {
        InflateRect((RECT*)prc, -uInflateXBy, -uInflateYBy );
    }

    if(prc->left > prc->right)
    {
        prc->left = prc->right;
    }
    if(prc->top > prc->bottom)
    {
        prc->top = prc->bottom;
    }

Cleanup:
    return S_OK;
}

HRESULT BRAdjustRectForBorder(CTransform* pTransform, RECT* prc, fmBorderStyle BorderStyle)
{
    return BRAdjustRectForBorderActual(pTransform, prc, BorderStyle, FALSE);
}

HRESULT BRAdjustRectForBorderRev(CTransform* pTransform, RECT* prc, fmBorderStyle BorderStyle)
{
    return BRAdjustRectForBorderActual(pTransform, prc, BorderStyle, TRUE);
}

HRESULT BRAdjustRectlForBorder(RECTL* prcl, fmBorderStyle BorderStyle)
{
    if(BorderStyle > fmBorderStyleSingle)
    {
        // Compute border size with zooming.
        InflateRect((RECT*)prcl, -BORDEREFFECT_IN_HIMETRICS, -BORDEREFFECT_IN_HIMETRICS);
    }
    // if no effect count for a single (for now) border line
    else if(BorderStyle == fmBorderStyleSingle)
    {
        // Compute border size with zooming.
        InflateRect((RECT*)prcl, -HAIRLINE_IN_HIMETRICS, -HAIRLINE_IN_HIMETRICS);
    }

    return S_OK;
}

HRESULT BRAdjustSizelForBorder(SIZEL* psizel, fmBorderStyle BorderStyle, BOOL fSubtractAdd)
{
    int iSubtractAdd = fSubtractAdd? 1: -1;

    if(BorderStyle > fmBorderStyleSingle)
    {
        psizel->cx += iSubtractAdd * 2 * BORDEREFFECT_IN_HIMETRICS;
        psizel->cy += iSubtractAdd * 2 * BORDEREFFECT_IN_HIMETRICS;
    }
    // if no effect count for a single (for now) border line
    else if(BorderStyle == fmBorderStyleSingle)
    {
        psizel->cx += iSubtractAdd * 2 * HAIRLINE_IN_HIMETRICS;
        psizel->cy += iSubtractAdd * 2 * HAIRLINE_IN_HIMETRICS;
    }

    // Don't allow a negative size
    psizel->cx = max(psizel->cx, 0L);
    psizel->cy = max(psizel->cy, 0L);

    return S_OK;
}

void DrawRect(HDC hdc, HBRUSH hbr, int x1, int y1, int x2, int y2, BOOL fReleaseBrush=TRUE)
{
    if(x2>x1 && y2>y1)
    {
        HBRUSH hbrOld = (HBRUSH)SelectObject(hdc, hbr);
        PatBlt(hdc, x1, y1, x2-x1, y2-y1, PATCOPY);
        SelectObject(hdc, hbrOld);
    }

    if(fReleaseBrush)
    {
        ReleaseCachedBrush(hbr);
    }
}

void CUtilityButton::Invalidate(HWND hWnd, const RECT &rc, DWORD dwFlags)
{
    if(hWnd)
    {
        RedrawWindow(hWnd, &rc, NULL, RDW_INVALIDATE);
    }
}

void CUtilityButton::DrawNull(HDC hdc, HBRUSH hbr, BUTTON_GLYPH glyph, const RECT& rcBounds, const SIZEL& sizel)
{
}

void CUtilityButton::DrawDotDotDot(HDC hdc, HBRUSH hbr, BUTTON_GLYPH glyph, const RECT& rcBounds, const SIZEL& sizel)
{
    long xStart, yStart, width, height;

    Assert(rcBounds.right > rcBounds.left);
    Assert(rcBounds.bottom > rcBounds.top);

    // 106 is about 3 points, which is how far from the bottom
    // the dots starts.
    //
    // 53 is about 1.5 points, which the height/width of the dots
    yStart = rcBounds.bottom - MulDiv(rcBounds.bottom-rcBounds.top, 106, sizel.cy);

    height = MulDiv(rcBounds.bottom-rcBounds.top, 53, sizel.cy);
    width  = MulDiv(rcBounds.right-rcBounds.left, 54, sizel.cx);

    xStart = rcBounds.left + (rcBounds.right-rcBounds.left)*2/13;
    DrawRect(hdc, hbr, xStart, yStart, xStart+width, yStart+height, FALSE);

    xStart = rcBounds.left + (rcBounds.right-rcBounds.left)*5/13;
    DrawRect(hdc, hbr, xStart, yStart, xStart+width, yStart+height, FALSE);

    xStart = rcBounds.left + (rcBounds.right-rcBounds.left)*8/13;
    DrawRect(hdc, hbr, xStart, yStart, xStart+width, yStart+height, FALSE);
}

void CUtilityButton::DrawReduce(HDC hdc, HBRUSH hbr, BUTTON_GLYPH glyph, const RECT& rcBounds, const SIZEL& sizel)
{
    long xStart, yStart, width, height;

    Assert(rcBounds.right > rcBounds.left);
    Assert(rcBounds.bottom > rcBounds.top);

    // 106 is about 3 points, which is how far from the bottom
    // the bar starts.
    //
    // 53 is about 1.5 points, which the height of the bar
    yStart = rcBounds.bottom - MulDiv( rcBounds.bottom-rcBounds.top, 108, sizel.cy );

    height = MulDiv(rcBounds.bottom-rcBounds.top, 54, sizel.cy);

    xStart = rcBounds.left + (rcBounds.right-rcBounds.left)*3/13;
    width  = (rcBounds.right - rcBounds.left) * 6 / 13;

    DrawRect(hdc, hbr, xStart, yStart, xStart+width, yStart+height, FALSE);
}

void CUtilityButton::DrawArrow(HDC hdc, HBRUSH hbr, BUTTON_GLYPH glyph, const RECT& rcBounds, const SIZEL& sizel)
{
    long i;

    Assert(rcBounds.right > rcBounds.left);
    Assert(rcBounds.bottom > rcBounds.top);

    Assert(glyph==BG_UP || glyph==BG_DOWN || glyph==BG_LEFT || glyph==BG_RIGHT);

    if(glyph==BG_UP || glyph==BG_DOWN)
    {
        // Determine the height of the arrow by computing the largest
        // arrows we allow for the given width and height, and then taking
        // the smaller of the two.  Also make sure it is non zero.
        long arrow_height = max((long)1,
            (long)min((rcBounds.bottom-rcBounds.top+2)/3,
            ((rcBounds.right-rcBounds.left)*5/8+1)/2));

        // Locate where the top of the arrow starts and where it is
        // centered horizontally
        long sy = rcBounds.top + (rcBounds.bottom-rcBounds.top+1-arrow_height)/2;

        long cx = rcBounds.left + (rcBounds.right-rcBounds.left-1)/2;

        // Draw the arrow from top to bottom in successive strips
        for(i=0; i<arrow_height; i++)
        {
            long y = glyph==BG_UP ? sy+i : sy+arrow_height-i-1;

            DrawRect(hdc, hbr, cx-i, y, cx-i+1+i*2, y+1, FALSE);
        }
    }
    else    
    {
        // Determine the width of the arrow by computing the largest
        // arrows we allow for the given width and height, and then taking
        // the smaller of the two.  Also make sure it iz non zero.
        long arrow_width = max((long)1,
            (long)min((rcBounds.right-rcBounds.left+2)/3,
            ((rcBounds.bottom-rcBounds.top)*5/8+1)/2));

        // Locate where the left of the arrow starts and where it is
        // centered vertically
        long sx = rcBounds.left + (rcBounds.right-rcBounds.left+1-arrow_width)/2;

        long cy = rcBounds.top + (rcBounds.bottom-rcBounds.top)/2;

        // Draw the arrow from top to bottom in successive strips
        for(long i=0; i<arrow_width; i++)
        {
            long x = glyph==BG_LEFT ? sx+i : sx+arrow_width-i-1;

            DrawRect(hdc, hbr, x, cy-i, x+1, cy+i+1, FALSE);
        }
    }
}

void CUtilityButton::DrawButton(CDrawInfo* pDI, HWND hWnd, BUTTON_GLYPH glyph,
                                BOOL pressed, BOOL enabled, BOOL focused,
                                const RECT& rcBounds, const SIZEL& sizelExtent, unsigned long padding)
{
    if(hWnd)
    {
        Invalidate(hWnd, rcBounds);
    }
    else
    {
        void (CUtilityButton::*pmfDraw)(HDC, HBRUSH, BUTTON_GLYPH, const RECT&, const SIZEL&);

        ThreeDColors&   colors = GetColors();
        RECT		    rcGlyph = rcBounds;
        SIZEL		    sizelGlyph = sizelExtent;
        long		    dx=0, dy=0, xOffset=0, yOffset=0;

        // must have at least hdc
        Assert(pDI->_hdc);
        // should come in initialized
        Assert(pDI->_sizeSrc.cx); 

        // First, draw the border around the glyph and the background.
        if(pressed && !_fFlat)
        {
            RECT rcFill = rcBounds;

            // Draw a "single line" border then reduce rect by that size and
            // Fill the rest with the button background
            BRDrawBorder(pDI, LPRECT(&rcBounds), fmBorderStyleSingle, colors.BtnShadow(), 0, 0);
            BRAdjustRectForBorder(pDI, &rcFill, fmBorderStyleSingle);
            DrawRect(pDI->_hdc, colors.BrushBtnFace(),
                rcFill.left, rcFill.top, rcFill.right, rcFill.bottom);

            // Now, compute the reduced rect as if a sunken border were drawn.
            // This leaves the glyph rect in rcGlyph
            BRAdjustRectForBorder(pDI, &rcGlyph, fmBorderStyleRaised);
        }
        else
        {
            // Draw the sunken border and fill the rest with the button bg.
            // This leaves the arrow rect in rcGlyph

            // If we come here with fPressed = TRUE, this must be a flat scrollbar
            BRDrawBorder(pDI, LPRECT(&rcBounds), 
                (pressed)?fmBorderStyleSunken:fmBorderStyleRaised,
                0, &colors, (_fFlat)?BRFLAGS_MONO:0);
            BRAdjustRectForBorder(pDI, &rcGlyph, (_fFlat)?fmBorderStyleSingle:fmBorderStyleRaised);
            DrawRect(pDI->_hdc, colors.BrushBtnFace(),
                rcGlyph.left, rcGlyph.top, rcGlyph.right, rcGlyph.bottom);
        }

        // See if we have a null rect.
        if(rcGlyph.right<=rcGlyph.left || rcGlyph.bottom<=rcGlyph.top)
        {
            goto Cleanup;
        }

        // Adjust the extent to reflect the border
        BRAdjustSizelForBorder(&sizelGlyph, fmBorderStyleSunken);

        // A combo glyph looks like a down arrow
        if(glyph == BG_COMBO)
        {
            glyph = BG_DOWN;
        }

        // Select the draw member.  Default to arrow.
        pmfDraw = &CUtilityButton::DrawArrow;

        switch(glyph)
        {
        case BG_PLAIN:
            pmfDraw = &CUtilityButton::DrawNull;
            break;

            // Adjust the rect for padding.
            //
            // BUGBUG: the padding has not been zoomed....
        case BG_DOWN:	rcGlyph.bottom	-= padding; break;
        case BG_UP:		rcGlyph.top		+= padding; break;
        case BG_LEFT:	rcGlyph.left	+= padding; break;
        case BG_RIGHT:	rcGlyph.right	-= padding; break;

        case BG_REFEDIT:
            pmfDraw = &CUtilityButton::DrawDotDotDot;
            break;

        case BG_REDUCE:
            pmfDraw = &CUtilityButton::DrawReduce;
            break;

        default:
            Assert(0 && "Unknown glyph");
            goto Cleanup;
        }

        // In the rest of the code, if the scrollbar is flat, ignore
        // fPressed (no offset even if the button is pressed)
        if(_fFlat)
        {
            pressed = FALSE;
        }

        // Now that the glyph draw member has been selected, use it to draw.
        //
        //
        // If we are pressed or disabled, we need to compute the offset
        // for pressing or the highlight offset version of the glyph.
        //
        // Also, if we have to draw a focus, we need to know how much to 
        // offset the focus rect from the border.
        if(!enabled || pressed || focused)
        {
            // The offset is 27 HIMETRICS (~ 3/4 points).
            xOffset = MulDiv(rcGlyph.right-rcGlyph.left, 27, sizelGlyph.cx);
            yOffset = MulDiv(rcGlyph.bottom-rcGlyph.top, 27, sizelGlyph.cy);

            if(!enabled || pressed)
            {
                dx = xOffset;
                dy = yOffset;
            }
        }

        if(enabled)
        {
            RECT rc = rcGlyph;
            HBRUSH hbr = colors.BrushBtnText();

            if(pressed)
            {
                rc.left   += dx;
                rc.right  += dx;
                rc.top    += dy;
                rc.bottom += dy;
            }

            CALL_METHOD(this, pmfDraw, (pDI->_hdc, hbr, glyph, rc, sizelGlyph));

            ReleaseCachedBrush(hbr);
        }
        else
        {
            HBRUSH hbr;
            // Here we draw the disabled glyph.  First, draw the lighter
            // back ground version, the the darker foreground version.
            RECT rcBack =
            {
                rcGlyph.left+dx, rcGlyph.top+dy,
                rcGlyph.right+dx, rcGlyph.bottom+dy
            };

            hbr = colors.BrushBtnHighLight();
            CALL_METHOD(this, pmfDraw, (pDI->_hdc, hbr, glyph, rcBack, sizelGlyph));
            ReleaseCachedBrush( hbr );

            hbr = colors.BrushBtnShadow();
            CALL_METHOD(this, pmfDraw, (pDI->_hdc, hbr, glyph, rcGlyph, sizelGlyph));
            ReleaseCachedBrush(hbr);
        }

        // Draw any focus rect

        // BUGBUG: This relies heavily on the assumption that the arrow which is
        //         drawn leaves alot of space around the edge so that the focus
        //         can be drawn there.  Also, the focus rect is not zoomed here.
        if(focused)
        {
            RECT rcFocus = rcGlyph;

            rcFocus.left   += xOffset + dx;
            rcFocus.right  -= xOffset - dx;
            rcFocus.top    += yOffset + dy;
            rcFocus.bottom -= yOffset - dy;

            rcFocus.right  = max(rcFocus.right, rcFocus.right);
            rcFocus.bottom = max(rcFocus.bottom, rcFocus.top);

            if(rcFocus.right<=rcGlyph.right && rcFocus.bottom<=rcGlyph.bottom)
            {
                ::DrawFocusRect(pDI->_hdc, &rcFocus);
            }
        }
    }
Cleanup:
    return;
}

ThreeDColors& CUtilityButton::GetColors()
{
    static ThreeDColors colorsDefault(DEFAULT_BASE_COLOR);
    return colorsDefault;
}