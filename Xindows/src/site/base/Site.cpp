
#include "stdafx.h"
#include "Site.h"

#define _cxx_
#include "../gen/csite.hdl"


HRESULT CSite::focus()
{
    return super::focus();
}

HRESULT CSite::blur()
{
    return super::blur();
}

HRESULT CSite::get_clientHeight(long* p)
{
    return super::get_clientHeight(p);
}

HRESULT CSite::get_clientWidth(long* p)
{
    return super::get_clientWidth(p);
}

HRESULT CSite::get_clientTop(long* p)
{
    return super::get_clientTop(p);
}

HRESULT CSite::get_clientLeft(long* p)
{
    return super::get_clientLeft(p);
}


//+----------------------------------------------------------------------------
//
// Function:    CalcImgBgRect
//
// Synopsis:    Finds the rectangle to pass to Tile() to draw correct
//              background image with the attributes specified in the
//              fancy format (repeat-x, repeat-y, etc).
//
//-----------------------------------------------------------------------------
void CalcBgImgRect(
           CTreeNode*       pNode,
           CFormDrawInfo*   pDI,
           const SIZE*      psizeObj,
           const SIZE*      psizeImg,
           CPoint*          pptBackOrig,
           RECT*            prcBackClip)
{
    // pNode is used to a) extract formats to get the
    // background position values, and b) to get the font height so we
    // can handle em/en/ex units for position.
    const CFancyFormat* pFF = pNode->GetFancyFormat();

    // N.B. Per CSS spec, percentages work as follows:
    // (x%, y%) means that the (x%, y%) point in the image is
    // positioned at the (x%, y%) point in the bounded rectangle.
    if(pFF->_cuvBgPosX.GetUnitType() == CUnitValue::UNIT_PERCENT)
    {
        pptBackOrig->x =
            MulDivQuick(pFF->_cuvBgPosX.GetPercent(),
            psizeObj->cx - psizeImg->cx, 100);
    }
    else
    {
        pptBackOrig->x =
            pFF->_cuvBgPosX.GetPixelValue(pDI, CUnitValue::DIRECTION_CX, 0,
            pNode->GetFontHeightInTwips((CUnitValue*)&pFF->_cuvBgPosX));
    }

    if(pFF->_cuvBgPosY.GetUnitType() == CUnitValue::UNIT_PERCENT)
    {
        pptBackOrig->y =
            MulDivQuick(pFF->_cuvBgPosY.GetPercent(),
            psizeObj->cy - psizeImg->cy,
            100);
    }
    else
    {
        pptBackOrig->y =
            pFF->_cuvBgPosY.GetPixelValue(pDI, CUnitValue::DIRECTION_CY, 0,
            pNode->GetFontHeightInTwips((CUnitValue*)&pFF->_cuvBgPosY));
    }


    if(pFF->_fBgRepeatX)
    {
        prcBackClip->left  = 0;
        prcBackClip->right = psizeObj->cx;
    }
    else
    {
        prcBackClip->left  = pptBackOrig->x;
        prcBackClip->right = pptBackOrig->x + psizeImg->cx;
    }

    if(pFF->_fBgRepeatY)
    {
        prcBackClip->top    = 0;
        prcBackClip->bottom = psizeObj->cy;
    }
    else
    {
        prcBackClip->top    = pptBackOrig->y;
        prcBackClip->bottom = pptBackOrig->y + psizeImg->cy;
    }
}

//+---------------------------------------------------------------
//
//  Member:     CSite::CSite
//
//  Synopsis:   Normal constructor.
//
//  Arguments:  pParent  Site that's our parent
//
//---------------------------------------------------------------
CSite::CSite(ELEMENT_TAG etag, CDocument* pDoc) : CElement(etag, pDoc)
{
    _fSite = TRUE;
}

HRESULT CSite::Init()
{
    HRESULT hr;

    hr = super::Init();

    if(hr)
    {
        goto Cleanup;
    }

    // Root's layout is created, if required, outside Init().
    if(Tag() != ETAG_ROOT)
    {
        CreateLayout();
    }

    Assert(_fLayoutAlwaysValid);

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CSite::PrivateQueryInterface, IUnknown
//
//  Synopsis:   Private unknown QI.
//
//-------------------------------------------------------------------------
HRESULT CSite::PrivateQueryInterface(REFIID iid, void** ppv)
{
    RRETURN(super::PrivateQueryInterface(iid, ppv));
}



void GetBorderColorInfoHelper(CTreeNode* pNodeContext, CDocInfo* pdci, CBorderInfo* pborderinfo)
{
    const CFancyFormat* pFF = pNodeContext->GetFancyFormat();
    const CCharFormat* pCF = pNodeContext->GetCharFormat();
    int i;
    COLORREF clr, clrHilight, clrLight, clrDark, clrShadow;

    for(i=0; i<4 ; i++)
    {
        BOOL fNeedSysColor = FALSE;
        // Get the base color
        if(!pFF->_ccvBorderColors[i].IsDefined())
        {
            clr = pCF->_ccvTextColor.GetColorRef();
            fNeedSysColor = TRUE;
        }
        else
        {
            clr = pFF->_ccvBorderColors[i].GetColorRef();
        }

        // Set up the flat color
        pborderinfo->acrColors[i][1] = clr;

        // Set up the inner and outer colors
        switch(pborderinfo->abStyles[i])
        {
        case fmBorderStyleNone:
        case fmBorderStyleDouble:
        case fmBorderStyleSingle:
        case fmBorderStyleDotted:
        case fmBorderStyleDashed:
            pborderinfo->acrColors[i][0] = pborderinfo->acrColors[i][2] = clr;
            // Don't need inner/outer colors
            break;

        default:
            {
                // Set up the color variations
                if(pFF->_ccvBorderColorHilight.IsDefined() && !ISBORDERSIDECLRSETUNIQUE(pFF, i))
                {
                    clrHilight = pFF->_ccvBorderColorHilight.GetColorRef();
                }
                else
                {
                    if(fNeedSysColor)
                    {
                        clrHilight = GetSysColorQuick(COLOR_BTNHIGHLIGHT);
                    }
                    else
                        clrHilight = clr;
                }
                if(pFF->_ccvBorderColorDark.IsDefined() && !ISBORDERSIDECLRSETUNIQUE(pFF, i))
                {
                    clrDark = pFF->_ccvBorderColorDark.GetColorRef();
                }
                else
                {
                    if(fNeedSysColor)
                    {
                        clrDark = GetSysColorQuick(COLOR_3DDKSHADOW);
                    }
                    else
                    {
                        clrDark = (((clr&0xff0000)>>1)&0xff0000 ) |
                            (((clr&0x00ff00)>>1)&0x00ff00) | (((clr&0x0000ff)>>1)&0x0000ff);
                    }
                }
                if(pFF->_ccvBorderColorShadow.IsDefined() && !ISBORDERSIDECLRSETUNIQUE(pFF, i))
                {
                    clrShadow = pFF->_ccvBorderColorShadow.GetColorRef();
                }
                else
                {
                    if(fNeedSysColor)
                    {
                        clrShadow = GetSysColorQuick(COLOR_BTNSHADOW);
                    }
                    else
                    {
                        clrShadow = (((clr&0xff0000)>>2)&0xff0000) |
                            (((clr&0x00ff00)>>2)&0x00ff00) | (((clr&0x0000ff)>>2)&0x0000ff);
                    }
                }

                // If the Light color isn't set synthesise a light color 3/4 of clr
                if(pFF->_ccvBorderColorLight.IsDefined() && !ISBORDERSIDECLRSETUNIQUE(pFF, i))
                {
                    clrLight = pFF->_ccvBorderColorLight.GetColorRef();
                }
                else
                {
                    if(fNeedSysColor)
                    {
                        clrLight = GetSysColorQuick(COLOR_BTNFACE);
                    }
                    else
                    {
                        clrLight = clrShadow + clrDark;
                    }
                }

                if(i==BORDER_TOP || i==BORDER_LEFT)
                {
                    // Top/left edges
                    if(pFF->_bBorderSoftEdges || (pborderinfo->wEdges&BF_SOFT))
                    {
                        switch(pborderinfo->abStyles[i])
                        {
                        case fmBorderStyleRaisedMono:
                            pborderinfo->acrColors[i][0] = pborderinfo->acrColors[i][2] = clrHilight;
                            break;
                        case fmBorderStyleSunkenMono:
                            pborderinfo->acrColors[i][0] = pborderinfo->acrColors[i][2] = clrDark;
                            break;
                        case fmBorderStyleRaised:
                            pborderinfo->acrColors[i][0] = clrHilight;
                            pborderinfo->acrColors[i][2] = clrLight;
                            break;
                        case fmBorderStyleSunken:
                            pborderinfo->acrColors[i][0] = clrDark;
                            pborderinfo->acrColors[i][2] = clrShadow;
                            break;
                        case fmBorderStyleEtched:
                            pborderinfo->acrColors[i][0] = clrDark;
                            pborderinfo->acrColors[i][2] = clrLight;
                            break;
                        case fmBorderStyleBump:
                            pborderinfo->acrColors[i][0] = clrHilight;
                            pborderinfo->acrColors[i][2] = clrShadow;
                            break;
                        }
                    }
                    else
                    {
                        switch(pborderinfo->abStyles[i])
                        {
                        case fmBorderStyleRaisedMono:
                            pborderinfo->acrColors[i][0] = pborderinfo->acrColors[i][2] = clrLight;
                            break;
                        case fmBorderStyleSunkenMono:
                            pborderinfo->acrColors[i][0] = pborderinfo->acrColors[i][2] = clrShadow;
                            break;
                        case fmBorderStyleRaised:
                            pborderinfo->acrColors[i][0] = clrLight;
                            pborderinfo->acrColors[i][2] = clrHilight;
                            break;
                        case fmBorderStyleSunken:
                            pborderinfo->acrColors[i][0] = clrShadow;
                            pborderinfo->acrColors[i][2] = clrDark;
                            break;
                        case fmBorderStyleEtched:
                            pborderinfo->acrColors[i][0] = clrShadow;
                            pborderinfo->acrColors[i][2] = clrHilight;
                            break;
                        case fmBorderStyleBump:
                            pborderinfo->acrColors[i][0] = clrLight;
                            pborderinfo->acrColors[i][2] = clrDark;
                            break;
                        }
                    }
                }
                else
                {
                    // Bottom/right edges
                    switch(pborderinfo->abStyles[i])
                    {
                    case fmBorderStyleRaisedMono:
                        pborderinfo->acrColors[i][0] = pborderinfo->acrColors[i][2] = clrDark;
                        break;
                    case fmBorderStyleSunkenMono:
                        pborderinfo->acrColors[i][0] = pborderinfo->acrColors[i][2] = clrHilight;
                        break;
                    case fmBorderStyleRaised:
                        pborderinfo->acrColors[i][0] = clrDark;
                        pborderinfo->acrColors[i][2] = clrShadow;
                        break;
                    case fmBorderStyleSunken:
                        pborderinfo->acrColors[i][0] = clrHilight;
                        pborderinfo->acrColors[i][2] = clrLight;
                        break;
                    case fmBorderStyleEtched:
                        pborderinfo->acrColors[i][0] = clrHilight;
                        pborderinfo->acrColors[i][2] = clrShadow;
                        break;
                    case fmBorderStyleBump:
                        pborderinfo->acrColors[i][0] = clrDark;
                        pborderinfo->acrColors[i][2] = clrLight;
                        break;
                    }
                }
            }
        }
    }

    if(pFF->_bBorderSoftEdges)
    {
        pborderinfo->wEdges |= BF_SOFT;
    }
}

//+---------------------------------------------------------------------------
//
//  Function:   CompareElementsByZIndex
//
//  Synopsis:   Comparison function used by qsort to compare the zIndex of
//              two elements.
//
//----------------------------------------------------------------------------
#define ELEMENT1_ONTOP  1
#define ELEMENT2_ONTOP  -1
#define ELEMENTS_EQUAL  0

int CompareElementsByZIndex(const void* pv1, const void* pv2)
{
    int         i, z1, z2;
    HWND        hwnd1, hwnd2;
    CElement*   pElement1 = *(CElement**)pv1;
    CElement*   pElement2 = *(CElement**)pv2;

    // Only compare elements which have the same ZParent
    // BUGBUG: For now, since table elements (e.g., TDs, TRs, CAPTIONs) cannot be
    //         positioned, it is Ok if they all end up in the same list - even if
    //         their ZParent-age is different.
    //         THIS MUST BE RE-VISITED ONCE WE SUPPORT POSITIONING ON TABLE ELEMENTS.
    //         (brendand)
    Assert(pElement1->GetFirstBranch()->ZParent()==pElement2->GetFirstBranch()->ZParent()
        || (pElement1->GetFirstBranch()->ZParent()->Tag()==ETAG_TR
        && pElement2->GetFirstBranch()->ZParent()->Tag()==ETAG_TR)
        || (pElement1->GetFirstBranch()->ZParent()->Tag()==ETAG_TABLE
        && pElement2->GetFirstBranch()->ZParent()->Tag()==ETAG_TR)
        || (pElement2->GetFirstBranch()->ZParent()->Tag()==ETAG_TABLE
        && pElement1->GetFirstBranch()->ZParent()->Tag()==ETAG_TR));

    // Sites with windows are _always_ above sites without windows.
    hwnd1 = pElement1->GetHwnd();
    hwnd2 = pElement2->GetHwnd();

    if((hwnd1==NULL) != (hwnd2==NULL))
    {
        return (hwnd1!=NULL)?ELEMENT1_ONTOP:ELEMENT2_ONTOP;
    }

    // If one element contains the other, then the containee is on top.
    //
    // Since table cells cannot be positioned, we ignore any case where
    // something is contained inside a table cell. That way they essentially
    // become 'peers' of the cells and can be positioned above or below them.
    //
    // BUGBUG (lylec) -- These elements are already scoped, so there's a
    // possibility we might walk up the wrong tree! This should be fixed
    // when Eric's proposed change is made regarding using Tree Nodes
    // instead of elements in the tree.
    if(pElement1->Tag()!=ETAG_TD && pElement2->Tag()!=ETAG_TD &&    // Cell
        pElement1->Tag()!=ETAG_TH && pElement2->Tag()!=ETAG_TH &&   // Header
        pElement1->Tag()!=ETAG_TC && pElement2->Tag()!=ETAG_TC)     // Caption
    {
        if(pElement1->GetFirstBranch()->SearchBranchToRootForScope(pElement2))
        {
            return ELEMENT1_ONTOP;
        }

        if(pElement2->GetFirstBranch()->SearchBranchToRootForScope(pElement1))
        {
            return ELEMENT2_ONTOP;
        }
    }

    // Only pay attention to the z-index attribute if the element is positioned
    //
    // The higher z-index is on top, which means the higher z-index value
    // is "greater".
    z1 = !pElement1->IsPositionStatic()
        ? pElement1->GetFirstBranch()->GetCascadedzIndex() : 0;

    z2 = !pElement2->IsPositionStatic()
        ? pElement2->GetFirstBranch()->GetCascadedzIndex() : 0;

    i = z1 - z2;

    if(i==ELEMENTS_EQUAL &&
        pElement1->IsPositionStatic()!=pElement2->IsPositionStatic())
    {
        // The non-static element has a z-index of 0, so we must make
        // sure it stays above anything in the flow (static).
        i = (!pElement1->IsPositionStatic()) ? ELEMENT1_ONTOP : ELEMENT2_ONTOP;
    }

    // Make sure that the source indices are up to date before accessing them
    Assert(pElement1->Doc() == pElement2->Doc());

    // If the zindex is the same, then sort by source order.
    //
    // Later in the source is on top, which means the higher source-index
    // value is "greater".
    if(i == ELEMENTS_EQUAL)
    {
        i = pElement1->GetSourceIndex() - pElement2->GetSourceIndex();
    }

    // Different elements should never be exactly equal.
    //
    // If this assert fires it's likely due to the element collection not
    // having been built yet.
    Assert(i!=ELEMENTS_EQUAL || pElement1==pElement2);

    return i;
}

DWORD GetBorderInfoHelper(
              CTreeNode*      pNodeContext,
              CDocInfo*       pdci,
              CBorderInfo*    pborderinfo,
              BOOL            fAll/*=FALSE*/)
{
    Assert(pNodeContext);
    Assert(pborderinfo);

    int   i;
    int   iBorderWidth = 0;
    const CFancyFormat* pFF = pNodeContext->GetFancyFormat();
    const CCharFormat* pCF = pNodeContext->GetCharFormat();

    Assert(pFF && pCF);

    for(i=BORDER_TOP; i<=BORDER_LEFT; i++)
    {
        if(pFF->_bBorderStyles[i] != (BYTE)-1)
        {
            pborderinfo->abStyles[i] = pFF->_bBorderStyles[i];
        }
        if(!pborderinfo->abStyles[i])
        {
            pborderinfo->aiWidths[i] = 0;
            continue;
        }

        switch(pFF->_cuvBorderWidths[i].GetUnitType())
        {
        case CUnitValue::UNIT_NULLVALUE:
            continue;
        case CUnitValue::UNIT_ENUM:
            {
                // Pick up the default border width here.
                CUnitValue cuv;
                cuv.SetValue((pFF->_cuvBorderWidths[i].GetUnitValue()+1)*2, CUnitValue::UNIT_PIXELS);
                iBorderWidth = cuv.GetPixelValue(NULL,
                    ((i==BORDER_TOP)||(i==BORDER_BOTTOM))
                    ? CUnitValue::DIRECTION_CY
                    : CUnitValue::DIRECTION_CX,
                    0, pCF->_yHeight);
            }
            break;
        default:
            iBorderWidth = pFF->_cuvBorderWidths[i].GetPixelValue(NULL,
                ((i==BORDER_TOP)||(i==BORDER_BOTTOM))
                ? CUnitValue::DIRECTION_CY
                : CUnitValue::DIRECTION_CX,
                0, pCF->_yHeight);

            // If user sets tiny borderwidth, set smallest width possible (1px) instead of zero (IE5,5865).
            if(!iBorderWidth && pFF->_cuvBorderWidths[i].GetUnitValue()>0)
            {
                iBorderWidth = 1;
            }
        }
        if(iBorderWidth >= 0)
        {
            pborderinfo->aiWidths[i] = iBorderWidth<=MAX_BORDER_SPACE ? iBorderWidth : MAX_BORDER_SPACE;
        }
    }

    // Now pick up the edges if we set the border-style for that edge to "none"
    pborderinfo->wEdges &= ~BF_RECT;

    if(pborderinfo->aiWidths[BORDER_TOP])
    {
        pborderinfo->wEdges |= BF_TOP;
    }
    if(pborderinfo->aiWidths[BORDER_RIGHT])
    {
        pborderinfo->wEdges |= BF_RIGHT;
    }
    if(pborderinfo->aiWidths[BORDER_BOTTOM])
    {
        pborderinfo->wEdges |= BF_BOTTOM;
    }
    if(pborderinfo->aiWidths[BORDER_LEFT])
    {
        pborderinfo->wEdges |= BF_LEFT;
    }

    if(fAll)
    {
        GetBorderColorInfoHelper(pNodeContext,pdci,pborderinfo);
    }

    if(pborderinfo->wEdges
        || pborderinfo->rcSpace.top    > 0
        || pborderinfo->rcSpace.bottom > 0
        || pborderinfo->rcSpace.left   > 0
        || pborderinfo->rcSpace.right  > 0)
    {
        return (pborderinfo->wEdges&BF_RECT
            && pborderinfo->aiWidths[BORDER_TOP]==pborderinfo->aiWidths[BORDER_BOTTOM]
        && pborderinfo->aiWidths[BORDER_LEFT]==pborderinfo->aiWidths[BORDER_RIGHT]
        && pborderinfo->aiWidths[BORDER_TOP]==pborderinfo->aiWidths[BORDER_LEFT]
        ? DISPNODEBORDER_SIMPLE : DISPNODEBORDER_COMPLEX);
    }
    return DISPNODEBORDER_NONE;
}