
#include "stdafx.h"
#include "BodyLayout.h"

const CLayout::LAYOUTDESC CBodyLayout::s_layoutdesc =
{
    LAYOUTDESC_FLOWLAYOUT, // _dwFlags
};

//+---------------------------------------------------------------------------
//
//  Member:     CBodyLayout::HandleMessage
//
//  Synopsis:   Check if we have a setcursor or mouse down in the
//              border (outside of the client rect) so that we can
//              pass the message to the containing frameset if we're
//              hosted in one.
//
//----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CBodyLayout::HandleMessage(CMessage* pMessage)
{
    HRESULT hr = S_FALSE;

    // have any mouse messages
    if(pMessage->message>=WM_MOUSEFIRST && pMessage->message<=WM_MOUSELAST && pMessage->message!=WM_MOUSEMOVE && ElementOwner()==Doc()->_pElemCurrent)
    {
        RequestFocusRect(FALSE);
    }

    if((pMessage->message>=WM_MOUSEFIRST && pMessage->message!=WM_MOUSEWHEEL && pMessage->message<=WM_MOUSELAST) || pMessage->message==WM_SETCURSOR)
    {
        RECT rc;
        GetRect(&rc, COORDSYS_GLOBAL);

        if(pMessage->htc!=HTC_HSCROLLBAR && pMessage->htc!=HTC_VSCROLLBAR
            && !PtInRect(&rc, pMessage->pt) && Doc()->GetCaptureObject()!=ElementOwner())
        {
            hr = S_OK;
            goto Cleanup;
        }
    }

    hr = super::HandleMessage(pMessage);

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     CBodyLayout::UpdateScrollbarInfo, protected
//
//  Synopsis:   Update CDispNodeInfo to reflect the correct scroll-related settings
//
//-------------------------------------------------------------------------
void CBodyLayout::UpdateScrollInfo(CDispNodeInfo* pdni) const
{
    CDocument* pDoc = Doc();

    if(pdni->_overflowX==styleOverflowNotSet || pdni->_overflowY==styleOverflowNotSet)
    {
        DWORD dwFrameOptions = pDoc->_dwFrameOptions &
            (FRAMEOPTIONS_SCROLL_NO|FRAMEOPTIONS_SCROLL_YES|FRAMEOPTIONS_SCROLL_AUTO);

        if(pDoc->_dwFlagsHostInfo & DOCHOSTUIFLAG_SCROLL_NO)
        {
            dwFrameOptions = FRAMEOPTIONS_SCROLL_NO;
        }
        else
        {
            switch(((CBodyLayout*)this)->Body()->GetAAscroll())
            {
            case bodyScrollno:
                dwFrameOptions = FRAMEOPTIONS_SCROLL_NO;
                break;

            case bodyScrollyes:
                dwFrameOptions = FRAMEOPTIONS_SCROLL_YES;
                break;

            case bodyScrollauto:
                dwFrameOptions = FRAMEOPTIONS_SCROLL_AUTO;
                break;

            case bodyScrolldefault:
                if(!dwFrameOptions)
                {
                    dwFrameOptions = FRAMEOPTIONS_SCROLL_YES;
                }
                break;
            }
        }

        switch(dwFrameOptions)
        {
        case FRAMEOPTIONS_SCROLL_NO:
            if(pdni->_overflowX == styleOverflowNotSet)
            {
                pdni->_overflowX = styleOverflowHidden;
            }
            if(pdni->_overflowY == styleOverflowNotSet)
            {
                pdni->_overflowY = styleOverflowHidden;
            }
            break;

        case FRAMEOPTIONS_SCROLL_AUTO:
            if(pdni->_overflowX == styleOverflowNotSet)
            {
                pdni->_overflowX = styleOverflowAuto;
            }
            if(pdni->_overflowY == styleOverflowNotSet)
            {
                pdni->_overflowY = styleOverflowAuto;
            }
            break;

        case FRAMEOPTIONS_SCROLL_YES:
        default:
            pdni->_sp._fHSBAllowed = TRUE;
            pdni->_sp._fHSBForced  = FALSE;
            pdni->_sp._fVSBAllowed = pdni->_sp._fVSBForced  = TRUE;
            break;
        }
    }

    // If an overflow value was set or generated, set the scrollbar properties using it
    if(pdni->_overflowX!=styleOverflowNotSet || pdni->_overflowY!=styleOverflowNotSet)
    {
        GetDispNodeScrollbarProperties(pdni);
    }
}

// JS OnSelection event
HRESULT CBodyLayout::OnSelectionChange(void)
{
    DYNCAST(CBodyElement, ElementOwner())->Fire_onselect(); // JS event
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CalcSize
//
//  Synopsis:   see CSite::CalcSize documentation
//
//----------------------------------------------------------------------------
DWORD CBodyLayout::CalcSize(CCalcInfo* pci, SIZE* psize, SIZE* psizeDefault)
{
    CScopeFlag csfCalcing(this);
    DWORD dwRet = super::CalcSize(pci, psize, psizeDefault);

    // Update the focus rect
    if(_fFocusRect && ElementOwner()==Doc()->_pElemCurrent)
    {
        RedrawFocusRect();
    }

    return dwRet;
}

//+---------------------------------------------------------------------------
//
//  Member:     CBodyLayout::NotifyScrollEvent
//
//  Synopsis:   Respond to a change in the scroll position of the display node
//
//----------------------------------------------------------------------------
void CBodyLayout::NotifyScrollEvent(RECT* prcScroll, SIZE* psizeScrollDelta)
{
    // Update the focus rect
    if(_fFocusRect && ElementOwner()==Doc()->_pElemCurrent)
    {
        RedrawFocusRect();
    }

    super::NotifyScrollEvent(prcScroll, psizeScrollDelta);
}

//+---------------------------------------------------------------------------
//
//  Member:     RequestFocusRect
//
//  Synopsis:   Turns on/off the focus rect of the body.
//
//  Arguments:  fOn     flag for requested state
//
//----------------------------------------------------------------------------
void CBodyLayout::RequestFocusRect(BOOL fOn)
{
    if(!_fFocusRect != !fOn)
    {
        _fFocusRect = fOn;
        RedrawFocusRect();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     RedrawFocusRect
//
//  Synopsis:   Redraw the focus rect of the body.
//
//----------------------------------------------------------------------------
void CBodyLayout::RedrawFocusRect()
{
    Assert(ElementOwner() == Doc()->_pElemCurrent);
    CView* pView = GetView();

    // Force update of focus shape
    pView->SetFocus(NULL, 0);
    pView->SetFocus(ElementOwner(), 0);
    pView->InvalidateFocus();
}

//+---------------------------------------------------------------------------
//
//  Member:     CBodyLayout::GetFocusShape
//
//  Synopsis:   Returns the shape of the focus outline that needs to be drawn
//              when this element has focus. This function creates a new
//              CShape-derived object. It is the caller's responsibility to
//              release it.
//
//----------------------------------------------------------------------------
HRESULT CBodyLayout::GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape)
{
    CRect       rc;
    CRectShape* pShape;
    HRESULT     hr = S_FALSE;

    Assert(ppShape);
    *ppShape = NULL;

    if(!_fFocusRect)
    {
        goto Cleanup;
    }

    GetClientRect(&rc, CLIENTRECT_BACKGROUND);
    if(rc.IsEmpty())
    {
        goto Cleanup;
    }

    pShape = new CRectShape;
    if(!pShape)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pShape->_rect = rc;
    pShape->_cThick = 2; // always draw extra thick for BODY
    *ppShape = pShape;

    hr = S_OK;

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+------------------------------------------------------------------------
//
//  Member:     GetBackgroundInfo
//
//  Synopsis:   Fills out a background info for which has details on how
//              to display a background color &| background image.
//
//-------------------------------------------------------------------------
BOOL CBodyLayout::GetBackgroundInfo(
                               CFormDrawInfo*   pDI,
                               BACKGROUNDINFO*  pbginfo,
                               BOOL             fAll,
                               BOOL             fRightToLeft)
{
    Assert(pDI || !fAll);

    super::GetBackgroundInfo(pDI, pbginfo, fAll, fRightToLeft);

    if(pbginfo->crBack == CLR_INVALID)
    {
        pbginfo->crBack = Doc()->_pOptionSettings->crBack();
        Assert(pbginfo->crBack != CLR_INVALID);
    }

    return TRUE;
}
