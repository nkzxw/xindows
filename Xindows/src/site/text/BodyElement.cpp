
#include "stdafx.h"
#include "BodyElement.h"

#define _cxx_
#include "../gen/body.hdl"

HRESULT InitTextSubSystem();

// implementation of CBodyElement::CreateLayout()
IMPLEMENT_LAYOUT_FNS(CBodyElement, CBodyLayout)

CElement::ACCELS CBodyElement::s_AccelsBodyDesign = CElement::ACCELS (&CElement::s_AccelsElementDesign, 0);
CElement::ACCELS CBodyElement::s_AccelsBodyRun    = CElement::ACCELS (&CElement::s_AccelsElementRun,    0);

const CElement::CLASSDESC CBodyElement::s_classdesc =
{
    {
        &CLSID_HTMLBody,                // _pclsid
        s_acpi,                         // _pcpi
        ELEMENTDESC_TEXTSITE  |
        ELEMENTDESC_BODY      |
        ELEMENTDESC_CANSCROLL |
        ELEMENTDESC_NOTIFYENDPARSE,     // _dwFlags
        &IID_IHTMLBodyElement,          // _piidDispinterface
        &s_apHdlDescs,                  // _apHdlDesc
    },
    &CBodyElement::s_AccelsBodyDesign,  // _pAccelsDesign
    &CBodyElement::s_AccelsBodyRun      // _pAccelsRun
};

const long s_adispCommonProps[CBodyElement::NUM_COMMON_PROPS][2] =
{
    { DISPID_CDoc_bgColor,    DISPID_CBodyElement_bgColor},
    { DISPID_CDoc_linkColor,  DISPID_CBodyElement_link},
    { DISPID_CDoc_alinkColor, DISPID_CBodyElement_aLink},
    { DISPID_CDoc_vlinkColor, DISPID_CBodyElement_vLink},
    { DISPID_CDoc_fgColor,    DISPID_CBodyElement_text}
};

//+------------------------------------------------------------------------
//
//  Member:     CBodyElement::PrivateQueryInterface, IUnknown
//
//  Synopsis:   Private unknown QI.
//
//-------------------------------------------------------------------------
HRESULT CBodyElement::PrivateQueryInterface(REFIID iid, void** ppv)
{
    RRETURN(super::PrivateQueryInterface(iid, ppv));
}

HRESULT CBodyElement::CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElement)
{
    HRESULT hr;

    Assert(pht->IsBegin(ETAG_BODY));
    Assert(ppElement);

    hr = InitTextSubSystem();
    if(hr)
    {
        goto Cleanup;
    }

    *ppElement = new CBodyElement(pDoc);

    hr = (*ppElement) ? S_OK : E_OUTOFMEMORY;

Cleanup:

    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CBodyElement::CBodyElement
//
//---------------------------------------------------------------
CBodyElement::CBodyElement(CDocument* pDoc) : CTxtSite(ETAG_BODY, pDoc) 
{
    _fSynthetic     = FALSE;
    _fInheritFF     = TRUE;
}

//+---------------------------------------------------------------
//
//  Member:     CBodyElement::Init2
//
//  Synopsis:   Final initializer
//
//---------------------------------------------------------------
HRESULT CBodyElement::Init2(CInit2Context* pContext)
{
    HRESULT         hr;
    int             i;
    CDocument*      pDoc = Doc();

    // before we do anything copy potentially initialized values
    // from the document's aa
    if(pDoc)
    {
        CAttrArray* pAA = *(pDoc->GetAttrArray());
        if(pAA)
        {
            CAttrValue* pAV = NULL;

            for(i=0; i<NUM_COMMON_PROPS; i++)
            {
                pAV = pAA->Find(s_adispCommonProps[i][0]);
                if(pAV)
                {
                    // Implicit assumption that we're always dealing with I4's
                    hr = AddSimple(s_adispCommonProps[i][1], pAV->GetLong(), pAV->GetAAType());
                }
            }
        }
    }

    hr = super::Init2(pContext);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     Notify
//
//-------------------------------------------------------------------------
void CBodyElement::Notify(CNotification* pNF)
{
    CDocument* pDoc = Doc();

    super::Notify(pNF);

    switch(pNF->Type())
    {
    case NTYPE_ELEMENT_QUERYTABBABLE:
        {
            CQueryFocus* pQueryFocus = (CQueryFocus*)pNF->DataAsPtr();

            pQueryFocus->_fRetVal = FALSE;
        }
        break;

    case NTYPE_ELEMENT_SETFOCUS:
        {
            CSetFocus*  pSetFocus   = (CSetFocus*)pNF->DataAsPtr();
            CMessage*   pMessage    = pSetFocus->_pMessage;

            // We want to turn on the focus rect only during certain
            // conditions:
            // 1) we are the current site
            // 2) we got here due to a tab/frametab key
            // 3) There is no selection or 0 len selection ( a Caret ! )
            Layout()->RequestFocusRect(
                pDoc->_pElemCurrent==this
                && pMessage && pMessage->message==WM_KEYDOWN
                && (pDoc->HasSelection() ? ( pDoc->GetSelectionType()==SELECTION_TYPE_Caret) : TRUE));
        }
        break;

    case NTYPE_ELEMENT_ENTERTREE:
        {
            // Alert the view that the top client element may have changed
            pDoc->OpenView();

            Layout()->_fContentsAffectSize = FALSE;

            // BUGBUG
            //
            // Major hack code to simulate the setting of the top site
            CMarkup*    pMarkup = GetMarkup();
            CElement*   pElemClient = pMarkup->GetElementClient();
            CSite*      pSite = this;

            // This could be NULL during DOM operations!
            if(pElemClient)
            {
                pElemClient->ResizeElement(NFLAGS_SYNCHRONOUSONLY);
            }

            if(pMarkup->IsPrimaryMarkup())
            {
                // Let the topsite become current/uiactive if our current
                // state allows that.
                pSite->GetCurLayout()->_fSizeThis = TRUE;

                // Notify the view of a new possible top-most element
                SendNotification(NTYPE_VIEW_ATTACHELEMENT);

                if(!pDoc->_fCurrencySet)
                {
                    // EnterTree is not a good place to set currency, especially in
                    // design mode, because it could force a synchronous recalc by
                    // trying to set the caret. So, we delay setting the currency.
                    GWPostMethodCall(pDoc, ONCALL_METHOD(CDocument, DeferSetCurrency, DeferSetCurrency),
                        0, FALSE, "CDoc::DeferSetCurrency");
                }
            }
            // End BUGBUG
        }
        break;

    case NTYPE_ELEMENT_EXITTREE_1:
        // Notify the view that the top element may have left
        SendNotification(NTYPE_VIEW_DETACHELEMENT);
        break;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CBodyElement::GetBorderInfo, public
//
//  Synopsis:   Returns information about what borders we have.
//
//----------------------------------------------------------------------------
#define WIDTH_3DBORDER  2
inline void Set3DBorderEdgeInfo(BOOL fNeedBorderEdge, int cEdge, CBorderInfo* pborderinfo)
{
    if(fNeedBorderEdge)
    {
        pborderinfo->abStyles[cEdge] = fmBorderStyleSunken;
        pborderinfo->aiWidths[cEdge] = WIDTH_3DBORDER;
    }
}

DWORD CBodyElement::GetBorderInfo(CDocInfo* pdci, CBorderInfo* pborderinfo, BOOL fAll)
{
    CDocument* pDoc = Doc();

    if((pDoc->_dwFrameOptions&FRAMEOPTIONS_NO3DBORDER)==0 &&
        (pDoc->_dwFlagsHostInfo&DOCHOSTUIFLAG_NO3DBORDER)==0)
    {
        pDoc->_b3DBorder = NEED3DBORDER_TOP|NEED3DBORDER_LEFT
            |NEED3DBORDER_BOTTOM|NEED3DBORDER_RIGHT;

        Set3DBorderEdgeInfo(
            (pDoc->_b3DBorder&NEED3DBORDER_TOP)!=0,
            BORDER_TOP,
            pborderinfo);
        Set3DBorderEdgeInfo(
            (pDoc->_b3DBorder&NEED3DBORDER_LEFT)!=0,
            BORDER_LEFT,
            pborderinfo);
        Set3DBorderEdgeInfo(
            (pDoc->_b3DBorder&NEED3DBORDER_BOTTOM)!=0,
            BORDER_BOTTOM,
            pborderinfo);
        Set3DBorderEdgeInfo(
            (pDoc->_b3DBorder&NEED3DBORDER_RIGHT)!=0,
            BORDER_RIGHT,
            pborderinfo);

        pborderinfo->wEdges = BF_RECT;
    }

    CElement::GetBorderInfo(pdci, pborderinfo, fAll);

    if(pborderinfo->wEdges
        || pborderinfo->rcSpace.top    > 0
        || pborderinfo->rcSpace.bottom > 0
        || pborderinfo->rcSpace.left   > 0
        || pborderinfo->rcSpace.right  > 0)
    {
        return (pborderinfo->wEdges&(BF_TOP|BF_RIGHT|BF_BOTTOM|BF_LEFT)
            && pborderinfo->aiWidths[BORDER_TOP]==pborderinfo->aiWidths[BORDER_BOTTOM]
        && pborderinfo->aiWidths[BORDER_LEFT]==pborderinfo->aiWidths[BORDER_RIGHT]
        && pborderinfo->aiWidths[BORDER_TOP]==pborderinfo->aiWidths[BORDER_LEFT]
        ? DISPNODEBORDER_SIMPLE : DISPNODEBORDER_COMPLEX);
    }
    return DISPNODEBORDER_NONE;
}

//+------------------------------------------------------------------------
//
//  Member:     CBodyElement::ApplyDefaultFormat
//
//  Synopsis:   Applies default formatting properties for that element to
//              the char and para formats passed in
//
//  Arguments:  pCFI - Format Info needed for cascading
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CBodyElement::ApplyDefaultFormat(CFormatInfo* pCFI)
{
    HRESULT hr          = S_OK;
    DWORD   dwWidth     = 0xffffffff;
    DWORD   dwHeight    = 0xffffffff;
    const CFancyFormat* pFFParent = NULL;
    CTreeNode*  pNodeContext = pCFI->_pNodeContext;
    CTreeNode*  pNode;
    CDocument*  pDoc = Doc();

    Assert(pNodeContext && SameScope(this, pNodeContext));

    pCFI->PrepareFancyFormat();

    // Climb up the tree to find a background color, inherit an image url
    // from our parent.
    for(pNode=pNodeContext->Parent(); pNode; pNode=pNode->Parent())
    {
        pFFParent = pNode->GetFancyFormat();

        if(!pCFI->_ff()._lImgCtxCookie)
        {
            pCFI->_ff()._lImgCtxCookie = pFFParent->_lImgCtxCookie;
        }

        if(pFFParent->_ccvBackColor.IsDefined())
        {
            break;
        }
    }

    pCFI->_ff()._ccvBackColor = pFFParent->_ccvBackColor;
    Assert(pCFI->_ff()._ccvBackColor.IsDefined());

    pCFI->_ff()._ccvBorderColorLight.SetSysColor(COLOR_3DLIGHT);
    pCFI->_ff()._ccvBorderColorShadow.SetSysColor(COLOR_BTNSHADOW);
    pCFI->_ff()._ccvBorderColorHilight.SetSysColor(COLOR_BTNHIGHLIGHT);
    pCFI->_ff()._ccvBorderColorDark.SetSysColor(COLOR_3DDKSHADOW);

    pCFI->UnprepareForDebug();

    if(GetAAscroll() == bodyScrollno)
    {
        pCFI->PrepareFancyFormat();
        pCFI->_ff()._fNoScroll = TRUE;
        pCFI->UnprepareForDebug();
    }

    hr = super::ApplyDefaultFormat(pCFI);
    if(hr)
    {
        goto Cleanup;
    }

    {
        pCFI->PrepareFancyFormat();

        // BUGBUG (srinib) - for frames, if right/bottom margin is not specified and left/top
        // are specified through attributes then use left/top as default values.

        // If an explicit top margin is not specified, set it to the default value
        if(!pCFI->_pff->_fExplicitTopMargin)
        {
            if(dwHeight == 0xffffffff)
            {
                pCFI->_ff()._cuvMarginTop.SetRawValue(s_propdescCBodyElementtopMargin.a.ulTagNotPresentDefault);
            }
            else
            {
                pCFI->_ff()._cuvMarginTop.SetValue(long(dwHeight), CUnitValue::UNIT_PIXELS);
            }
        }

        // if an explicit bottom margin is not specified, set default bottom margin.
        if(!pCFI->_pff->_fExplicitBottomMargin)
        {
            if(dwHeight == 0xffffffff)
            {
                pCFI->_ff()._cuvMarginBottom.SetRawValue(s_propdescCBodyElementbottomMargin.a.ulTagNotPresentDefault);
            }
            else
            {
                pCFI->_ff()._cuvMarginBottom.SetValue(long(dwHeight), CUnitValue::UNIT_PIXELS);
            }
        }

        if(pCFI->_pff->_cuvMarginLeft.IsNull())
        {
            if(dwWidth ==  0xffffffff) // margin specified on the frame
            {
                pCFI->_ff()._cuvMarginLeft.SetRawValue(s_propdescCBodyElementleftMargin.a.ulTagNotPresentDefault);
            }
            else
            {
                pCFI->_ff()._cuvMarginLeft.SetValue(long(dwWidth), CUnitValue::UNIT_PIXELS);
            }
        }

        if(pCFI->_pff->_cuvMarginRight.IsNull())
        {
            if(dwWidth ==  0xffffffff)  // margin specified on the frame
            {
                pCFI->_ff()._cuvMarginRight.SetRawValue(s_propdescCBodyElementrightMargin.a.ulTagNotPresentDefault);
            }
            else
            {
                pCFI->_ff()._cuvMarginRight.SetValue(long(dwWidth), CUnitValue::UNIT_PIXELS);
            }
        }

        // cache percent attribute information
        if(pCFI->_pff->_cuvMarginTop.IsPercent()
            || pCFI->_pff->_cuvMarginBottom.IsPercent())
        {
            pCFI->_ff()._fPercentVertPadding = TRUE;
        }

        if(pCFI->_pff->_cuvMarginLeft.IsPercent()
            || pCFI->_pff->_cuvMarginRight.IsPercent())
        {
            pCFI->_ff()._fPercentHorzPadding = TRUE;
        }

        pCFI->_ff()._fHasMargins = TRUE;

        pCFI->UnprepareForDebug();
    }

    pCFI->PrepareFancyFormat();
    pCFI->_ff()._fHeightPercent = TRUE;
    pCFI->_ff()._fWidthPercent = TRUE;
    pCFI->UnprepareForDebug();

Cleanup:
    RRETURN (hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CBodyElement::OnPropertyChange
//
//  Synopsis:   Handles change of property on body tag
//
//  Arguments:  dispid:  id of the property changing
//              dwFlags: change flags
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CBodyElement::OnPropertyChange(DISPID dispid, DWORD dwFlags)
{
    HRESULT hr;

    hr = super::OnPropertyChange(dispid, dwFlags);
    if(hr != S_OK)
    {
        goto Cleanup;
    }

    switch(dispid)
    {
    case DISPID_BACKCOLOR:
        {
            CDocument* pDoc = Doc();
            //redraw the window
            if(pDoc && pDoc->_pInPlace && pDoc->_pInPlace->_hwnd)
            {
                pDoc->GetView()->Invalidate();
            }
        }
        break;
    }

Cleanup:
    return hr;
}

//+---------------------------------------------------------------------------
//
//  Member:     CBodyElement::GetFocusShape
//
//  Synopsis:   Returns the shape of the focus outline that needs to be drawn
//              when this element has focus. This function creates a new
//              CShape-derived object. It is the caller's responsibility to
//              release it.
//
//----------------------------------------------------------------------------
HRESULT CBodyElement::GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape)
{
    RRETURN1(Layout()->GetFocusShape(lSubDivision, pdci, ppShape), S_FALSE);
}

//+----------------------------------------------------------------------------
//
// Member: WaitForRecalc
//
//-----------------------------------------------------------------------------
void CBodyElement::WaitForRecalc()
{
    CFlowLayout* pFlowLayout = Layout();

    if(pFlowLayout)
    {
        CDisplay* pdp = pFlowLayout->GetDisplay();

        if(pdp)
        {
            pdp->WaitForRecalc(GetLastCp(), -1);
        }
    }
}

//+----------------------------------------------------------------------------
// Text subsystem initialization
//-----------------------------------------------------------------------------
void RegisterFETCs();

// System static variables
extern void DeInitFontCache();

HRESULT InitTextSubSystem()
{
    static BOOL fTSInitted = FALSE;

    if(!fTSInitted)
    {
        RegisterFETCs(); // Register new clipboard formats

        // BUGBUG - EricVas: This stuff is not deinitialized anywhere

        fTSInitted = TRUE;
    }

    return S_OK;
}

void DeinitTextSubSystem()
{
    DeInitFontCache();
}
