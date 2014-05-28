
#include "stdafx.h"
#include "TextArea.h"

#define _cxx_
#include "../gen/textarea.hdl"

IMPLEMENT_LAYOUT_FNS(CTextArea, CTextAreaLayout)
IMPLEMENT_LAYOUT_FNS(CRichtext, CRichtextLayout)

CElement::ACCELS CRichtext::s_AccelsTextareaDesign = CElement::ACCELS(NULL, 106/*IDR_ACCELS_INPUTTXT_DESIGN wlw note*/);
CElement::ACCELS CRichtext::s_AccelsTextareaRun    = CElement::ACCELS(NULL, 105/*IDR_ACCELS_INPUTTXT_RUN wlw note*/);

//+------------------------------------------------------------------------
//
//  Member:     CElement::s_classdesc
//
//  Synopsis:   class descriptor
//
//-------------------------------------------------------------------------
const CElement::CLASSDESC CTextArea::s_classdesc =
{
    {
        &CLSID_HTMLTextAreaElement,     // _pclsid
        s_acpi,                         // _pcpi
        ELEMENTDESC_TEXTSITE |
        ELEMENTDESC_DONTINHERITSTYLE |
        ELEMENTDESC_SHOWTWS |
        ELEMENTDESC_CANSCROLL |
        ELEMENTDESC_VPADDING |
        ELEMENTDESC_NOTIFYENDPARSE |
        ELEMENTDESC_NOANCESTORCLICK |
        ELEMENTDESC_HASDEFDESCENT,
        &IID_IHTMLTextAreaElement,      // _piidDispinterface
        &s_apHdlDescs,                  // _apHdlDesc
    },
    (void*)s_apfnpdIHTMLTextAreaElement,// _pfnTearOff
    &s_AccelsTextareaDesign,            // _pAccelsDesign
    &s_AccelsTextareaRun                // _pAccelsRun
};

//+------------------------------------------------------------------------
//
//  Member:     CTextArea::CreateElement()
//
//  Synopsis:   called by the parser to create an instance
//
//-------------------------------------------------------------------------
HRESULT CTextArea::CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElement)
{
    Assert(ppElement);
    *ppElement = new CTextArea(pht->GetTag(), pDoc);
    RRETURN((*ppElement)?S_OK:E_OUTOFMEMORY);
}

//+------------------------------------------------------------------------
//
//  Member:     CTextArea::Save
//
//  Synopsis:   Save the element to the stream
//
//-------------------------------------------------------------------------
HRESULT CTextArea::Save(CStreamWriteBuff* pStmWrBuff, BOOL fEnd)
{
    HRESULT hr = S_OK;

    Assert(!_fInSave);
    _fInSave = TRUE;

    if(!fEnd)
    {
        pStmWrBuff->BeginPre();
    }

    hr = super::Save(pStmWrBuff, fEnd);

    if(hr)
    {
        goto Cleanup;
    }

    if(fEnd)
    {
        pStmWrBuff->EndPre();
    }

    Assert(_fInSave); // this will catch recursion
    _fInSave = FALSE;

Cleanup:
    RRETURN(hr);
}

void CTextArea::GetPlainTextWithBreaks(TCHAR* pchBuff)
{
    CTxtPtr     tp(GetMarkup(), Layout()->GetContentFirstCp());
    long        i;
    CDisplay*   pdp = Layout()->GetDisplay();
    long        c = pdp->Count();
    CLine*      pLine;

    Assert(Layout()->GetMultiLine());
    Assert(pchBuff);

    for(i=0; i<c; i++)
    {
        pLine = pdp->Elem(i);

        if(pLine->_cch)
        {
            tp.GetRawText(pLine->_cch, pchBuff);
            tp.AdvanceCp(pLine->_cch);
            pchBuff += pLine->_cch;

            // If the line ends in a '\r', we need to append a '\r'.
            // Otherwise, it must be having a soft break (unless it is
            // the last line), so we need to append a '\r\n'
            if(pLine->_fHasBreak)
            {
                Assert(*(pchBuff-1) == _T('\r'));
                *pchBuff++ = _T('\n');
            }
            else if(i < c-1)
            {
                *pchBuff++ = _T('\r');
                *pchBuff++ = _T('\n');
            }
        }
    }
    // Null-terminate
    *pchBuff = 0;
}

long CTextArea::GetPlainTextLengthWithBreaks()
{
    long        len = 1; // for trailing '\0'
    long        i;
    CDisplay*   pdp = Layout()->GetDisplay();
    long        c = pdp->Count();
    CLine*      pLine;

    Assert(HasLayout());
    Assert(Layout()->GetMultiLine());
    for(i=0; i<c; i++)
    {
        // Every line except the last must be non-empty.
        Assert(i==c-1 || pdp->Elem(i)->_cch>0);

        pLine = pdp->Elem(i);
        Assert(pLine);
        len += pLine->_cch;

        // If the line ends in a '\r', we need to append a '\r'.
        // Otherwise, it must be having a soft break (unless it is
        // the last line), so we need to append a '\r\n'
        if(pLine->_fHasBreak)
        {
            len++;
        }
        else if(i < c-1)
        {
            len += 2;
        }
    }

    return len;
}

//+------------------------------------------------------------------------
//
//  Member:     CTextArea::GetSubmitValue()
//
//  Synopsis:   returns the inner HTML
//
//-------------------------------------------------------------------------
HRESULT CTextArea::GetSubmitValue(CString* pstr)
{
    HRESULT hr = S_OK;

    // Make sure the site is not detached. Otherwise, we
    // cannot get to the runs
    if(!IsInMarkup())
    {
        goto Cleanup;
    }

    if(GetAAwrap() == htmlWrapHard)
    {
        long len = GetPlainTextLengthWithBreaks();
        if(len > 0)
        {
            pstr->SetLengthNoAlloc(0);
            hr = pstr->ReAlloc(len);
            if(hr)
            {
                goto Cleanup;
            }
            GetPlainTextWithBreaks(*pstr);
        }
    }
    else
    {
        hr = GetValueHelper(pstr);
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CTextArea::ApplyDefaultFormat()
//
//  Synopsis:
//
//-------------------------------------------------------------------------
HRESULT CTextArea::ApplyDefaultFormat(CFormatInfo* pCFI)
{
    CDocument*          pDoc     = Doc();
    CODEPAGESETTINGS*   pCS      = pDoc->_pCodepageSettings;
    COLORREF            crWindow = GetSysColorQuick(COLOR_WINDOW);
    HRESULT             hr       = S_OK;

    if(!pCFI->_pff->_ccvBackColor.IsDefined()
        || pCFI->_pff->_ccvBackColor.GetColorRef()!=crWindow)
    {
        pCFI->PrepareFancyFormat();
        pCFI->_ff()._ccvBackColor = crWindow;
    }

    pCFI->PrepareCharFormat();
    pCFI->PrepareParaFormat();

    pCFI->_cf()._ccvTextColor.SetSysColor(COLOR_WINDOWTEXT);

    // our intrinsics shouldn't inherit the cursor property. they have a 'default'
    pCFI->_cf()._bCursorIdx = styleCursorAuto;

    pCFI->_cf()._fBold = FALSE;
    pCFI->_cf()._wWeight = FW_NORMAL; //FW_NORMAL = 400
    pCFI->_cf()._yHeight = 200;       // 10 * 20, 10 points

    // Thai does not have a fixed pitch font. Leave it as proportional
    if(pDoc->GetCodePage() != CP_THAI)
    {
        pCFI->_cf()._bPitchAndFamily = FIXED_PITCH;
        pCFI->_cf()._latmFaceName = pCS->latmFixedFontFace;
    }
    pCFI->_cf()._bCharSet = pCS->bCharSet;
    pCFI->_cf()._fNarrow = IsNarrowCharSet(pCFI->_cf()._bCharSet);
    pCFI->_pf()._cuvTextIndent.SetPoints(0);
    pCFI->_fPre = TRUE;

    pCFI->UnprepareForDebug();

    hr = super::ApplyDefaultFormat(pCFI);
    if(hr)
    {
        goto Cleanup;
    }

    pCFI->PrepareCharFormat();

    // font height in CharFormat is already nonscaling size in twips
    pCFI->_cf().SetHeightInNonscalingTwips(pCFI->_cf()._yHeight);

    pCFI->UnprepareForDebug();

Cleanup:
    RRETURN(hr);
}

// Richtext member functions
const CElement::CLASSDESC CRichtext::s_classdesc =
{
    {
        &CLSID_HTMLRichtextElement,      // _pclsid
        s_acpi,                          // _pcpi
        ELEMENTDESC_TEXTSITE          |
        ELEMENTDESC_SHOWTWS           |
        ELEMENTDESC_ANCHOROUT         |
        ELEMENTDESC_NOBKGRDRECALC     |
        ELEMENTDESC_CANSCROLL         |
        ELEMENTDESC_HASDEFDESCENT     |
        ELEMENTDESC_NOTIFYENDPARSE,      // _dwFlags
        &IID_IHTMLTextAreaElement,       // _piidDispinterface
        &s_apHdlDescs,                   // _apHdlDesc
    },
    (void*)s_apfnpdIHTMLTextAreaElement, // _pfnTearOff
    &s_AccelsTextareaDesign,             // _pAccelsDesign
    &s_AccelsTextareaRun                 // _pAccelsRun
};

//+------------------------------------------------------------------------
//
//  Member:     CRichtext::CreateElement()
//
//  Synopsis:   called by the parser to create instance
//
//-------------------------------------------------------------------------
HRESULT CRichtext::CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElement)
{
    Assert(ppElement);
    *ppElement = new CRichtext(pht->GetTag(), pDoc);
    RRETURN((*ppElement)?S_OK:E_OUTOFMEMORY);
}

//+------------------------------------------------------------------------
//
//  Member:     CRichtext::Init2()
//
//  Synopsis:   called by the parser to initialize instance
//
//-------------------------------------------------------------------------
HRESULT CRichtext::Init2(CInit2Context* pContext)
{
    HRESULT hr = S_OK;

    hr = super::Init2(pContext);
    if(!OK(hr))
    {
        goto Cleanup;
    }

    {
        Assert(Tag() == ETAG_TEXTAREA);
        SetAAtype(htmlInputTextarea);
    }

    if(!GetAAreadOnly())
    {
        _fEditAtBrowse = TRUE;
    }

Cleanup:
    RRETURN1(hr, S_INCOMPLETE);
}

//+------------------------------------------------------------------------
//
//  Member:     OnPropertyChange
//
//  Note:       Called after a property has changed to notify derived classes
//              of the change.  All properties (except those managed by the
//              derived class) are consistent before this call.
//
//              Also, fires a property change notification to the site.
//
//-------------------------------------------------------------------------
HRESULT CRichtext::OnPropertyChange(DISPID dispid, DWORD dwFlags)
{
    HRESULT hr = S_OK;

    hr = super::OnPropertyChange(dispid, dwFlags);
    if(hr)
    {
        goto Cleanup;
    }

    switch(dispid)
    {
    case DISPID_CRichtext_wrap:
        Layout()->SetWrap();
        break;

    case DISPID_CRichtext_readOnly:
        _fEditAtBrowse = !GetAAreadOnly();
        // update the editability in the edit context
        if(HasCurrency())
        {
            Doc()->EnsureEditContext(this);
        }
        break;
    }

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CRichtext::QueryStatus
//
//  Synopsis:   Called to discover if a given command is supported
//              and if it is, what's its state.  (disabled, up or down)
//
//--------------------------------------------------------------------------
HRESULT CRichtext::QueryStatus(
       GUID*        pguidCmdGroup,
       ULONG        cCmds,
       MSOCMD       rgCmds[],
       MSOCMDTEXT*  pcmdtext)
{
    Assert(CBase::IsCmdGroupSupported(pguidCmdGroup));
    Assert(cCmds == 1);

    MSOCMD* pCmd = &rgCmds[0];
    ULONG cmdID;

    Assert(!pCmd->cmdf);

    cmdID = CBase::IDMFromCmdID(pguidCmdGroup, pCmd->cmdID);
    switch(cmdID)
    {
    case IDM_INSERTOBJECT:
        // Don't allow objects to be inserted in TEXTAREA
        pCmd->cmdf = MSOCMDSTATE_DISABLED;
        return S_OK;

    default:
        RRETURN_NOTRACE(super::QueryStatus(
            pguidCmdGroup,
            1,
            pCmd,
            pcmdtext));
    }
}

HRESULT CRichtext::Exec(
        GUID*       pguidCmdGroup,
        DWORD       nCmdID,
        DWORD       nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut)
{
    Assert(CBase::IsCmdGroupSupported(pguidCmdGroup));

    int     idm = CBase::IDMFromCmdID(pguidCmdGroup, nCmdID);
    HRESULT hr  = MSOCMDERR_E_NOTSUPPORTED;

    switch(idm)
    {
    case IDM_SELECTALL:
        select();
        hr = S_OK;
        break;
    }

    if(hr == MSOCMDERR_E_NOTSUPPORTED)
    {
        hr = super::Exec(
            pguidCmdGroup,
            nCmdID,
            nCmdexecopt,
            pvarargIn,
            pvarargOut);
    }

    RRETURN_NOTRACE(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CRichtext::Notify()
//
//  Synopsis:   handle notifications
//
//-------------------------------------------------------------------------
void CRichtext::Notify(CNotification* pNF)
{
    IMarkupPointer* pStart = NULL;
    IHTMLElement* pIElement = NULL;
    HRESULT hr = S_OK;
    BOOL fRefTaken = FALSE;

    if(pNF->Type() != NTYPE_DELAY_LOAD_HISTORY)
    {
        super::Notify(pNF);
    }
    switch(pNF->Type())
    {
    case NTYPE_ELEMENT_GOTMNEMONIC:
        {
            {
                CDocument* pDoc = Doc();
                fRefTaken = TRUE;
                hr = pDoc->CreateMarkupPointer(&pStart);
                if(hr)
                {
                    goto Cleanup;
                }

                hr = this->QueryInterface(IID_IHTMLElement, (void**)&pIElement);
                if(hr)
                {
                    goto Cleanup;
                }

                hr = pStart->MoveAdjacentToElement(pIElement, ELEM_ADJ_BeforeEnd);
                if(hr)
                {
                    goto Cleanup;
                }

                hr = pDoc->Select(pStart, pStart, SELECTION_TYPE_Caret);                
                if(hr)
                {
                    goto Cleanup;
                }
            }                 
        }            
        break;

    case NTYPE_ELEMENT_LOSTMNEMONIC:
        {
            {
                Doc()->NotifySelection(SELECT_NOTIFY_DESTROY_ALL_SELECTION, NULL);
            }                 
        }        
        break;

    case NTYPE_DELAY_LOAD_HISTORY:
        super::Notify(pNF);
        break;

    case NTYPE_END_PARSE:
        GetValueHelper(&_cstrDefaultValue);
        _fLastValueSet = FALSE;
        break;

        // BUGBUG: this might have problems with undo (jbeda)
    case NTYPE_ELEMENT_EXITTREE_1:
        break;

    case NTYPE_SET_CODEPAGE:
        _fChangingEncoding = TRUE;
        break;
    }

Cleanup:
    if(fRefTaken)
    {
        ReleaseInterface(pStart);
        ReleaseInterface(pIElement);
    }
}

//+------------------------------------------------------------------------
//
//  Member:     CRichtext::GetValueHelper()
//
//  Synopsis:   returns the inner HTML
//
//-------------------------------------------------------------------------
HRESULT CRichtext::GetValueHelper(CString* pstr)
{
    HRESULT hr = S_OK;

    if(Tag()==ETAG_TEXTAREA)
    {
        hr = GetPlainTextInScope(pstr);
    }
    else
    {
        BSTR bStrValue;

        hr = GetText(&bStrValue, WBF_NO_WRAP|WBF_NO_TAG_FOR_CONTEXT);
        if(hr)
        {
            goto Cleanup;
        }

        Assert(pstr);
        hr = pstr->SetBSTR(bStrValue);
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CRichtext::GetWordWrap()
//
//  Synopsis:   Callback to tell if the word wrap should be done
//
//-------------------------------------------------------------------------
BOOL CRichtext::GetWordWrap() const
{
    return  htmlWrapOff==GetAAwrap()?FALSE:TRUE;
}

HRESULT CRichtext::SetValueHelperInternal(CString* pstr, BOOL fOM/*=TRUE*/)
{
    HRESULT hr = S_OK;
    int c= pstr->Length();

    Assert(Tag()==ETAG_TEXTAREA);
    Assert(pstr);
    hr = Inject(Inside, FALSE, *pstr, c);

    // Set this to prevent OnPropertyChange(_Value_) from firing twice
    // when value is set through OM. This flag is cleared in OnTextChange().
    _fFiredValuePropChange = fOM;

    // bugbug, make sure this is covered when turns the HTMLAREA on
    _cstrLastValue.Set(*pstr, c);
    _fLastValueSet = TRUE;
    _fTextChanged = FALSE;

    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CRichtext::SetValueHelper()
//
//  Synopsis:   set the inner HTML
//
//-------------------------------------------------------------------------
HRESULT CRichtext::SetValueHelper(CString* pstr)
{
    return SetValueHelperInternal(pstr);
}

//+------------------------------------------------------------------------
//
//  Member:     CRichtext::ApplyDefaultFormat()
//
//  Synopsis:
//
//-------------------------------------------------------------------------
HRESULT CRichtext::ApplyDefaultFormat(CFormatInfo* pCFI)
{
    CUnitValue  uv;
    DWORD       dwRawValue;
    int         i;
    HRESULT     hr = S_OK;

    pCFI->PrepareFancyFormat();

    // No vertical spacing between para's
    pCFI->_ff()._cuvSpaceBefore.SetPoints(0);
    pCFI->_ff()._cuvSpaceAfter.SetPoints(0);

    // Border info
    pCFI->_ff()._ccvBorderColorLight.SetSysColor(COLOR_3DLIGHT);
    pCFI->_ff()._ccvBorderColorDark.SetSysColor(COLOR_3DDKSHADOW);
    pCFI->_ff()._ccvBorderColorHilight.SetSysColor(COLOR_BTNHIGHLIGHT);
    pCFI->_ff()._ccvBorderColorShadow.SetSysColor(COLOR_BTNSHADOW);

    uv.SetValue(2, CUnitValue::UNIT_PIXELS);
    dwRawValue = uv.GetRawValue();

    for(i=BORDER_TOP; i<=BORDER_LEFT; i++)
    {
        pCFI->_ff()._cuvBorderWidths[i].SetRawValue(dwRawValue);
        pCFI->_ff()._bBorderStyles[i] = fmBorderStyleSunken;
    }
    // End Border info

    // Add default padding and scrolling
    Assert(TEXT_INSET_DEFAULT_LEFT == TEXT_INSET_DEFAULT_RIGHT);
    Assert(TEXT_INSET_DEFAULT_LEFT == TEXT_INSET_DEFAULT_TOP);
    Assert(TEXT_INSET_DEFAULT_LEFT == TEXT_INSET_DEFAULT_BOTTOM);

    uv.SetValue(TEXT_INSET_DEFAULT_LEFT, CUnitValue::UNIT_PIXELS);

    pCFI->_ff()._cuvPaddingLeft   =
        pCFI->_ff()._cuvPaddingRight  =
        pCFI->_ff()._cuvPaddingTop    =
        pCFI->_ff()._cuvPaddingBottom = uv;

    pCFI->UnprepareForDebug();

    hr = super::ApplyDefaultFormat(pCFI);
    if(hr)
    {
        goto Cleanup;
    }

    if(pCFI->_pff->_bOverflowX==(BYTE)styleOverflowNotSet
        || pCFI->_pff->_bOverflowY==(BYTE)styleOverflowNotSet)
    {
        pCFI->PrepareFancyFormat();

        if(pCFI->_pff->_bOverflowX == (BYTE)styleOverflowNotSet)
        {
            pCFI->_ff()._bOverflowX = GetAAwrap()!=htmlWrapOff
                ? (BYTE)styleOverflowHidden : (BYTE)styleOverflowScroll;
        }

        if(pCFI->_pff->_bOverflowY == (BYTE)styleOverflowNotSet)
        {
            pCFI->_ff()._bOverflowY = (BYTE)styleOverflowScroll;
        }

        pCFI->UnprepareForDebug();
    }

Cleanup:
    RRETURN(hr);
}

HRESULT STDMETHODCALLTYPE CRichtext::select(void)
{
    HRESULT         hr          = S_OK;
    CMarkup*        pMarkup     = GetMarkup();
    CDocument*      pDoc        = Doc();
    CMarkupPointer  ptrStart(pDoc);
    CMarkupPointer  ptrEnd(pDoc); 
    IMarkupPointer* pIStart; 
    IMarkupPointer* pIEnd; 

    if(!pMarkup)
    {
        goto Cleanup;
    }

    // BUGBUG (MohanB) We need to make this current because that's the only
    // way selection works right now. GetCurrentSelectionRenderingServices()
    // looks for the current element. MarkA should fix this.
    hr = BecomeCurrent(0);
    if(hr)
    {
        goto Cleanup;
    }

    hr = ptrStart.MoveToCp(GetFirstCp(), pMarkup);
    if(hr)
    {
        goto Cleanup;
    }
    hr = ptrEnd.MoveToCp(GetLastCp(), pMarkup);
    if(hr)
    {
        goto Cleanup;
    }
    Verify(S_OK == ptrStart.QueryInterface(IID_IMarkupPointer, (void**)&pIStart));
    Verify(S_OK == ptrEnd.QueryInterface(IID_IMarkupPointer, (void**)&pIEnd));
    hr = pDoc->Select(pIStart, pIEnd, SELECTION_TYPE_Selection);
    pIStart->Release();
    pIEnd->Release();
Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CRichtext::RequestYieldCurrency
//
//  Synopsis:   Check if OK to Relinquish currency
//
//  Arguments:  BOOl fForce -- if TRUE, force change and ignore user cancelling the
//                             onChange event
//
//  Returns:    S_OK: ok to yield currency
//
//--------------------------------------------------------------------------
HRESULT CRichtext::RequestYieldCurrency(BOOL fForce)
{
    CString cstr;
    HRESULT hr = S_OK;

    if((hr=GetValueHelper(&cstr)) == S_OK)
    {
        BOOL fFire = FormsStringCmp(cstr, _fLastValueSet?_cstrLastValue : _cstrDefaultValue) != 0;

        if(!fFire)
        {
            goto Cleanup;
        }

        if(!Fire_onchange()) // JS event
        {
            hr = E_FAIL;
            goto Cleanup;
        }

        hr = super::RequestYieldCurrency(fForce);
        if(hr == S_OK)
        {
            _cstrLastValue.Set(cstr);
            _fLastValueSet = TRUE;
        }
    }

Cleanup:
    if(fForce && FAILED(hr))
    {
        hr = S_OK;
    }

    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     CRichText::BecomeUIActive
//
//  Synopsis:   Check imeMode to set state of IME.
//
//  Notes:      This is the method that external objects should call
//              to force sites to become ui active.
//
//--------------------------------------------------------------------------
HRESULT CRichtext::BecomeUIActive()
{
    HRESULT hr = S_OK;

    hr = super::BecomeUIActive();
    if(hr)
    {
        goto Cleanup;
    }

    hr = SetImeState();
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:    
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CRichtext::YieldCurrency
//
//  Synopsis:   Relinquish currency
//
//  Arguments:  pSiteNew    New site that wants currency
//
//  Returns:    HRESULT
//
//--------------------------------------------------------------------------
HRESULT CRichtext::YieldCurrency(CElement* pElemNew)
{
    HRESULT hr;

    _fDoReset = FALSE;

    Assert(HasLayout());

    hr = super::YieldCurrency(pElemNew);

    RRETURN(hr);
}

HRESULT STDMETHODCALLTYPE CRichtext::HandleMessage(CMessage* pMessage)
{
    HRESULT hr = S_FALSE;
    BOOL fEditable = IsEditable(TRUE);

    if(!CanHandleMessage() || (!fEditable && !IsEnabled()))
    {
        goto Cleanup;
    }

    if(!fEditable && _fDoReset)
    {
        if(pMessage->message==WM_KEYDOWN && pMessage->wParam==VK_ESCAPE)
        {
        }
        else if(pMessage->message>=WM_KEYFIRST &&
            pMessage->message<=WM_KEYLAST && pMessage->wParam!=VK_ESCAPE)
        {
            _fDoReset = FALSE;
        }
    }

    switch(pMessage->message)
    {
    case WM_RBUTTONDOWN:
        // Ignore right-click (single click) if the input text box has focus.
        // This prevents the selected contents in an input text control from
        // being loosing selection.
        {
            hr = BecomeCurrent(pMessage->lSubDivision);
        }                
        goto Cleanup;

        // We handle all WM_CONTEXTMENUs
    case WM_CONTEXTMENU:
        hr = OnContextMenu(
            (short)LOWORD(pMessage->lParam),
            (short)HIWORD(pMessage->lParam),
            CONTEXT_MENU_CONTROL);
        goto Cleanup;
    }

    if(!fEditable &&
        pMessage->message==WM_KEYDOWN &&
        pMessage->wParam==VK_ESCAPE)
    {
        _fDoReset = TRUE;
        SetValueHelperInternal(_fLastValueSet?&_cstrLastValue:&_cstrDefaultValue, FALSE);
        hr = S_FALSE;
        goto Cleanup;
    }

    // Let supper take care of event firing
    // Since we let TxtEdit handle messages we do JS events after
    // it comes back
    hr = super::HandleMessage(pMessage);

Cleanup:
    RRETURN1(hr, S_FALSE);
}

HRESULT CRichtext::put_status(VARIANT status)
{
    switch(status.vt)
    {
    case VT_NULL:
        _vStatus.vt = VT_NULL;
        break;
    case VT_BOOL:
        _vStatus.vt = VT_BOOL;
        V_BOOL(&_vStatus) = V_BOOL(&status);
        break;
    default:
        _vStatus.vt = VT_BOOL;
        V_BOOL(&_vStatus) = VB_TRUE;
    }

    Verify(S_OK == OnPropertyChange(DISPID_CRichtext_status, 0));

    RRETURN(S_OK);
}

HRESULT CRichtext::get_status(VARIANT* pStatus)
{
    if (_vStatus.vt==VT_NULL)
    {
        pStatus->vt = VT_NULL;
    }
    else
    {
        pStatus->vt = VT_BOOL;
        V_BOOL(pStatus) = V_BOOL(&_vStatus);
    }
    RRETURN(S_OK);
}

HRESULT CRichtext::DoReset(void)
{
    RRETURN(SetValueHelperInternal(&_cstrDefaultValue));
}