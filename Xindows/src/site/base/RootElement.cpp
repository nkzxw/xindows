
#include "stdafx.h"
#include "RootElement.h"

const CElement::CLASSDESC CRootElement::s_classdesc =
{
    {
        NULL,                   // _pclsid
        NULL,                   // _pcpi
        ELEMENTDESC_NOLAYOUT,   // _dwFlags
        NULL,                   // _piidDispinterface
        NULL
    },
    NULL,

    NULL,                       // _paccelsDesign
    NULL                        // _paccelsRun
};

void CRootElement::Notify(CNotification* pNF)
{
    NOTIFYTYPE ntype = pNF->Type();

    CMarkup* pMarkup = GetMarkup();

    super::Notify(pNF);

    switch(ntype)
    {
    case NTYPE_SET_CODEPAGE:
        // Directly switch the codepage (do not call SwitchCodePage)
        {
            CDocument* pDoc = Doc();
            ULONG ulData;

            pNF->Data(&ulData);

            pDoc->_codepage = CODEPAGE(ulData);
            pDoc->_codepageFamily = WindowsCodePageFromCodePage(pDoc->_codepage);
        }
        break;

    case NTYPE_CLEAR_FORMAT_CACHES:
        GetFirstBranch()->VoidCachedInfo();
        break;

    case NTYPE_END_PARSE:
        pMarkup->SetLoaded(TRUE);

        break;
    }
    return;
}

HRESULT CRootElement::ComputeFormats(CFormatInfo* pCFI, CTreeNode* pNodeTarget)
{
    CDocument*      pDoc = Doc();
    THREADSTATE*    pts = GetThreadState();
    CColorValue     cv;
    COLORREF        cr;
    HRESULT         hr = S_OK;

    Assert(pCFI);
    Assert(SameScope( this, pNodeTarget));
    Assert(pCFI->_eExtraValues!=ComputeFormatsType_Normal || (pNodeTarget->_iCF==-1 && pNodeTarget->_iPF == -1 && pNodeTarget->_iFF==-1));

    pCFI->Reset();
    pCFI->_pNodeContext = pNodeTarget;

    // Setup Char Format
    if(pDoc->_icfDefault < 0)
    {
        hr = pDoc->CacheDefaultCharFormat();
        if(hr)
        {
            goto Cleanup;
        }
    }

    pCFI->_icfSrc = pDoc->_icfDefault;
    pCFI->_pcfSrc = pCFI->_pcf = pDoc->_pcfDefault;

    // Setup Para Format
    pCFI->_ipfSrc = pts->_ipfDefault;
    pCFI->_ppfSrc = pCFI->_ppf = pts->_ppfDefault;

    // Setup Fancy Format
    pCFI->_iffSrc = pts->_iffDefault;
    pCFI->_pffSrc = pCFI->_pff = pts->_pffDefault;

    cv = pDoc->GetAAbgColor();

    if(cv.IsDefined())
    {
        cr = cv.GetColorRef();
    }
    else
    {
        cr = pDoc->_pOptionSettings->crBack();
    }

    if(!pCFI->_pff->_ccvBackColor.IsDefined() || pCFI->_pff->_ccvBackColor.GetColorRef()!=cr)
    {
        pCFI->PrepareFancyFormat();
        pCFI->_ff()._ccvBackColor = cr;
        pCFI->UnprepareForDebug();
    }

    Assert(pCFI->_pff->_ccvBackColor.IsDefined());

    if(pCFI->_eExtraValues == ComputeFormatsType_Normal)
    {
        hr = pNodeTarget->CacheNewFormats(pCFI);
        if(hr)
        {
            goto Cleanup;
        }

        // If the doc codepage is Hebrew visual order, set the flag.
        pDoc->_fVisualOrder = (pDoc->GetCodePage() == CP_ISO_8859_8);
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CRootElement::YieldCurrency
//
//  Synopsis:
//
//----------------------------------------------------------------------------
HRESULT CRootElement::YieldCurrency(CElement* pElemNew)
{
    return super::YieldCurrency(pElemNew);
}

//+---------------------------------------------------------------------------
//
//  Member:     CRootElement::YieldUI
//
//  Synopsis:
//
//----------------------------------------------------------------------------
void CRootElement::YieldUI(CElement* pElemNew)
{
}

//+---------------------------------------------------------------------------
//
//  Member:     CRootElement::BecomeUIActive
//
//  Synopsis:
//
//----------------------------------------------------------------------------
HRESULT CRootElement::BecomeUIActive()
{
    HRESULT hr = S_OK;


    if(Doc()->_pElemUIActive==this && GetFocus()==Doc()->InPlace()->_hwnd)
    {
        return S_OK;
    }

    // Tell the document that we are now the UI active site.
    // This will deactivate the current UI active site.
    hr = Doc()->SetUIActiveElement(this);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}
