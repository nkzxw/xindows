
#include "stdafx.h"
#include "TxtSlave.h"

const CElement::CLASSDESC CTxtSlave::s_classdesc =
{
    {
        NULL,                           // _pclsid
        s_acpi,                         // _pcpi
        ELEMENTDESC_TEXTSITE |          // _dwFlags
        ELEMENTDESC_DONTINHERITSTYLE |
        ELEMENTDESC_SHOWTWS |
        ELEMENTDESC_VPADDING |
        ELEMENTDESC_HASDEFDESCENT,
        NULL,                           // _piidDispinterface
        NULL,                           // _apHdlDesc
    },
    NULL,                               // _pAccelsDesign
    NULL                                // _pAccelsRun
};

HRESULT CTxtSlave::CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElement)
{
    Assert(ppElement);

    Assert(pht->GetTag() == ETAG_TXTSLAVE);
    *ppElement = new CTxtSlave(pht->GetTag(), pDoc);

    RRETURN((*ppElement)?S_OK:E_OUTOFMEMORY);
}

HRESULT CTxtSlave::Init2(CInit2Context* pContext)
{
    // CAUTION: Need to keep track of changes in Init2 of the
    // base classes and port them as appropriate.
    return S_OK;
}

DWORD CTxtSlave::GetBorderInfo(CDocInfo* pdci, CBorderInfo* pborderinfo, BOOL fAll)
{
    return MarkupMaster()->GetBorderInfo(pdci, pborderinfo, fAll);
}

HRESULT CTxtSlave::ComputeFormats(CFormatInfo* pCFI, CTreeNode* pNodeTarget)
{
    HRESULT         hr = S_OK;
    CTreeNode*      pNodeMaster = NULL;
    CElement*       pElemMaster = MarkupMaster();
    CDocument*      pDoc = Doc();
    THREADSTATE*    pts  = GetThreadState();
    BOOL            fComputeFFOnly = pNodeTarget->_iCF != -1;
    COMPUTEFORMATSTYPE eExtraValues = pCFI->_eExtraValues;

    Assert(pCFI);
    Assert(SameScope(this, pNodeTarget));
    Assert(eExtraValues!=ComputeFormatsType_Normal || ((pNodeTarget->_iCF==-1 && pNodeTarget->_iPF==-1) || pNodeTarget->_iFF==-1));

    if(pElemMaster)
    {
        pNodeMaster = pElemMaster->GetFirstBranch();

        // Get the format of our master before applying our own format.
        if(pNodeMaster)
        {
            // If the master node has not computed formats yet, recursively compute them
            if(pNodeMaster->_iCF==-1 || pNodeMaster->_iFF==-1
                || eExtraValues==ComputeFormatsType_GetInheritedValue)
            {

                hr = pElemMaster->ComputeFormats(pCFI, pNodeMaster);

                if(hr)
                {
                    goto Cleanup;
                }
            }

            Assert(pNodeMaster->_iCF >= 0);
            Assert(pNodeMaster->_iPF >= 0);
            Assert(pNodeMaster->_iFF >= 0);
        }
    }

    // NOTE: From this point forward any errors must goto Error instead of Cleanup!
    pCFI->Reset();
    pCFI->_pNodeContext = pNodeTarget;

    if(pNodeMaster)
    {
        // Inherit para format directly from the master node.
        pCFI->_iffSrc = pNodeMaster->_iFF;
        pCFI->_pffSrc = pCFI->_pff = &(*pts->_pFancyFormatCache)[pCFI->_iffSrc];
        pCFI->_fHasExpandos = (pCFI->_pff->_iExpandos >= 0);

        if(!fComputeFFOnly)
        {
            // Inherit the Char and Para formats from the master node
            pCFI->_icfSrc = pNodeMaster->_iCF;
            pCFI->_pcfSrc = pCFI->_pcf = &(*pts->_pCharFormatCache)[pCFI->_icfSrc];
            pCFI->_ipfSrc = pNodeMaster->_iPF;
            pCFI->_ppfSrc = pCFI->_ppf = &(*pts->_pParaFormatCache)[pCFI->_ipfSrc];

            // If the parent had layoutness, clear the inner formats
            if(pCFI->_pcf->_fHasDirtyInnerFormats)
            {
                pCFI->PrepareCharFormat();
                pCFI->_cf().ClearInnerFormats();
                pCFI->UnprepareForDebug();
            }
            if(pCFI->_ppf->_fHasDirtyInnerFormats)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf().ClearInnerFormats();
                pCFI->UnprepareForDebug();
            }
            if(pCFI->_ppf->_fPre != pCFI->_ppf->_fPreInner
                || pCFI->_ppf->_fInclEOLWhite != pCFI->_ppf->_fInclEOLWhiteInner
                || pCFI->_ppf->_bBlockAlign != pCFI->_ppf->_bBlockAlignInner)
            {
                pCFI->PrepareParaFormat();
                pCFI->_pf()._fPre = pCFI->_pf()._fPreInner;
                pCFI->_pf()._fInclEOLWhite = pCFI->_pf()._fInclEOLWhiteInner;
                pCFI->_pf()._bBlockAlign = pCFI->_pf()._bBlockAlignInner;
                pCFI->UnprepareForDebug();
            }

            if(pCFI->_pcf->_fNoBreak != pCFI->_pcf->_fNoBreakInner)
            {
                pCFI->PrepareCharFormat();
                pCFI->_cf()._fNoBreak = pCFI->_cf()._fNoBreakInner;
                pCFI->UnprepareForDebug();
            }
        }
        else
        {
            pCFI->_icfSrc = pDoc->_icfDefault;
            pCFI->_pcfSrc = pCFI->_pcf = pDoc->_pcfDefault;
            pCFI->_ipfSrc = pts->_ipfDefault;
            pCFI->_ppfSrc = pCFI->_ppf = pts->_ppfDefault;
        }
    }
    else
    {
        pCFI->_iffSrc = pts->_iffDefault;
        pCFI->_pffSrc = pCFI->_pff = pts->_pffDefault;
        pCFI->_icfSrc = pDoc->_icfDefault;
        pCFI->_pcfSrc = pCFI->_pcf = pDoc->_pcfDefault;
        pCFI->_ipfSrc = pts->_ipfDefault;
        pCFI->_ppfSrc = pCFI->_ppf = pts->_ppfDefault;

        Assert(pCFI->_pffSrc->_pszFilters == NULL);
    }

    if(pCFI->_pff->_fHasLayout || !pCFI->_pff->_fBlockNess)
    {
        pCFI->PrepareFancyFormat();
        pCFI->_ff()._fHasLayout = FALSE;
        pCFI->_ff()._fBlockNess = TRUE;
        pCFI->UnprepareForDebug();
    }

    if(eExtraValues == ComputeFormatsType_Normal)
    {
        hr = pNodeTarget->CacheNewFormats(pCFI);
        if(hr)
        {
            goto Error;
        }

        // Cache whether an element is a block element or not for fast retrieval.
        pNodeTarget->_fBlockNess = TRUE;

        pNodeTarget->_fHasLayout = FALSE;
    }

Cleanup:
    RRETURN(hr);

Error:
    pCFI->Cleanup();
    goto Cleanup;
}