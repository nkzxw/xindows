
#include "stdafx.h"
#include "TxtRange.h"



// Helper functions so that measurer can make empty lines
// have the same height they will have after we spring
// load them and put text in them.
long GetSpringLoadedHeight(IMarkupPointer* pmpPosition, CFlowLayout* pFlowLayout, short* pyDescentOut)
{
    CDocument*  pDoc = pFlowLayout->Doc();
    CTreeNode*  pNode = pFlowLayout->GetFirstBranch();
    CCcs*       pccs;
    CBaseCcs*   pBaseCcs;
    CCharFormat cfLocal = *(pNode->GetCharFormat());
    CCalcInfo   CI;
    int         yHeight = -1;
    CVariant    varIn, varOut;
    GUID        guidCmdGroup = CGID_MSHTML;
    HRESULT     hr;

    V_VT(&varIn) = VT_UNKNOWN;
    V_UNKNOWN(&varIn) = pmpPosition;

    hr = pDoc->Exec(&guidCmdGroup, IDM_COMPOSESETTINGS, 0, &varIn, &varOut);
    V_VT(&varIn) = VT_NULL;
    if(hr || V_VT(&varOut)==VT_NULL)
    {
        goto Cleanup;
    }

    // We now know that we have to apply compose settings on this line.
    // Get font size.
    V_VT(&varIn) = VT_I4;
    V_I4(&varIn) = IDM_FONTSIZE;
    hr = pDoc->Exec(&guidCmdGroup, IDM_COMPOSESETTINGS, 0, &varIn, &varOut);
    if(hr)
    {
        goto Cleanup;
    }

    // If we got a valid font size, apply it to the local charformat.
    if(V_VT(&varOut)==VT_I4 && V_I4(&varOut)!=-1)
    {
        int iFontSize = ConvertHtmlSizeToTwips(V_I4(&varOut));
        cfLocal.SetHeightInTwips(iFontSize);
    }

    // Get font name.
    V_VT(&varIn) = VT_I4;
    V_I4(&varIn) = IDM_FONTNAME;
    hr = pDoc->Exec(&guidCmdGroup, IDM_COMPOSESETTINGS, 0, &varIn, &varOut);
    if(hr)
    {
        goto Cleanup;
    }

    // If we got a valid font name, apply it to the local charformat.
    if(V_VT(&varOut) == VT_BSTR)
    {
        TCHAR* pstrFontName = V_BSTR(&varOut);
        cfLocal.SetFaceName(pstrFontName);
    }

    cfLocal._bCrcFont = cfLocal.ComputeFontCrc();

    CI.Init(pFlowLayout);

    pccs = fc().GetCcs(CI._hdc, &CI, &cfLocal);

    if(pccs)
    {
        pBaseCcs = pccs->GetBaseCcs();
        yHeight = pBaseCcs->_yHeight + pBaseCcs->_yOffset;

        if(pyDescentOut)
        {
            *pyDescentOut = (INT)pBaseCcs->_yDescent;
        }

        pccs->Release();
    }

Cleanup:
    return yHeight;
}

long GetSpringLoadedHeight(CCalcInfo* pci, CFlowLayout* pFlowLayout, CTreePos* ptp, long cp, short* pyDescentOut)
{
    CElement*   pElementContent = pFlowLayout->ElementContent();
    int         yHeight;

    Assert(pyDescentOut);

    if(pElementContent && pElementContent->HasFlag(TAGDESC_ACCEPTHTML))
    {
        CMarkup*        pMarkup = pFlowLayout->GetContentMarkup();
        CDocument*      pDoc = pMarkup->Doc();
        CMarkupPointer  mpComposeFont(pDoc);
        HRESULT         hr;

        hr = mpComposeFont.MoveToCp(cp, pMarkup);
        if(hr)
        {
            yHeight = -1;
            goto Cleanup;
        }

        yHeight = GetSpringLoadedHeight(&mpComposeFont, pFlowLayout, pyDescentOut);
    }
    else
    {
        DEBUG_ONLY(CMarkup* pMarkup = pFlowLayout->GetContentMarkup());
        DEBUG_ONLY(LONG junk);

        Assert(ptp);
        Assert(ptp == pMarkup->TreePosAtCp(cp, &junk));
        const CCharFormat* pCF = ptp->GetBranch()->GetCharFormat();
        CCcs* pccs = fc().GetCcs(pci->_hdc, pci, pCF);
        CBaseCcs* pBaseCcs;

        if(!pccs)
        {
            yHeight = -1;
            goto Cleanup;
        }
        pBaseCcs = pccs->GetBaseCcs();
        yHeight = pBaseCcs->_yHeight;
        *pyDescentOut = pBaseCcs->_yDescent;

        pccs->Release();
    }

Cleanup:
    return yHeight;
}