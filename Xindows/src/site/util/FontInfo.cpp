
#include "stdafx.h"
#include "FontInfo.h"

HRESULT CFontInfoCache::AddInfoToAtomTable(LPCTSTR pchFaceName, LONG* plIndex)
{
    HRESULT hr = S_OK;
    LONG lIndex;
    CFontInfo* pfi;

    for(lIndex=0; lIndex<Size(); lIndex++)
    {
        pfi = (CFontInfo*)Deref(sizeof(CFontInfo), lIndex);
        if(!StrCmpIC(pchFaceName, pfi->_cstrFaceName))
        {
            break;
        }
    }

    if(lIndex == Size())
    {
        CFontInfo fi;
        // Not found, so add element to array.
        hr = AppendIndirect(&fi);
        if(hr)
        {
            goto Cleanup;
        }

        lIndex = Size() - 1;
        pfi = (CFontInfo*)Deref(sizeof(CFontInfo), lIndex);
        pfi->_sids = sidsNotSet;
        hr = pfi->_cstrFaceName.Set(pchFaceName);
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(plIndex)
    {
        *plIndex = lIndex;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CFontInfoCache::GetAtomFromName(LPCTSTR pch, LONG* plIndex)
{
    LONG        lIndex;
    HRESULT     hr = S_OK;
    CFontInfo*  pfi;

    for(lIndex=0; lIndex<Size(); lIndex++)
    {
        pfi = (CFontInfo*)Deref(sizeof(CFontInfo), lIndex);

        if(!StrCmpIC(pfi->_cstrFaceName, pch))
        {
            break;
        }
    }

    if(lIndex == Size())
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    if(plIndex)
    {
        *plIndex = lIndex;
    }

Cleanup:    
    RRETURN(hr);
}

HRESULT CFontInfoCache::GetInfoFromAtom(LONG lIndex, CFontInfo** ppfi)
{
    HRESULT hr = S_OK;
    CFontInfo* pfi;

    if(Size() <= lIndex)
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    pfi = (CFontInfo*)Deref(sizeof(CFontInfo), lIndex);
    *ppfi = pfi;

Cleanup:
    RRETURN1(hr, DISP_E_MEMBERNOTFOUND);
}

void CFontInfoCache::Free()
{
    CFontInfo* pCFontInfo;
    LONG i;

    for(i=0; i<Size(); i++)
    {
        pCFontInfo = (CFontInfo*)Deref(sizeof(CFontInfo), i);
        pCFontInfo->_cstrFaceName.Free();
    }
    DeleteAll();
}