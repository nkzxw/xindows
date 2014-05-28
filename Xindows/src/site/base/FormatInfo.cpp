
#include "stdafx.h"
#include "FormatInfo.h"

void CFormatInfo::Cleanup()
{
    if(_pAAExpando)
    {
        _pAAExpando->Free();
        _pAAExpando = NULL;
    }

    _cstrBgImgUrl.Free();
    _cstrLiImgUrl.Free();
    _cstrFilters.Free();
}

CAttrArray* CFormatInfo::GetAAExpando()
{
    if(_pAAExpando == NULL)
    {
        memset(&_AAExpando, 0, sizeof(_AAExpando));
        _pAAExpando = &_AAExpando;

        if(_pff->_iExpandos >= 0)
        {
            _pAAExpando->CopyExpandos(GetExpandosAttrArrayFromCacheEx(_pff->_iExpandos));
        }

        _fHasExpandos = TRUE;
    }

    return _pAAExpando;
}

void CFormatInfo::PrepareCharFormatHelper()
{
    Assert(_pcfSrc!=NULL && _pcf==_pcfSrc);
    _pcf = &_cfDst;
    memcpy(&_cfDst, _pcfSrc, sizeof(CCharFormat));
    _fPreparedCharFormat = TRUE;
}

void CFormatInfo::PrepareParaFormatHelper()
{
    Assert(_ppfSrc!=NULL && _ppf==_ppfSrc);
    _ppf = &_pfDst;
    memcpy(&_pfDst, _ppfSrc, sizeof(CParaFormat));
    _fPreparedParaFormat = TRUE;
}

void CFormatInfo::PrepareFancyFormatHelper()
{
    Assert(_pffSrc!=NULL && _pff==_pffSrc);
    _pff = &_ffDst;
    memcpy(&_ffDst, _pffSrc, sizeof(CFancyFormat));
    _fPreparedFancyFormat = TRUE;
}

HRESULT CFormatInfo::ProcessImgUrl(CElement* pElem, LPCTSTR lpszUrl, DISPID dispID, LONG* plCookie, BOOL fHasLayout)
{
    HRESULT hr = S_OK;

    if(lpszUrl && *lpszUrl)
    {
        hr = pElem->GetImageUrlCookie(lpszUrl, plCookie);
        if(hr)
        {
            goto Cleanup;
        }
    }
    else
    {
        // Return a null cookie.
        *plCookie = 0;
    }

    pElem->AddImgCtx(dispID, *plCookie);
    pElem->_fHasImage = (*plCookie != 0);

    if(dispID == DISPID_A_LIURLIMGCTXCACHEINDEX)
    {
        // url images require a request resize when modified
        pElem->_fResizeOnImageChange = *plCookie!=0;
    }
    else if(dispID == DISPID_A_BGURLIMGCTXCACHEINDEX)
    {
        // sites draw their own background, so we don't have to inherit
        // their background info
        if(!fHasLayout)
        {
            PrepareCharFormat();
            _cf()._fHasBgImage = (*plCookie!=0);
            UnprepareForDebug();
        }
    }

Cleanup:
    RRETURN(hr);
}
