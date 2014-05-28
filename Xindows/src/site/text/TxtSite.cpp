
#include "stdafx.h"
#include "TxtSite.h"

#define _cxx_
#include "../gen/txtedit.hdl"

//+------------------------------------------------------------------------
//
//  Member:     wConvScroll
//
//  Synopsis:   Interchange horizontal and vertical commands for vertical
//              text site.
//
//-------------------------------------------------------------------------
WORD wConvScroll(WORD wparam)
{
    switch(wparam)
    {
    case SB_BOTTOM:
        return SB_TOP;

    case SB_LINEDOWN:
        return SB_LINEUP;

    case SB_LINEUP:
        return SB_LINEDOWN;

    case SB_PAGEDOWN:
        return SB_PAGEUP;

    case SB_PAGEUP:
        return SB_PAGEDOWN;

    case SB_TOP:
        return SB_BOTTOM;

    default:
        return wparam;
    }
}

CElement::ACCELS CTxtSite::s_AccelsTxtSiteDesign = CElement::ACCELS(&CElement::s_AccelsElementDesign, 0);
CElement::ACCELS CTxtSite::s_AccelsTxtSiteRun    = CElement::ACCELS(&CElement::s_AccelsElementRun,    0);

//+------------------------------------------------------------------------
//
//  Member:     CTxtSite, ~CTxtSite
//
//  Synopsis:   Constructor/Destructor
//
//-------------------------------------------------------------------------
CTxtSite::CTxtSite(ELEMENT_TAG etag, CDocument* pDoc) : super(etag, pDoc)
{
    _fOwnsRuns = TRUE;
}

//+------------------------------------------------------------------------
//
//  Member:     CTxtSite::PrivateQueryInterface, IUnknown
//
//  Synopsis:   Private unknown QI.
//
//-------------------------------------------------------------------------
HRESULT CTxtSite::PrivateQueryInterface(REFIID iid, void** ppv)
{
    HRESULT hr;

    *ppv = NULL;

    if(iid == CLSID_CTextSite)
    {
        *ppv = this; // weak ref.
        return S_OK;
    }
    else
    {
        RRETURN(super::PrivateQueryInterface(iid, ppv));
    }

    (*(IUnknown**)ppv)->AddRef();
    return hr;
}

//+------------------------------------------------------------------
//
//  Members: [get/put]_scroll[top/left] and get_scroll[height/width]
//
//  Synopsis : IHTMLTextContainer members. _dp is in pixels.
//
//------------------------------------------------------------------
HRESULT CTxtSite::get_scrollHeight(long* plValue)
{
    RRETURN(super::get_scrollHeight(plValue));
}

HRESULT CTxtSite::get_scrollWidth(long* plValue)
{
    RRETURN(super::get_scrollWidth(plValue));
}

HRESULT CTxtSite::get_scrollTop(long* plValue)
{
    RRETURN(super::get_scrollTop(plValue));
}

HRESULT CTxtSite::get_scrollLeft(long* plValue)
{
    RRETURN(super::get_scrollLeft(plValue));
}

HRESULT CTxtSite::put_scrollTop(long lPixels)
{
    RRETURN(super::put_scrollTop(lPixels));
}

HRESULT CTxtSite::put_scrollLeft(long lPixels)
{
    RRETURN(super::put_scrollLeft(lPixels));
}

//+-----------------------------------------------------------------
//
//  member [put_|get_]onscroll
//
//  synopsis - just delegate to CElement. these are here because this
//      property migrated from here to elemetn.
//+-----------------------------------------------------------------
STDMETHODIMP CTxtSite::put_onscroll(VARIANT v)
{
    RRETURN(super::put_onscroll(v));
}

STDMETHODIMP CTxtSite::get_onscroll(VARIANT* p)
{
    RRETURN(super::get_onscroll(p));
}