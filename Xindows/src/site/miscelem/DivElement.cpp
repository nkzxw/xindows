
#include "stdafx.h"
#include "DivElement.h"

#define _cxx_
#include "../gen/div.hdl"

const CElement::CLASSDESC CDivElement::s_classdesc =
{
    {
        &CLSID_HTMLDivElement,              // _pclsid
        s_acpi,                             // _pcpi
        0,                                  // _dwFlags
        &IID_IHTMLDivElement,               // _piidDispinterface
        &s_apHdlDescs,                      // _apHdlDesc
    },
    (void*)s_apfnpdIHTMLDivElement,         // _pfnTearOff
    NULL,                                   // _pAccelsDesign
    NULL                                    // _pAccelsRun
};

HRESULT CDivElement::CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElement)
{
    Assert(pht->Is(ETAG_DIV));
    Assert(ppElement);
    *ppElement = new CDivElement(pDoc);
    return *ppElement?S_OK:E_OUTOFMEMORY;
}

//+----------------------------------------------------------------------------
//
//  Member:     CDivElement::PrivateQueryInterface, IUnknown
//
//  Synopsis:   Private unknown QI.
//
//-----------------------------------------------------------------------------
HRESULT CDivElement::PrivateQueryInterface(REFIID iid, void** ppv)
{
    HRESULT hr;

    *ppv = NULL;

    if(iid == IID_IHTMLControlElement)
    {

        hr = CreateTearOffThunk(
            this,
            s_apfnpdIHTMLControlElement,
            NULL,
            ppv,
            (void*)s_ppropdescsInVtblOrderIHTMLControlElement);
        if(hr)
        {
            RRETURN(hr);
        }
    }
    else if(iid == IID_IHTMLTextContainer)
    {
        hr = CreateTearOffThunk(
            this,
            (void*)s_apfnIHTMLTextContainer,
            NULL,
            ppv);
        if(hr)
        {
            RRETURN(hr);
        }
    }
    else
    {
        RRETURN(super::PrivateQueryInterface(iid, ppv));
    }

    if(!*ppv)
    {
        RRETURN(E_NOINTERFACE);
    }

    ((IUnknown*)*ppv)->AddRef();

    return S_OK;
}