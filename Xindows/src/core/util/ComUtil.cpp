
#include "stdafx.h"
#include "ComUtil.h"

//+----------------------------------------------------------------------------
//
//  Member:     GetTypeInfoFromCoClass
//
//  Synopsis:   Return either the default dispinterface or default source
//              interface ITypeInfo from a coclass
//
//  Arguments:  pTICoClass - ITypeInfo for containing coclass
//              fSource    - Return either source (TRUE) or default (FALSE) interface
//              ppTI       - Location at which to return the interface (may be NULL)
//              piid       - Location at which to return ther interface IID (may be NULL)
//
//  Returns:    S_OK, E_xxxx
//
//-----------------------------------------------------------------------------
HRESULT GetTypeInfoFromCoClass(ITypeInfo* pTICoClass, BOOL fSource, ITypeInfo** ppTI, IID* piid)
{
    ITypeInfo*  pTI = NULL;
    TYPEATTR*   pTACoClass = NULL;
    TYPEATTR*   pTA = NULL;
    IID         iid;
    HREFTYPE    href;
    int         i;
    int         flags;
    HRESULT     hr;

    Assert(pTICoClass);

    if(!ppTI)
    {
        ppTI = &pTI;
    }
    if(!piid)
    {
        piid = &iid;
    }

    *ppTI = NULL;
    *piid = IID_NULL;

    hr = pTICoClass->GetTypeAttr(&pTACoClass);
    if(hr)
    {
        goto Cleanup;
    }
    Assert(pTACoClass->typekind == TKIND_COCLASS);

    for(i=0; i<pTACoClass->cImplTypes; i++)
    {
        hr = pTICoClass->GetImplTypeFlags(i, &flags);
        if(hr)
        {
            goto Cleanup;
        }

        if((flags&IMPLTYPEFLAG_FDEFAULT) &&
            ((fSource&&(flags&IMPLTYPEFLAG_FSOURCE)) ||
            (!fSource&&!(flags&IMPLTYPEFLAG_FSOURCE))))
        {
            hr = pTICoClass->GetRefTypeOfImplType(i, &href);
            if(hr)
            {
                goto Cleanup;
            }

            hr = pTICoClass->GetRefTypeInfo(href, ppTI);
            if(hr)
            {
                goto Cleanup;
            }

            hr = (*ppTI)->GetTypeAttr(&pTA);
            if(hr)
            {
                goto Cleanup;
            }

            *piid = pTA->guid;
            goto Cleanup;
        }
    }

    hr = E_FAIL;

Cleanup:
    if(pTA)
    {
        Assert(*ppTI);
        (*ppTI)->ReleaseTypeAttr(pTA);
    }
    ReleaseInterface(pTI);
    if(pTACoClass)
    {
        Assert(pTICoClass);
        pTICoClass->ReleaseTypeAttr(pTACoClass);
    }
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   ValidateInvoke
//
//  Synopsis:   Validates arguments to a call of IDispatch::Invoke.  A call
//              to this function takes less space than the function itself.
//
//----------------------------------------------------------------------------
HRESULT ValidateInvoke(DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
{
    if(pvarResult)
    {
        VariantInit(pvarResult);
    }

    if(pexcepinfo)
    {
        InitEXCEPINFO(pexcepinfo);
    }

    if(puArgErr)
    {
        *puArgErr = 0;
    }

    if(!pdispparams)
    {
        RRETURN(E_INVALIDARG);
    }

    return S_OK;
}

//+----------------------------------------------------------------------------
//
//  Function    :   ReadyStateInvoke
//
//  Synopsis    :   This helper function is called by the various invokes for 
//      those classes that support the ready state property. this centralizes the
//      logic and code for handling this case.
//
//  RETURNS :   S_OK,           readyState-get and no errors
//              E_INVALIDARG    readystate-get and errors
//              S_FALSE         not readystate-get
//
//-----------------------------------------------------------------------------
HRESULT ReadyStateInvoke(DISPID dispid, WORD wFlags, long lReadyState, VARIANT* pvarResult)
{
    HRESULT hr = S_FALSE;

    if(dispid == DISPID_READYSTATE)
    {
        if(pvarResult && (wFlags&DISPATCH_PROPERTYGET))
        {
            V_VT(pvarResult) = VT_I4;
            V_I4(pvarResult) = lReadyState;
            hr = S_OK;
        }
        else
        {
            hr = E_INVALIDARG;
        }
    }

    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------
//
// Function:    IsSameObject
//
// Synopsis:    Checks for COM identity
//
// Arguments:   pUnkLeft, pUnkRight
//
//+---------------------------------------------------------------
BOOL IsSameObject(IUnknown* pUnkLeft, IUnknown* pUnkRight)
{
    IUnknown *pUnk1, *pUnk2;

    if(pUnkLeft == pUnkRight)
    {
        return TRUE;
    }

    if(pUnkLeft==NULL || pUnkRight==NULL)
    {
        return FALSE;
    }

    if(SUCCEEDED(pUnkLeft->QueryInterface(IID_IUnknown, (LPVOID*)&pUnk1)))
    {
        pUnk1->Release();
        if(pUnk1 == pUnkRight)
        {
            return TRUE;
        }
        if(SUCCEEDED(pUnkRight->QueryInterface(IID_IUnknown, (LPVOID*)&pUnk2)))
        {
            pUnk2->Release();
            return pUnk1 == pUnk2;
        }
    }
    return FALSE;
}

//+-------------------------------------------------------------------------
//
//  Member:     FormSetClipboard(IDataObject *pdo)
//
//  Synopsis:   helper function to set the clipboard contents
//
//--------------------------------------------------------------------------
HRESULT FormSetClipboard(IDataObject* pdo)
{
    HRESULT hr;
    hr = OleSetClipboard(pdo);

    if(!hr && !GetPrimaryObjectCount())
    {
        hr = OleFlushClipboard();
    }
    else
    {
        TLS(pDataClip) = pdo;
    }

    RRETURN(hr);
}