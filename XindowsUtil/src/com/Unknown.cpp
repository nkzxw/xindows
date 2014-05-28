
#include "stdafx.h"
#include "Unknown.h"

//+------------------------------------------------------------------------
//
//  Function:   ClearInterfaceFn
//
//  Synopsis:   Sets an interface pointer to NULL, after first calling
//              Release if the pointer was not NULL initially
//
//  Arguments:  [ppUnk]     *ppUnk is cleared
//
//-------------------------------------------------------------------------
void ClearInterfaceFn(IUnknown** ppUnk)
{
    IUnknown* pUnk;

    pUnk = *ppUnk;
    *ppUnk = NULL;
    if(pUnk)
    {
        pUnk->Release();
    }
}

//+------------------------------------------------------------------------
//
//  Function:   ClearClassFn
//
//  Synopsis:   Nulls a pointer to a class, releasing the class via the
//              provided IUnknown implementation if the original pointer
//              is non-NULL.
//
//  Arguments:  [ppv]
//              [pUnk]
//
//-------------------------------------------------------------------------
void ClearClassFn(void** ppv, IUnknown* pUnk)
{
    *ppv = NULL;
    if(pUnk)
    {
        pUnk->Release();
    }
}

//+------------------------------------------------------------------------
//
//  Function:   ReplaceInterfaceFn
//
//  Synopsis:   Replaces an interface pointer with a new interface,
//              following proper ref counting rules:
//
//              = *ppUnk is set to pUnk
//              = if *ppUnk was not NULL initially, it is Release'd
//              = if pUnk is not NULL, it is AddRef'd
//
//              Effectively, this allows pointer assignment for ref-counted
//              pointers.
//
//  Arguments:  [ppUnk]
//              [pUnk]
//
//-------------------------------------------------------------------------
void ReplaceInterfaceFn(IUnknown** ppUnk, IUnknown* pUnk)
{
    IUnknown* pUnkOld = *ppUnk;

    *ppUnk = pUnk;

    // Note that we do AddRef before Release; this avoids
    // accidentally destroying an object if this function
    // is passed two aliases to it
    if(pUnk)
    {
        pUnk->AddRef();
    }

    if(pUnkOld)
    {
        pUnkOld->Release();
    }
}

//+------------------------------------------------------------------------
//
//  Function:   ReleaseInterface
//
//  Synopsis:   Releases an interface pointer if it is non-NULL
//
//  Arguments:  [pUnk]
//
//-------------------------------------------------------------------------
void ReleaseInterface(IUnknown* pUnk)
{
    if(pUnk)
    {
        pUnk->Release();
    }
}
