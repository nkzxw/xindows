
#include "stdafx.h"
#include "DropTarget.h"

//+---------------------------------------------------------------------------
//
//  Member:   CDropTarget::CDropTarget
//
//  Synopsis: Point to CServer for D&D Method delegation
//
//----------------------------------------------------------------------------
CDropTarget::CDropTarget(CServer* pServer)
{
    _pServer = pServer;
    _ulRef = 1;
    pServer->SubAddRef();

    MemSetName((this, "CDropTarget _pServer=%08x", pServer));
}

//+----------------------------------------------------------------------------
//
//  Member:     CDropTarget::~CDropTarget
//
//  Synopsis:
//
//-----------------------------------------------------------------------------
CDropTarget::~CDropTarget()
{
}

//+----------------------------------------------------------------------------
//
//  Member:     CDropTarget::QueryInterface
//
//  Synopsis:   Returns only IDropTarget and IUnknown interfaces. Does not
//              delegate QI calss to pServer
//
//----------------------------------------------------------------------------
STDMETHODIMP CDropTarget::QueryInterface(REFIID riid, void** ppv)
{
    if(IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDropTarget))
    {
        *ppv = (IDropTarget*)this;
    }
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDropTarget::AddRef
//
//  Synopsis:   AddRefs the parent server
//
//----------------------------------------------------------------------------
STDMETHODIMP_(ULONG) CDropTarget::AddRef()
{
    _pServer->SubAddRef();
    return ++_ulRef;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDropTarget::Release
//
//  Synopsis:   Releases the parent server
//
//----------------------------------------------------------------------------
STDMETHODIMP_(ULONG) CDropTarget::Release()
{
    _pServer->SubRelease();

    if(0 == --_ulRef)
    {
        delete this;
        return 0;
    }

    return _ulRef;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDropTarget::DragEnter
//
//  Synopsis:   Delegates to _pServer
//
//----------------------------------------------------------------------------
STDMETHODIMP CDropTarget::DragEnter( IDataObject* pIDataSource, DWORD grfKeyState,
                                    POINTL pt, DWORD* pdwEffect)
{
    return _pServer->DragEnter(pIDataSource, grfKeyState, pt, pdwEffect);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDropTarget::DragOver
//
//  Synopsis:   Delegates to _pServer
//
//----------------------------------------------------------------------------
STDMETHODIMP CDropTarget::DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
{
    return _pServer->DragOver(grfKeyState, pt, pdwEffect);
}

//+---------------------------------------------------------------------------
//
//  Member:     CDropTarget::DragLeave
//
//  Synopsis:   Delegates to _pServer
//
//----------------------------------------------------------------------------
STDMETHODIMP CDropTarget::DragLeave()
{
    return _pServer->DragLeave(FALSE); // fDrop is FALSE
}

//+---------------------------------------------------------------------------
//
//  Member:     CDropTarget::Drop
//
//  Synopsis:   Delegates to _pServer
//
//----------------------------------------------------------------------------
STDMETHODIMP CDropTarget::Drop(IDataObject* pIDataSource, DWORD grfKeyState,
                               POINTL pt, DWORD* pdwEffect)
{
    return _pServer->Drop(pIDataSource, grfKeyState, pt, pdwEffect);
}