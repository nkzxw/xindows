//+------------------------------------------------------------------------
//
//  Copyright (C) Microsoft Corporation, 1998.
//
//  File:       MSHTMLED.CXX
//
//  Contents:   Implementation of Mshtml Editing Component
//
//  Classes:    CMshtmlEd
//
//  History:    7-Jan-98   raminh  Created
//             12-Mar-98   raminh  Converted over to use ATL
//-------------------------------------------------------------------------

#include "stdafx.h"
#include "mshtmled.h"

#include "edutil.h"
#include "edcmd.h"
#include "blockcmd.h"

extern HRESULT ShowEditDialog(UINT idm, VARIANT * pvarExecArgIn, HWND hwndParent, VARIANT * pvarArgReturn);

extern "C" const GUID CGID_EditStateCommands;

//+---------------------------------------------------------------------------
//
//  CMshtmlEd Constructor
//
//----------------------------------------------------------------------------
CMshtmlEd::CMshtmlEd( CXindowsEditor * pEd )
                        : _sl(this)
{ 
    _pEd = pEd;
    _pContext = NULL;
}


CMshtmlEd::~CMshtmlEd()
{
//  delete _psl;
}


HRESULT
CMshtmlEd::Initialize( IUnknown * pContext )
{
    _pContext = pContext;
    return(S_OK);
}    


//////////////////////////////////////////////////////////////////////////
//
//  Public Interface CCaret::IUnknown's Implementation
//
//////////////////////////////////////////////////////////////////////////

STDMETHODIMP_(ULONG)
CMshtmlEd::AddRef( void )
{
    return( ++_cRef );
}


STDMETHODIMP_(ULONG)
CMshtmlEd::Release( void )
{
    --_cRef;

    if( 0 == _cRef )
    {
        delete this;
        return 0;
    }

    return _cRef;
}


STDMETHODIMP
CMshtmlEd::QueryInterface(
    REFIID              iid, 
    LPVOID *            ppv )
{
    if (!ppv)
        RRETURN(E_INVALIDARG);

    *ppv = NULL;
    
    if( iid == IID_IUnknown )
    {
        *ppv = static_cast< IUnknown * >( this );
    }
    else if( iid == IID_IOleCommandTarget )
    {
        *ppv = static_cast< IOleCommandTarget * >( this );
    }

    if (*ppv)
    {
        ((IUnknown *)*ppv)->AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
    
}

BOOL
CMshtmlEd::IsDialogCommand(DWORD nCmdexecopt, DWORD nCmdID, VARIANT *pvarargIn)
{
    BOOL bResult = FALSE;
    
    if (nCmdID == IDM_HYPERLINK)
    {
        bResult = (pvarargIn == NULL) || (nCmdexecopt != OLECMDEXECOPT_DONTPROMPTUSER);
    }
    else if (nCmdID == IDM_IMAGE || nCmdID == IDM_FONT)
    {
        bResult = (nCmdexecopt != OLECMDEXECOPT_DONTPROMPTUSER);
    }
    
    return bResult;
}


//+---------------------------------------------------------------------------
//
//  CMshtmlEd IOleCommandTarget Implementation for Exec() 
//
//----------------------------------------------------------------------------
STDMETHODIMP
CMshtmlEd::Exec( const GUID *       pguidCmdGroup,
                       DWORD        nCmdID,
                       DWORD        nCmdexecopt,
                       VARIANTARG * pvarargIn,
                       VARIANTARG * pvarargOut)
{
    HRESULT hr = OLECMDERR_E_NOTSUPPORTED;
    CCommand* theCommand = NULL;

    // ShowHelp is not implemented 
    if(nCmdexecopt == OLECMDEXECOPT_SHOWHELP)
        goto Cleanup;

    Assert( *pguidCmdGroup == CGID_MSHTML );
    Assert( _pEd );
    Assert( _pEd->GetDoc() );
    Assert( _pContext );

    if (IsDialogCommand(nCmdexecopt, nCmdID, pvarargIn) )
    {
        // Special case for image or hyperlink or font dialogs
        theCommand = _pEd->GetCommandTable()->Get( ~nCmdID );    
    }
    else
    {
        theCommand = _pEd->GetCommandTable()->Get( nCmdID );
    }

    if ( theCommand )
    {
        hr = theCommand->Exec( nCmdexecopt, pvarargIn, pvarargOut, this );
    }
    else
    {
        hr = OLECMDERR_E_NOTSUPPORTED;
    }

Cleanup:
    RRETURN( hr );
}

//+---------------------------------------------------------------------------
//
//  CMshtmlEd IOleCommandTarget Implementation for QueryStatus()
//
//  Note: QueryStatus() is still being handled by Trident
//----------------------------------------------------------------------------
STDMETHODIMP CMshtmlEd::QueryStatus(
        const GUID * pguidCmdGroup,
        ULONG cCmds,
        OLECMD rgCmds[],
        OLECMDTEXT * pcmdtext)
{
    HRESULT             hr = OLECMDERR_E_NOTSUPPORTED ;
    CCommand  *         theCommand = NULL;
    OLECMD  *           pCmd = &rgCmds[0];

    Assert( *pguidCmdGroup == CGID_MSHTML );
    Assert( _pEd );
    Assert( _pEd->GetDoc() );
    Assert( _pContext );
    
    // BUGBUG: The dialog commands are hacked with strange tagId's.  So, for now we just
    // make sure the right command gets the query status [ashrafm]
    
    if (pCmd->cmdID == IDM_FONT)
    {
        theCommand = _pEd->GetCommandTable()->Get( ~(pCmd->cmdID)  );
    }
    else
    {
        theCommand = _pEd->GetCommandTable()->Get( pCmd->cmdID  );
    }
    
    if (theCommand )
    {
        hr = theCommand->QueryStatus( pCmd, pcmdtext, this );                      
    }
    else 
    {
        hr = OLECMDERR_E_NOTSUPPORTED;
    }

    RRETURN ( hr ) ;

}

HRESULT CMshtmlEd::GetSegmentList(ISegmentList** ppSegmentList) 
{ 
    HRESULT hr = S_OK;
    ISegmentList* pOut = NULL;
    hr = _pContext->QueryInterface(IID_ISegmentList, (void**)&pOut);
    *ppSegmentList = pOut;
    RRETURN(hr);
}
