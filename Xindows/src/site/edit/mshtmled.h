//+---------------------------------------------------------------------------
//
//  Copyright (C) Microsoft Corporation, 1998.
//
//  Class:      CMshtmlEd
//
//  Contents:   Definition of CMshtmlEd interfaces. This class is used to 
//                dispatch IOleCommandTarget commands to Tridnet range objects.
//
//  History:    7-Jan-98   raminh  Created
//----------------------------------------------------------------------------

#ifndef __XINDOWS_SITE_EDIT_MSHTMLED_H__
#define __XINDOWS_SITE_EDIT_MSHTMLED_H__

#include "sload.h"

class CSpringLoader;
class CXindowsEditor;


class CMshtmlEd : public IOleCommandTarget
{
public:

    CMshtmlEd( CXindowsEditor * pEd ) ;
    ~CMshtmlEd();

    DECLARE_MEMCLEAR_NEW_DELETE()

    // --------------------------------------------------
    // IUnknown Interface
    // --------------------------------------------------

    STDMETHODIMP_(ULONG)
    AddRef( void ) ;

    STDMETHODIMP_(ULONG)
    Release( void ) ;

    STDMETHODIMP
    QueryInterface(
        REFIID              iid, 
        LPVOID *            ppv ) ;


    // --------------------------------------------------
    // IOleCommandTarget methods
    // --------------------------------------------------

    STDMETHOD (QueryStatus)(
        const GUID * pguidCmdGroup,
        ULONG cCmds,
        OLECMD rgCmds[],
        OLECMDTEXT * pcmdtext);

    STDMETHOD (Exec)(
        const GUID * pguidCmdGroup,
        DWORD nCmdID,
        DWORD nCmdexecopt,
        VARIANTARG * pvarargIn,
        VARIANTARG * pvarargOut);


    HRESULT Initialize( IUnknown * pContext );

    CSpringLoader * GetSpringLoader() { return &_sl; }
    CXindowsEditor   * GetEditor()       { return _pEd; }
    HRESULT         GetSegmentList( ISegmentList** ppSegmentList ); 

private:
    CMshtmlEd() { }            // Protect the default constructor
    
    BOOL IsDialogCommand(DWORD nCmdexecopt, DWORD nCmdId, VARIANT *pvarargIn);

    LONG            _cRef;
    CXindowsEditor   * _pEd;      // The editor that we work for
    IUnknown      * _pContext; // The segment list context of this command router
    CSpringLoader   _sl;       // The spring loader                                        
};

#endif //_MSHTMLED_HXX_
