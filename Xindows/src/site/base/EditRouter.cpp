
#include "stdafx.h"
#include "EditRouter.h"

//+-------------------------------------------------------------------------
//
//  CEditRouter constructor and destructors
//
//--------------------------------------------------------------------------
CEditRouter::CEditRouter()
{
    _pInternalCmdTarget = NULL;
    _pHostCmdTarget = NULL;
    _fHostChecked = FALSE;
}

void CEditRouter::Passivate()
{    
    ReleaseInterface(_pInternalCmdTarget);
    _pInternalCmdTarget = NULL;
    ReleaseInterface(_pHostCmdTarget);
    _pHostCmdTarget = NULL;
}

CEditRouter::~CEditRouter()
{
    Passivate(); // Dup call just in case it wasn't already done.
}

//+-------------------------------------------------------------------------
//
//  Method:     CEditRouter::ExecEditCommand
//
//  Synopsis:   Routes and editing command to the appropriate edit handler(s);
//              last parameter provides the context in which the edit operation
//              will be carried on by the edit handler.
//
// Notes: If the host provides an edit handler, then commands are routed there 
//        and creation of the default edit handler (Mshtmled) is deferred to later. 
//        In fact, the default handler may never be created, for instance, when
//        all commands are supported by the host handler.
//
//--------------------------------------------------------------------------
HRESULT CEditRouter::ExecEditCommand(
         GUID*          pguidCmdGroup,
         DWORD          nCmdID,
         DWORD          nCmdexecopt,
         VARIANTARG*    pvarargIn,
         VARIANTARG*    pvarargOut,
         IUnknown*      punkContext,
         CDocument*     pDoc)
{
    int idm;
    HRESULT hr = OLECMDERR_E_NOTSUPPORTED;

    if(punkContext==NULL || !pDoc)
    {
        hr = E_FAIL;
        goto Cleanup;
    }

    // If there is no host to route to yet...
    if(_pHostCmdTarget==NULL && !_fHostChecked)
    {
        hr = SetHostEditHandler(punkContext, pDoc);
        _fHostChecked = TRUE;
    }

    // Commands are described using a pair of CmdGroup and CmdID. IDMFromCmdID an IDM, which is
    // the canonical form of the CmdGroup + CmdID pair. 
    // If IDM is unknown it can only be sent to the host, since it may be a new command that the
    // host can handle. In this case the original CmdGroup + CmdID is sent to the host. Otherwise
    // the canonical form (IDM + CGID_MSHTML) is sent to the host.
    // MshtmlEd is only sent the canonical form, since it will not be able to handle unknown commands.
    idm = CBase::IDMFromCmdID(pguidCmdGroup, nCmdID);

    if(_pHostCmdTarget)
    {
        // Route command to the host edit handler first
        if(idm == IDM_UNKNOWN)
        {
            hr = _pHostCmdTarget->Exec(
                pguidCmdGroup, 
                nCmdID, 
                nCmdexecopt, 
                pvarargIn, pvarargOut);
        }
        else
        {
            hr = _pHostCmdTarget->Exec(
                &CGID_MSHTML, 
                idm, 
                nCmdexecopt, 
                pvarargIn, pvarargOut);
        }

        if(!FAILED(hr))
        {
            goto Cleanup; // The host handled the command so we're done.
        }
    }

    // Okay, the host didn't like that one. Lets send it to our internal editor
    // If we don't know what it is, bail...
    if(idm == IDM_UNKNOWN)
    {
        hr = OLECMDERR_E_NOTSUPPORTED;
        goto Cleanup;
    }

    // Dispatch to our internal handler based on the passed in punk
    if(_pInternalCmdTarget == NULL)
    {
        hr = SetInternalEditHandler(punkContext, pDoc, TRUE); // If we get here, our only hope is the editor
        if(hr)
        {
            goto Cleanup;
        }
    }

    hr = _pInternalCmdTarget->Exec(
        &CGID_MSHTML, 
        idm, 
        nCmdexecopt, 
        pvarargIn, pvarargOut);

Cleanup:
    SRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CEditRouter::QueryStatusEditCommand
//
//              Rout the QueryStatus to MshtmlEd.
//
//+-------------------------------------------------------------------------
HRESULT CEditRouter::QueryStatusEditCommand(
        GUID*       pguidCmdGroup,
        ULONG       cCmds,
        MSOCMD      rgCmds[],
        MSOCMDTEXT* pcmdtext,
        IUnknown*   punkContext,
        CDocument*  pDoc)
{ 
    ULONG   idm;
    HRESULT hr = S_OK;
    MSOCMD* pCmd = &rgCmds[0];
    MSOCMD  newCommand;

    // If there is no host to route to yet...
    if(_pHostCmdTarget==NULL && !_fHostChecked)
    {
        hr = SetHostEditHandler(punkContext, pDoc);
        _fHostChecked = TRUE;
    }

    //  Commands are described using a pair of CmdGroup and CmdID. IDMFromCmdID an IDM, which is
    //  the canonical form of the CmdGroup + CmdID pair. 
    //  If IDM is unknown it can only be sent to the host, since it may be a new command that the
    //  host can handle. In this case the original CmdGroup + CmdID is sent to the host. Otherwise
    //  the canonical form (IDM + CGID_MSHTML) is sent to the host.
    //  MshtmlEd is only sent the canonical form, since it will not be able to handle unknown commands.
    idm = CBase::IDMFromCmdID(pguidCmdGroup,pCmd->cmdID);

    newCommand.cmdf = 0;    // just initialize (alexa)
    newCommand.cmdID = idm; // store the new Command Id.

    if(_pHostCmdTarget)
    {
        // Route command to the host edit handler first
        if(idm == IDM_UNKNOWN)
        {
            hr = _pHostCmdTarget->QueryStatus(
                pguidCmdGroup, 
                cCmds,
                &newCommand ,
                pcmdtext);
        }
        else
        {
            hr = _pHostCmdTarget->QueryStatus(
                &CGID_MSHTML, 
                cCmds,
                &newCommand, 
                pcmdtext);
        }

        if(!FAILED(hr))
        {
            goto Cleanup; // The host handled the command so we're done.
        }
    }

    // If we don't know what it is, bail...
    if(idm == IDM_UNKNOWN)
    {
        hr = OLECMDERR_E_NOTSUPPORTED;
        goto Cleanup;
    }

    // Dispatch to our internal handler
    if(_pInternalCmdTarget == NULL)
    {
        hr = SetInternalEditHandler(punkContext, pDoc, FALSE); // Do NOT force creation of the editor
        if(hr)
        {
            goto Cleanup;
        }
    }

    if(_pInternalCmdTarget == NULL) // We are STILL null after last check...
    {
        switch(idm)
        {
            // PUT ALL THE COMMANDS THAT SHOULD BE ENABLED EVEN IF THERE ISN'T AN EDITOR TO CHECK WITH HERE
        case IDM_SELECTALL:
            newCommand.cmdf = MSOCMDSTATE_UP;
            break;

        default:
            newCommand.cmdf = MSOCMDSTATE_DISABLED;
        }
        goto Cleanup;
    }

    // _pInternalCmdTarget  can be null if an event fired that 
    // changed the document. sadness
    if(_pInternalCmdTarget)
    {
        hr = _pInternalCmdTarget->QueryStatus(
            &CGID_MSHTML,
            cCmds, &newCommand,
            pcmdtext);
    }

Cleanup:
    if(hr == S_OK)
    {
        pCmd->cmdf = newCommand.cmdf;
    }

    SRETURN(hr);
}

const CLSID SID_SEditCommandTarget =
{0x3050f4b5,0x98b5,0x11cf,0xbb,0x82,0x00,0xaa,0x00,0xbd,0xce,0x0b};
const CLSID CGID_EditStateCommands =
{0x3050f4b6,0x98b5,0x11cf,0xbb,0x82,0x00,0xaa,0x00,0xbd,0xce,0x0b};
//+-------------------------------------------------------------------------
//
//  Method:     CEditRouter::SetHostEditHandler
//
//  Synopsis:   Helper routine to get an IOleCommandTarget from the host 
//              (if any), where editing commands will be routed.
//--------------------------------------------------------------------------
HRESULT CEditRouter::SetHostEditHandler(IUnknown* punkContext, CDocument* pDoc)
{
    IServiceProvider* pServiceProvider = 0;
    VARIANTARG varParamIn;
    HRESULT hr = S_FALSE;

    hr = pDoc->QueryInterface(IID_IServiceProvider, (void**)&pServiceProvider);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pServiceProvider->QueryService(SID_SEditCommandTarget, 
        IID_IOleCommandTarget, (void**)&_pHostCmdTarget);
    if(hr)
    {
        goto Cleanup;
    }

    V_VT(&varParamIn) = VT_UNKNOWN;
    V_UNKNOWN(&varParamIn) = punkContext;
    hr = _pHostCmdTarget->Exec(&CGID_EditStateCommands, 
        IDM_CONTEXT, 0, &varParamIn, NULL);
    if(hr)
    {
        // Host will not be able to process edit commands properly, 
        // since it failed to handle EditStateCommand. Nullify HostEditHandler.
        ClearInterface(&_pHostCmdTarget);
        goto Cleanup;
    }

Cleanup:
    ReleaseInterface(pServiceProvider);

    RRETURN(hr);
}

HRESULT CEditRouter::SetInternalEditHandler(IUnknown* punkContext, CDocument* pDoc, BOOL fForceCreate)
{
    HRESULT         hr = S_OK;
    IHTMLEditor*    ped = NULL;
    IUnknown*       punk = NULL;
    IHTMLDocument*  pTest = NULL;

    AssertSz(_pInternalCmdTarget==NULL , "CEditRouter::SetInternalEditHandler called when it already has one.");
    if(_pInternalCmdTarget != NULL)
    {
        goto Cleanup;
    }

    // See if we are being passed in a range as our context
    hr = punkContext->QueryInterface(IID_IHTMLDocument , (void**)&pTest);
    ReleaseInterface(pTest);

    if(!hr)
    {
        // our context is the Document. Dispatch to the editor.
        ped = pDoc->GetHTMLEditor(fForceCreate);
        if(ped == NULL)
        {   
            hr = S_OK;
            _pInternalCmdTarget = NULL;
            goto Cleanup;
        }

        hr = ped->QueryInterface(IID_IOleCommandTarget, (void**)&_pInternalCmdTarget);
    }
    else
    {
        // our context is probably a range. Get a range target from the doc            
        ped = pDoc->GetHTMLEditor(TRUE);
        if(ped == NULL)
        {
            goto Cleanup;
        }

        // Get a segment list to pass to the CommandTarget
        // BugBug : We need a cleaner way to pass a weak ref to the range        
        hr = ped->GetRangeCommandTarget(static_cast<ISegmentList*>(punkContext), &punk);
        if(hr)
        {
            goto Cleanup;
        }

        hr = punk->QueryInterface(IID_IOleCommandTarget, 
            (void**)&_pInternalCmdTarget);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    ReleaseInterface(punk);
    RRETURN(hr);
}
