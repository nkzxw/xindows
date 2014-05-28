
#include "stdafx.h"
#include "Base.h"

#include <ActivScp.h>

// The index into the list of event iid's that is the primary event iid.
#define CPI_OFFSET_PRIMARYDISPEVENT     1

#ifdef _WIN64
DWORD g_dwCookieForWin64 = 0;
#endif

BEGIN_TEAROFF_TABLE(CBase, IOleCommandTarget)
    TEAROFF_METHOD(CBase, &QueryStatus, querystatus, (GUID* pguidCmdGroup, ULONG cCmds, MSOCMD rgCmds[], MSOCMDTEXT* pcmdtext))
    TEAROFF_METHOD(CBase, &Exec, exec, (GUID* pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANTARG* pvarargIn, VARIANTARG* pvarargOut))
END_TEAROFF_TABLE()

BEGIN_TEAROFF_TABLE(CBase, IObjectIdentity)
    TEAROFF_METHOD(CBase, &IsEqualObject, isequalobject, (IUnknown*))
END_TEAROFF_TABLE()

#ifdef _DEBUG
ULONG CBase::s_ulLastSSN = 0;
extern BOOL GWHasPostedMethod(void* pvObject);
#endif

//+---------------------------------------------------------------
//
//  Member:     CBase::CBase, public
//
//  Synopsis:   Constructor.
//
//----------------------------------------------------------------
CBase::CBase()
{
    _pAA = NULL;
    _ulRefs = 1;
    _ulAllRefs = 1;

#ifdef _DEBUG
    _ulSSN = ++s_ulLastSSN;
    MemSetName((this, "SSN=%d", _ulSSN));
#endif
}

CBase::~CBase()
{
    Assert("Ref count messed up in derived dtor?" &&
        (_ulAllRefs==ULREF_IN_DESTRUCTOR||_ulAllRefs==1) &&
        (_ulRefs==ULREF_IN_DESTRUCTOR||_ulRefs==1));
    Assert(_pAA == NULL);
    Assert(!GWHasPostedMethod(this));
}

//+---------------------------------------------------------------
//
//  Member:     CBase::Init
//
//  Synopsis:   Nothing happens.
//
//----------------------------------------------------------------
HRESULT CBase::Init()
{
    // Assert that the cpi classdesc is setup correctly.  It should either
    // be NULL or have at least two entries.  The second entry can be
    // CPI_ENTRY_NULL.
    Assert(!BaseDesc()->_pcpi || BaseDesc()->_pcpi[0].piid);

    return S_OK;
}

//+---------------------------------------------------------------
//
//  Member:     CBase::Passivate
//
//  Synopsis:   Shutdown main object by releasing references to
//              other objects and generally cleaning up.  This
//              function is called when the main reference count
//              goes to zero.  The destructor is called when
//              the reference count for the main object and all
//              embedded sub-objects goes to zero.
//
//---------------------------------------------------------------
void CBase::Passivate()
{
    Assert("CBase::Passivate called unexpectedly or refcnt messed up in derived Passivate" &&
        (_ulRefs==ULREF_IN_DESTRUCTOR||_ulAllRefs==1));

    delete _pAA;
    _pAA = NULL;
}

//+---------------------------------------------------------------
//
//  Member:     CBase::PrivateAddRef, IPrivateUnknown
//
//---------------------------------------------------------------
ULONG CBase::PrivateAddRef()
{
    Assert("CBase::PrivateAddRef called after CBase::Passivate." && _ulRefs!=0);
    _ulRefs += 1;
    return 0;
}

//+---------------------------------------------------------------
//
//  Member:     CBase::PrivateRelease, IPrivateUnknown
//
//---------------------------------------------------------------
ULONG CBase::PrivateRelease()
{
    ULONG ulRefs = --_ulRefs;
    if(ulRefs == 0)
    {
        _ulRefs = ULREF_IN_DESTRUCTOR;
        Passivate();
        Assert("Unexpected refcnt on return from CBase::Passivate" && _ulRefs==ULREF_IN_DESTRUCTOR);
        _ulRefs = 0;
        SubRelease();
    }

    return ulRefs;
}

//+---------------------------------------------------------------
//
//  Member:     CBase::PrivateQueryInterface, IPrivateUnknown
//
//---------------------------------------------------------------
HRESULT CBase::PrivateQueryInterface(REFIID iid, void** ppvObj)
{
    *ppvObj = NULL;
    return E_NOINTERFACE;
}

//+---------------------------------------------------------------------------
//
//  Member:     CBase::ExpandedRelativeUrlInVariant
//
//  Synopsis:   Used by CBase::getAttribute to expand an URL if the property
//              retrieved is an URL and the GETMEMBER_ABSOLUTE is specified.
//
//----------------------------------------------------------------------------
HRESULT CBase::ExpandedRelativeUrlInVariant(VARIANT* pVariantURL)
{
    HRESULT hr = S_OK;
    TCHAR   cBuf[pdlUrlLen];
    TCHAR*  pchUrl = cBuf;

    if(pVariantURL && V_VT(pVariantURL)==VT_BSTR)
    {
        BSTR            bstrURL;
        IHTMLElement*   pElement;
        CElement*       pCElement;

        // Are we really an element?
        if(!PrivateQueryInterface(IID_IHTMLElement, (void**)&pElement))
        {
            ReleaseInterface(pElement);

            pCElement = DYNCAST(CElement, this);

            hr = pCElement->Doc()->ExpandUrl(V_BSTR(pVariantURL), ARRAYSIZE(cBuf), pchUrl, pCElement);
            if(hr)
            {
                goto Cleanup;
            }

            hr = FormsAllocString(pchUrl, &bstrURL);
            if(hr)
            {
                goto Cleanup;
            }

            VariantClear(pVariantURL);

            V_BSTR(pVariantURL) = bstrURL;
            V_VT(pVariantURL) = VT_BSTR;
        }
    }
    else
    {
        hr = S_OK;
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     FindPropDescFromDispID (private)
//
//  Synopsis:   Find a PROPERTYDESC based on the dispid.
//
//  Arguments:  dispidMember    - PROPERTY/METHOD to find
//              ppDesc          - PROPERTYDESC/METHODDESC found
//              pwEntry         - Byte offset in v-table of virtual function
//
//  Returns:    S_OK            everything is fine
//              E_INVALIDARG    ppDesc is NULL
//              S_FALSE         dispidMember not found in PROPDESC array
//--------------------------------------------------------------------------
#define AUTOMATION_VTBL_ENTRIES     7 // Includes IUnknown + IDispatch functions

HRESULT CBase::FindPropDescFromDispID(DISPID dispidMember, PROPERTYDESC** ppDesc, WORD* pwEntry, WORD* pwIIDIndex)
{
    HRESULT hr = S_OK;
    const VTABLEDESC* pVTblArray = NULL;

    if(!ppDesc)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *ppDesc = NULL;
    if(pwEntry)
    {
        *pwEntry = 0;
    }
    if(pwIIDIndex)
    {
        *pwIIDIndex = 0;
    }

    // Check cache before we try a linear search?
    if(!pVTblArray)
    {
        // Not in cache, do it the hard way -- linear.
        pVTblArray = GetVTableArray();

        // Look for known property or method in vtable interface array.
        if(pVTblArray && pVTblArray->pPropDesc)
        {
            while(pVTblArray->pPropDesc)
            {
                if(dispidMember == pVTblArray->pPropDesc->GetDispid())
                {
                    goto Found;
                }

                pVTblArray++;
            }

            goto NotFound;
        }
    }

Found:
    if(pVTblArray)
    {
        *ppDesc = const_cast<PROPERTYDESC*>(pVTblArray->pPropDesc);

        if(pwIIDIndex) 
        {
            *pwIIDIndex = pVTblArray->uVTblEntry >> 8;
        }
        if(pwEntry)
        {
            *pwEntry = ((pVTblArray->uVTblEntry&0xff) + AUTOMATION_VTBL_ENTRIES) * sizeof(void*); // Adjust for IDispatch methods
        }
    }
    else
    {
        hr = S_FALSE;
        goto Cleanup;
    }

Cleanup:
    RRETURN1(hr, S_FALSE);

NotFound:
    // No match found.
    hr = S_FALSE;
    goto Cleanup;
}

//+-------------------------------------------------------------------------
//
//  Function:   GetArgsActual
//
//  Synopsis:   helper
//
//--------------------------------------------------------------------------
LONG GetArgsActual(DISPPARAMS* pdispparams)
{
    LONG cArgsActual;
    // If the parameters passed in has the named DISPID_THIS paramter, then we
    // don't want to include this parameter in the total number of parameter
    // count.  This named parameter is an additional parameter tacked on by the
    // script engine to handle scoping rules.
    cArgsActual = pdispparams->cArgs;
    if(pdispparams->cNamedArgs && *pdispparams->rgdispidNamedArgs==DISPID_THIS)
    {
        cArgsActual--; // Don't include DISPID_THIS in argument count.
    }

    return cArgsActual;
}

//+---------------------------------------------------------------
//
//  member: CBase ::IsEqualObject
//
//  synopsis : default IObjectIdentity implementation.
//
//----------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CBase::IsEqualObject(IUnknown* pUnk)
{
    HRESULT hr = E_POINTER;
    IUnknown* pUnkThis = NULL;
    IUnknown* pUnkTarget=NULL;

    if(!pUnk)
    {
        goto Cleanup;
    }

    hr = PrivateQueryInterface(IID_IUnknown, (void**)&pUnkThis);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pUnk->QueryInterface(IID_IUnknown, (void**)&pUnkTarget);
    if(hr)
    {
        goto Cleanup;
    }

    hr = (pUnkTarget==pUnkThis) ? S_OK : S_FALSE;

Cleanup:
    ReleaseInterface(pUnkThis);
    ReleaseInterface(pUnkTarget);
    RRETURN1(hr, S_FALSE);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::GetEnabled
//
//  Synopsis:   Helper method.  Many objects will simply override this.
//
//--------------------------------------------------------------------------
STDMETHODIMP CBase::GetEnabled(VARIANT_BOOL* pfEnabled)
{
    if(!pfEnabled)
    {
        RRETURN(E_INVALIDARG);
    }

    *pfEnabled = VB_TRUE;
    return DISP_E_MEMBERNOTFOUND;
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::GetValid
//
//  Synopsis:   Helper method.  Many objects will simply override this.
//
//--------------------------------------------------------------------------
STDMETHODIMP CBase::GetValid(VARIANT_BOOL* pfValid)
{
    if(!pfValid)
    {
        RRETURN(E_INVALIDARG);
    }

    *pfValid = VB_TRUE;
    return DISP_E_MEMBERNOTFOUND;
}

struct MAP 
{
    short idm;
    USHORT usCmdID;
};

static MAP amapCommandSet95[] =
{
    IDM_OPEN,                   OLECMDID_OPEN,
    IDM_NEW,                    OLECMDID_NEW,
    IDM_SAVE,                   OLECMDID_SAVE,
    IDM_SAVEAS,                 OLECMDID_SAVEAS,
    IDM_SAVECOPYAS,             OLECMDID_SAVECOPYAS,
    IDM_PRINT,                  OLECMDID_PRINT,
    IDM_PRINTPREVIEW,           OLECMDID_PRINTPREVIEW,
    IDM_PAGESETUP,              OLECMDID_PAGESETUP,
    IDM_SPELL,                  OLECMDID_SPELL,
    IDM_PROPERTIES,             OLECMDID_PROPERTIES,
    IDM_CUT,                    OLECMDID_CUT,
    IDM_COPY,                   OLECMDID_COPY,
    IDM_PASTE,                  OLECMDID_PASTE,
    IDM_PASTESPECIAL,           OLECMDID_PASTESPECIAL,
    IDM_UNDO,                   OLECMDID_UNDO,
    IDM_REDO,                   OLECMDID_REDO,
    IDM_SELECTALL,              OLECMDID_SELECTALL,
    IDM_CLEARSELECTION,         OLECMDID_CLEARSELECTION,
    IDM_STOP,                   OLECMDID_STOP,
    IDM_REFRESH,                OLECMDID_REFRESH,
    IDM_STOPDOWNLOAD,           OLECMDID_STOPDOWNLOAD,
    IDM_ENABLE_INTERACTION,     OLECMDID_ENABLE_INTERACTION,
    OLECMDID_ONUNLOAD,          OLECMDID_ONUNLOAD,
    OLECMDID_DONTDOWNLOADCSS,   OLECMDID_DONTDOWNLOADCSS,
};

//+-------------------------------------------------------------------------
//
//  Method:     CBase::IDMFromCmdID, static
//
//  Synopsis:   Compute menu item identifier from command set and command id.
//
//--------------------------------------------------------------------------
int CBase::IDMFromCmdID(const GUID* pguidCmdGroup, ULONG ulCmdID)
{
    MAP* pmap;
    int cmap;

    if(pguidCmdGroup == NULL)
    {
        pmap = amapCommandSet95;
        cmap = ARRAYSIZE(amapCommandSet95);
    }
    else if(*pguidCmdGroup == CGID_MSHTML)
    {
        // Command identifiers in the Forms3 command set map
        // directly to menu item identifiers.
        return ulCmdID;
    }
    else
    {
        return IDM_UNKNOWN;
    }

    for(; --cmap>=0; pmap++)
    {
        if(pmap->usCmdID == ulCmdID)
        {
            return pmap->idm;
        }
    }

    return IDM_UNKNOWN;
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::IsCmdGroupSupported, static
//
//  Synopsis:   Determine if the given command group is supported.
//
//--------------------------------------------------------------------------
BOOL CBase::IsCmdGroupSupported(const GUID* pguidCmdGroup)
{
    return pguidCmdGroup==NULL || *pguidCmdGroup==CGID_MSHTML;
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::FireCancelableEvent
//
//  Synopsis:   Tries to fire the event and returns the script return value that
//              is a combination of returned value and varReturnValue (if present)
//              If one of them is TRUE (that means no default action for onerrer)
//              TRUE is returned. If no value is returned from the script the
//               *pfRet is not changed.
//--------------------------------------------------------------------------
HRESULT CBase::FireCancelableEvent(DISPID dispidMethod, DISPID dispidProp,
                                   IDispatch* pEventObject, VARIANT_BOOL* pfRetVal,
                                   BYTE* pbTypes, ...)
{
    va_list  valParms;
    CVariant Var;
    HRESULT  hr;

    va_start(valParms, pbTypes);
    hr = FireEventV(dispidMethod, dispidProp, pEventObject, &Var, pbTypes, valParms);
    va_end(valParms);

    if(pfRetVal != NULL)
    {
        if(V_VT(&Var) == VT_BOOL)
        {
            *pfRetVal = V_BOOL(&Var);
        }
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandSupported
//
//  Synopsis:
//
//  Returns: returns true if given command (like bold) is supported
//----------------------------------------------------------------------------
HRESULT CBase::queryCommandSupported(const BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    HRESULT hr = S_OK;
    ULONG   uCmdId;

    if(pfRet == NULL)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = CmdIDFromCmdName(bstrCmdId, &uCmdId);
    if(hr == S_OK)
    {
        *pfRet = VB_TRUE;
    }
    else if(hr == E_INVALIDARG)
    {
        *pfRet = VB_FALSE;
        hr = S_OK;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandEnabled
//
//  Synopsis:
//
//  Returns: returns true if given command is currently enabled. For toolbar
//          buttons not being enabled means being grayed.
//----------------------------------------------------------------------------
HRESULT CBase::queryCommandEnabled(const BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    HRESULT hr = S_OK;
    DWORD   dwFlags;

    if(pfRet == NULL)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *pfRet = VB_FALSE;

    hr = QueryCommandHelper(bstrCmdId, &dwFlags, NULL);
    if(hr)
    {
        goto Cleanup;
    }

    if(dwFlags==MSOCMDSTATE_NINCHED || dwFlags==MSOCMDSTATE_UP || dwFlags==MSOCMDSTATE_DOWN)
    {
        *pfRet = VB_TRUE;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandState
//
//  Synopsis:
//
//  Returns: returns true if given command is on. For toolbar buttons this
//          means being down. Note that a command button can be disabled
//          and also be down.
//----------------------------------------------------------------------------
HRESULT CBase::queryCommandState(const BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    HRESULT hr = S_OK;
    DWORD   dwFlags;

    if(pfRet == NULL)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *pfRet = VB_FALSE;

    hr = QueryCommandHelper(bstrCmdId, &dwFlags, NULL);
    if(hr)
    {
        goto Cleanup;
    }

    if(dwFlags == MSOCMDSTATE_DOWN)
    {
        *pfRet = VB_TRUE;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandIndeterm
//
//  Synopsis:
//
//  Returns: returns true if given command is in indetermined state.
//          If this value is TRUE the value returnd by queryCommandState
//          should be ignored.
//----------------------------------------------------------------------------
HRESULT CBase::queryCommandIndeterm(const BSTR bstrCmdId, VARIANT_BOOL* pfRet)
{
    HRESULT hr = S_OK;
    DWORD   dwFlags;

    if(pfRet == NULL)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *pfRet = VB_FALSE;

    hr = QueryCommandHelper(bstrCmdId, &dwFlags, NULL);
    if(hr)
    {
        goto Cleanup;
    }

    if(dwFlags == MSOCMDSTATE_NINCHED)
    {
        *pfRet = VB_TRUE;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandText
//
//  Synopsis:
//
//  Returns: Returns the text that describes the command (eg bold)
//----------------------------------------------------------------------------
HRESULT CBase::queryCommandText(const BSTR bstrCmdId, BSTR* pcmdText)
{
    HRESULT hr = S_OK;

    if(pcmdText == NULL)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *pcmdText = NULL;

    hr = QueryCommandHelper(bstrCmdId, NULL, pcmdText);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandValue
//
//  Synopsis:
//
//  Returns: Returns the  command value like font name or size.
//----------------------------------------------------------------------------
HRESULT CBase::queryCommandValue(const BSTR bstrCmdId, VARIANT* pvarRet)
{
    HRESULT hr = S_OK;
    DWORD   dwCmdId;

    if(pvarRet == NULL)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = VariantClear(pvarRet);
    if(hr)
    {
        goto Cleanup;
    }

    // Convert the command ID from string to number
    hr = CmdIDFromCmdName(bstrCmdId, &dwCmdId);
    if(hr)
    {
        goto Cleanup;
    }

    // Set the appropriate variant type
    V_VT(pvarRet) = GetExpectedCmdValueType(dwCmdId);

    // Call QueryStatus instead of exec if the expected return value is boolean
    if(V_VT(pvarRet) == VT_BOOL)
    {
        MSOCMD msocmd;

        msocmd.cmdID = dwCmdId;
        msocmd.cmdf  = 0;

        hr = QueryStatus(const_cast<GUID*>(&CGID_MSHTML), 1, &msocmd, NULL);
        if(hr)
        {
            goto Cleanup;
        }

        V_BOOL(pvarRet) = (msocmd.cmdf==MSOCMDSTATE_NINCHED || msocmd.cmdf==MSOCMDSTATE_DOWN)
            ? VB_TRUE : VB_FALSE;
    }
    else
    {
        // Use exec to get the string on integer value

        // If GetExpectedCmdValueType returned a VT_BSTR we need to null out the value
        // VariantClear wont do that.  If the pvarRet passed in the VariantClear
        // the bstrVal would be bogus.
        V_BSTR(pvarRet) = NULL;

        hr = Exec(const_cast<GUID*>(&CGID_MSHTML), dwCmdId, MSOCMDEXECOPT_DONTPROMPTUSER, NULL, pvarRet);
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
//  Member:     execCommand
//
//  Synopsis:   Executes given command
//
//  Returns:
//----------------------------------------------------------------------------
HRESULT CBase::execCommand(const BSTR bstrCmdId, VARIANT_BOOL showUI, VARIANT value)
{
    DWORD    dwCmdOpt;
    DWORD    dwCmdId;
    VARIANT* pValue = NULL;
    HRESULT  hr;

    // Translate the "show UI" flag into appropriate option
    dwCmdOpt = (showUI==VB_FALSE) ? MSOCMDEXECOPT_DONTPROMPTUSER : MSOCMDEXECOPT_PROMPTUSER;

    // Convert the command ID from string to number
    hr = CmdIDFromCmdName(bstrCmdId, &dwCmdId);
    if(hr)
    {
        goto Cleanup;
    }

    pValue = (V_VT(&value)==(VT_BYREF|VT_VARIANT)) ? V_VARIANTREF(&value) : &value;

    // Some functions do not check for empty or error type variants
    if(V_VT(pValue)==VT_ERROR || V_VT(pValue)==VT_EMPTY)
    {
        pValue = NULL;
    }

    hr = Exec(const_cast<GUID*>(&CGID_MSHTML), dwCmdId, dwCmdOpt, pValue, NULL);

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     execCommandShowHelp
//
//  Synopsis:
//
//  Returns:
//----------------------------------------------------------------------------
HRESULT CBase::execCommandShowHelp(const BSTR bstrCmdId)
{
    HRESULT hr;
    DWORD   dwCmdId;

    // Convert the command ID from string to number
    hr = CmdIDFromCmdName(bstrCmdId, &dwCmdId);
    if(hr)
    {
        goto Cleanup;
    }

    hr = Exec(const_cast<GUID*>(&CGID_MSHTML), dwCmdId, MSOCMDEXECOPT_SHOWHELP, NULL, NULL);

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::ExecSetGetProperty and 
//              CBase::ExecSetGetKnownProp
//
//  Synopsis:   Helper functions for Exec(), it Set/Get a property for
//              control. pvarargIn, pvarargOut can both not NULL. they both
//              call CBase:: ExecSetGetHelper for the invoke logic.
//
//--------------------------------------------------------------------------
HRESULT CBase::ExecSetGetProperty(
          VARIANTARG*   pvarargIn,  // In parameter
          VARIANTARG*   pvarargOut, // Out parameter
          UINT          uPropName,  // property name
          VARTYPE       vt)         // Parameter type
{
    HRESULT     hr = S_OK;
    IDispatch*  pDispatch = NULL;
    DISPID      dispid;

    // Invalid parameters
    if(!pvarargIn && !pvarargOut)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // try to turn the uPropName into a dispid for the property
    hr = GetDispatchForProp(uPropName, &pDispatch, &dispid);
    if(hr)
    {
        goto Cleanup;
    }

    // call the helper to do the work
    hr = ExecSetGetHelper(pvarargIn, pvarargOut, pDispatch, dispid, vt);

Cleanup:
    ReleaseInterface(pDispatch);
    if(DISP_E_UNKNOWNNAME == hr)
    {
        hr = OLECMDERR_E_NOTSUPPORTED; // we listen for this error code
    }
    RRETURN(hr);
}

HRESULT CBase::ExecSetGetKnownProp(
           VARIANTARG*  pvarargIn,  // In parameter
           VARIANTARG*  pvarargOut, // Out parameter
           DISPID       dispidProp, 
           VARTYPE      vt)         // Parameter type
{
    HRESULT hr = S_OK;
    IDispatch* pDispatch = NULL;

    // Invalid parameters?
    if(!pvarargIn && !pvarargOut)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // now get the IDispatch for *this*
    hr = PrivateQueryInterface(IID_IDispatch, (void**)&pDispatch);
    if(hr)
    {
        goto Cleanup;
    }

    // now call the helper to do the work
    hr = ExecSetGetHelper(pvarargIn, pvarargOut, pDispatch, dispidProp, vt);

Cleanup:
    ReleaseInterface(pDispatch);
    RRETURN(hr);
}

HRESULT CBase::ExecSetGetHelper(
        VARIANTARG* pvarargIn,  // In parameter
        VARIANTARG* pvarargOut, // Out parameter
        IDispatch*  pDispatch,  // the IDispatch for *this*
        DISPID      dispid,     // the property dispid
        VARTYPE     vt)         // Parameter type
{
    HRESULT     hr  =S_OK;
    DISPPARAMS  dp = _afxGlobalData._Zero.dispparams;   // initialized be zero.
    DISPID      dispidPut = DISPID_PROPERTYPUT;         // Dispid of prop arg.

    // Set property
    if(pvarargIn)
    {
        // Fill in dp
        dp.rgvarg = pvarargIn;
        dp.rgdispidNamedArgs = &dispidPut;
        dp.cArgs = 1;
        dp.cNamedArgs = 1;

        hr = pDispatch->Invoke(
            dispid,
            IID_NULL,
            NULL,
            DISPATCH_PROPERTYPUT,
            &dp,
            NULL,
            NULL,
            NULL);

        if(hr)
        {
            goto Cleanup;
        }
    }

    // Get property
    if(pvarargOut)
    {
        // Get property requires different dp
        dp = _afxGlobalData._Zero.dispparams;

        hr = pDispatch->Invoke(
            dispid,
            IID_NULL,
            NULL,
            DISPATCH_PROPERTYGET,
            &dp,
            pvarargOut,
            NULL,
            NULL);

        if(hr)
        {
            // BUGBUG (EricVas) - This is a hack to prevent the member
            // not found from trickling up, causing a nasty message box to
            // appear
            hr = OLECMDERR_E_DISABLED;

            goto Cleanup;
        }

        // Update the VT if necessary
        V_VT(pvarargOut) = vt;
    }

Cleanup:
    if(DISP_E_UNKNOWNNAME == hr)
    {
        hr = OLECMDERR_E_NOTSUPPORTED; // we listen for this error code
    }
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::ExecSetPropertyCmd
//
//  Synopsis:   Helper function for Exec(), It is used for SpecialEffect
//              Commands, Justify (TextAlign). For these commands, there are
//              not input parameter.
//
//  Arguments:
//
//  Returns:
//
//--------------------------------------------------------------------------
HRESULT CBase::ExecSetPropertyCmd(UINT uPropName, DWORD value)
{
    VARIANT var;

    VariantInit(&var);
    V_VT(&var) = VT_I4;
    V_I4(&var) = value;

    return ExecSetGetProperty(&var, NULL, uPropName, VT_I4);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::ExecToggleCmd
//
//  Synopsis:   Helper function for exec(). It is used for cmdidBold,
//              cmdidItalic, cmdidUnderline. It always toggle property
//              value.
//
//  Arguments:
//
//  Returns:
//
//--------------------------------------------------------------------------
HRESULT CBase::ExecToggleCmd(UINT uPropName) // Dipatch ID for command
{
    HRESULT     hr ;
    IDispatch*  pDispatch = NULL;
    DISPPARAMS  dp = _afxGlobalData._Zero.dispparams;   // initialized be zero.
    DISPID      dispidPut = DISPID_PROPERTYPUT;         // Dispid of prop arg.
    VARIANT     var;
    DISPID      dispid;

    hr = GetDispatchForProp(uPropName, &pDispatch, &dispid);
    if(hr)
    {
        goto Cleanup;
    }

    VariantInit(&var);
    V_VT(&var) = VT_BOOL;

    // Get property value
    hr = pDispatch->Invoke(
        dispid,
        IID_NULL,
        NULL,
        DISPATCH_PROPERTYGET,
        &dp,
        &var,
        NULL,
        NULL);
    if(hr)
    {
        goto Cleanup;
    }

    // Toggle property value
    V_BOOL(&var) = !V_BOOL(&var);

    // Fill in dp
    dp.rgvarg = &var;
    dp.rgdispidNamedArgs = &dispidPut;
    dp.cArgs = 1;
    dp.cNamedArgs = 1;

    // Set new property value
    hr = pDispatch->Invoke(
        dispid,
        IID_NULL,
        NULL,
        DISPATCH_PROPERTYPUT,
        &dp,
        NULL,
        NULL,
        NULL);

Cleanup:
    ReleaseInterface(pDispatch);
    if(DISP_E_UNKNOWNNAME == hr)
    {
        hr = OLECMDERR_E_NOTSUPPORTED; // we listen for this error code
    }
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::QueryStatusProperty
//
//  Synopsis:   Helper function for QueryStatus(), it determines if a control
//              supports a property by checking whether you can get property. .
//
//  Arguments:
//
//  Returns:
//
//--------------------------------------------------------------------------
HRESULT CBase::QueryStatusProperty(MSOCMD* pCmd, UINT uPropName, VARTYPE vt)
{
    HRESULT     hr;
    IDispatch*  pDispatch = NULL;
    CVariant    var;
    DISPPARAMS  dp = _afxGlobalData._Zero.dispparams;         // initialized be zero.
    DISPID      dispid;

    hr = GetDispatchForProp(uPropName, &pDispatch, &dispid);
    if(hr)
    {
        goto Cleanup;
    }

    V_VT(&var) = vt;

    hr = pDispatch->Invoke(
        dispid,
        IID_NULL,
        NULL,
        DISPATCH_PROPERTYGET,
        &dp,
        &var,
        NULL,
        NULL);

Cleanup:
    if(!hr)
    {
        if(V_VT(&var)==VT_BOOL && V_BOOL(&var)==VB_TRUE)
        {
            pCmd->cmdf = MSOCMDSTATE_DOWN;
        }
        else
        {
            pCmd->cmdf = MSOCMDSTATE_UP;
        }
    }
    if(DISP_E_UNKNOWNNAME == hr)
    {
        hr = OLECMDERR_E_NOTSUPPORTED; // we listen for this error code
    }
    ReleaseInterface(pDispatch);
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::GetDispatchForProp
//
//  Synopsis:   Helper function for flavours of Exec()-s and QueryStatus()-es
//
//--------------------------------------------------------------------------
HRESULT CBase::GetDispatchForProp(UINT uPropName, IDispatch** ppDisp, DISPID* pdispid)
{
    HRESULT hr;
    TCHAR   achPropName[64];
    LPTSTR  pchPropName = achPropName;
    int     nLoadStringRes;

    *ppDisp = NULL;

    hr = PrivateQueryInterface(IID_IDispatch, (void**)ppDisp);
    if(hr)
    {
        goto Cleanup;
    }

    nLoadStringRes = LoadString(_afxGlobalData._hInstResource, uPropName, achPropName, ARRAYSIZE(achPropName));
    if(0 == nLoadStringRes)
    {
        hr = GetLastWin32Error();
        Assert (!OK(hr));
        goto Cleanup;
    }

    hr = (*ppDisp)->GetIDsOfNames(IID_NULL, &pchPropName, 1, _afxGlobalData._lcidUserDefault, pdispid);

Cleanup:
    if(hr)
    {
        ClearInterface(ppDisp);
    }

    RRETURN(hr);
}

// Table used to to translate command names to command IDs and get command value return types
// Notice that the last element defines the default cmdID and vt returned when the item is not found
CBase::CMDINFOSTRUCT CBase::cmdTable[] =
{
    _T("CreateBookmark"),       IDM_BOOKMARK,
    _T("CreateLink"),           IDM_HYPERLINK,
    _T("InsertImage"),          IDM_IMAGE,
    _T("Bold"),                 IDM_BOLD,
    _T("BrowseMode"),           IDM_BROWSEMODE,
    _T("EditMode"),             IDM_EDITMODE,
    _T("InsertButton"),         IDM_BUTTON,
    _T("InsertIFrame"),         IDM_IFRAME,
    _T("InsertInputButton"),    IDM_INSINPUTBUTTON,
    _T("InsertInputCheckbox"),  IDM_CHECKBOX,
    _T("InsertInputImage"),     IDM_INSINPUTIMAGE,
    _T("InsertInputRadio"),     IDM_RADIOBUTTON,
    _T("InsertInputText"),      IDM_TEXTBOX,
    _T("InsertSelectDropdown"), IDM_DROPDOWNBOX,
    _T("InsertSelectListbox"),  IDM_LISTBOX,
    _T("InsertTextArea"),       IDM_TEXTAREA,
    _T("Italic"),               IDM_ITALIC,
    _T("SizeToControl"),        IDM_SIZETOCONTROL,
    _T("SizeToControlHeight"),  IDM_SIZETOCONTROLHEIGHT,
    _T("SizeToControlWidth"),   IDM_SIZETOCONTROLWIDTH,
    _T("Underline"),            IDM_UNDERLINE,
    _T("Copy"),                 IDM_COPY,
    _T("Cut"),                  IDM_CUT,
    _T("Delete"),               IDM_DELETE,
    _T("Print"),                IDM_EXECPRINT,
    _T("JustifyCenter"),        IDM_JUSTIFYCENTER,
    _T("JustifyFull"),          IDM_JUSTIFYFULL,
    _T("JustifyLeft"),          IDM_JUSTIFYLEFT,
    _T("JustifyRight"),         IDM_JUSTIFYRIGHT,
    _T("JustifyNone"),          IDM_JUSTIFYNONE,
    _T("Paste"),                IDM_PASTE,
    _T("PlayImage"),            IDM_DYNSRCPLAY,
    _T("StopImage"),            IDM_DYNSRCSTOP,
    _T("InsertInputReset"),     IDM_INSINPUTRESET,
    _T("InsertInputSubmit"),    IDM_INSINPUTSUBMIT,
    _T("InsertInputFileUpload"),IDM_INSINPUTUPLOAD,
    _T("InsertFieldset"),       IDM_INSFIELDSET,
    _T("Unselect"),             IDM_CLEARSELECTION,
    _T("BackColor"),            IDM_BACKCOLOR,
    _T("ForeColor"),            IDM_FORECOLOR,
    _T("FontName"),             IDM_FONTNAME,
    _T("FontSize"),             IDM_FONTSIZE,
    _T("FormatBlock"),          IDM_BLOCKFMT,
    _T("Indent"),               IDM_INDENT,
    _T("InsertMarquee"),        IDM_MARQUEE,
    _T("InsertOrderedList"),    IDM_ORDERLIST,
    _T("InsertParagraph"),      IDM_PARAGRAPH,
    _T("InsertUnorderedList"),  IDM_UNORDERLIST,
    _T("Outdent"),              IDM_OUTDENT,
    _T("Redo"),                 IDM_REDO,
    _T("Refresh"),              IDM_REFRESH,
    _T("RemoveParaFormat"),     IDM_REMOVEPARAFORMAT,
    _T("RemoveFormat"),         IDM_REMOVEFORMAT,
    _T("SelectAll"),            IDM_SELECTALL,
    _T("StrikeThrough"),        IDM_STRIKETHROUGH,
    _T("Subscript"),            IDM_SUBSCRIPT,            
    _T("Superscript"),          IDM_SUPERSCRIPT,
    _T("Undo"),                 IDM_UNDO,
    _T("Unlink"),               IDM_UNLINK,
    _T("InsertHorizontalRule"), IDM_HORIZONTALLINE,
    _T("UnBookmark"),           IDM_UNBOOKMARK,
    _T("OverWrite"),            IDM_OVERWRITE,
    _T("InsertInputPassword"),  IDM_INSINPUTPASSWORD,
    _T("InsertInputHidden"),    IDM_INSINPUTHIDDEN,
    _T("DirLTR"),               IDM_DIRLTR,
    _T("DirRTL"),               IDM_DIRRTL,
    _T("BlockDirLTR"),          IDM_BLOCKDIRLTR,
    _T("BlockDirRTL"),          IDM_BLOCKDIRRTL,
    _T("InlineDirLTR"),         IDM_INLINEDIRLTR,
    _T("InlineDirRTL"),         IDM_INLINEDIRRTL,
    _T("SaveAs"),               IDM_SAVEAS,
    _T("Open"),                 IDM_OPEN,
    _T("Stop"),                 IDM_STOP,
    NULL,                       0
};

// Translates command name into command ID. If the command is not found and the
//  command string starts with a digit the command number is used.
HRESULT CBase::CmdIDFromCmdName(BSTR bstrCmdName, ULONG* pcmdValue)
{
    int i;
    HRESULT hr=S_OK;

    Assert(pcmdValue != NULL);
    *pcmdValue = 0;

    if(FormsIsEmptyString(bstrCmdName))
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    i = 0;
    while(cmdTable[i].cmdName != NULL)
    {
        if(StrCmpIC(cmdTable[i].cmdName, bstrCmdName) == 0)
        {
            break;
        }
        i++;
    }
    if(cmdTable[i].cmdName != NULL)
    {
        // The command name was found, use the value from the table
        *pcmdValue = cmdTable[i].cmdID;
        if(*pcmdValue == 0)
        {
            hr = E_INVALIDARG;
        }
    }
    else
    {
        hr = E_INVALIDARG;
    }

Cleanup:
    return hr;
}

// Returns the expected VARIANT type of the command value (like VT_BSTR for font name)
VARTYPE CBase::GetExpectedCmdValueType(ULONG uCmdID)
{
    // We do not need to set the pvarOut, except for IDM_FONTSIZE. But we still
    // use this logic to determine if we need to use the queryStatus or exec
    // to get the command value
    if(uCmdID==IDM_FONTSIZE || uCmdID==IDM_FORECOLOR || uCmdID==IDM_BACKCOLOR)
    {
        return VT_I4;
    }

    if(uCmdID==IDM_BLOCKFMT || uCmdID==IDM_FONTNAME) 
    {
        return VT_BSTR;
    }

    return VT_BOOL;
}

//+---------------------------------------------------------------------------
//
//  Member:     queryCommandHelper
//
//  Synopsis:   This function is called by QueryCommandXXX functions and does the
//              command ID conversions and returns the flags or the text for given
//              command. Only one of the return parameters must be not NULL.
//
//  Returns:    S_OK if the value is returned
//----------------------------------------------------------------------------
struct MSOCMDTEXT_WITH_TEXT
{
    MSOCMDTEXT  header;
    WCHAR       text[FORMS_BUFLEN];
};

HRESULT CBase::QueryCommandHelper(const BSTR bstrCmdId, DWORD* cmdf, BSTR* pcmdText)
{
    HRESULT                 hr = S_OK;
    MSOCMD                  msocmd;
    MSOCMDTEXT_WITH_TEXT    msocmdtext;

    Assert((cmdf==NULL && pcmdText!=NULL) || (cmdf!=NULL && pcmdText==NULL));

    // initialize the values so in case of error we return a NULL pointer
    if(pcmdText != NULL)
    {
        *pcmdText = NULL;
    }
    else
    {
        *cmdf = 0L;
    }

    // Fill the command structure converting the command ID from string to number
    hr = CmdIDFromCmdName(bstrCmdId, &(msocmd.cmdID));
    if(hr)
    {
        goto Cleanup;
    }
    msocmd.cmdf  = 0;

    if(pcmdText != NULL)
    {
        msocmdtext.header.cmdtextf = MSOCMDTEXTF_NAME;
        msocmdtext.header.cwBuf    = FORMS_BUFLEN;
        msocmdtext.header.cwActual = 0;
    }

    hr = QueryStatus(const_cast<GUID*>(&CGID_MSHTML), 1, &msocmd, (MSOCMDTEXT*)&msocmdtext);
    if(hr)
    {
        goto Cleanup;
    }

    if(pcmdText != NULL)
    {
        // Ignore the  msocmd value, just return the text
        if(msocmdtext.header.cwActual > 0)
        {
            // Allocate the return string
            *pcmdText = SysAllocString(msocmdtext.header.rgwz);
            if(*pcmdText == NULL)
            {
                hr = E_OUTOFMEMORY;
                goto Cleanup;
            }
        }
    }
    else
    {
        // return the flags
        *cmdf = msocmd.cmdf;
    }

Cleanup:
    RRETURN(hr);
}

HRESULT CBase::toString(BSTR* bstrString)
{
    HRESULT    hr;
    CVariant   var;

    if(!bstrString)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *bstrString = NULL;

    hr = var.CoerceVariantArg(VT_BSTR);
    if(hr)
    {
        goto Cleanup;
    }

    *bstrString = V_BSTR(&var);
    // Don't allow deletion of the BSTR
    V_VT(&var) = VT_EMPTY;

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------
//
//  Member:     CBase::SubRelease
//
//---------------------------------------------------------------
ULONG CBase::SubRelease()
{
#ifdef _DEBUG
    ULONG ulAllRefs = _ulAllRefs;
#endif
    if(--_ulAllRefs == 0)
    {
        _ulAllRefs = ULREF_IN_DESTRUCTOR;
        _ulRefs = ULREF_IN_DESTRUCTOR;
        delete this;
    }
#ifdef _DEBUG
    return ulAllRefs-1;
#else
    return 0;
#endif
}

//+---------------------------------------------------------------
//
//  Member:     CBase::InvokeDispatchWithThis
//
//  Synopsis:   invokes pDisp adding IUnknown of this
//              named parameter to args list
//
//---------------------------------------------------------------
HRESULT CBase::InvokeDispatchWithThis(IDispatch* pDisp, VARIANT* pExtraArg,
                                      REFIID riid, LCID lcid, WORD wFlags,
                                      DISPPARAMS* pdispparams, VARIANT* pvarResult,
                                      EXCEPINFO* pexcepinfo, IServiceProvider* pSrvProvider) 
{
    HRESULT         hr;
    IDispatchEx*    pDEX = NULL;
    DISPPARAMS*     pPassedParams = pdispparams;
    IUnknown*       pUnk = NULL;
    CDispParams     dispParams((pExtraArg?pdispparams->cArgs+1:pdispparams->cArgs), pdispparams->cNamedArgs+1);

    hr = pDisp->QueryInterface(IID_IDispatchEx, (void**)&pDEX);
    // Does the object support IDispatchEx then we'll need to pass
    // the this pointer of this object
    if(!hr)
    {
        Verify(!PrivateQueryInterface(IID_IDispatch, (void**)&pUnk));

        hr = dispParams.Create(pdispparams);
        if(hr)
        {
            goto Cleanup;
        }

        // Create the named this parameter
        Assert(dispParams.cNamedArgs > 0);
        V_VT(&dispParams.rgvarg[0]) = VT_DISPATCH;
        V_UNKNOWN(&dispParams.rgvarg[0]) = pUnk;
        dispParams.rgdispidNamedArgs[0] = DISPID_THIS;

        if(pExtraArg)
        {
            UINT uIdx = dispParams.cArgs - 1;

            memcpy(&(dispParams.rgvarg[uIdx]), pExtraArg, sizeof(VARIANT));
        }

        pPassedParams = &dispParams;
    }

    // Call the function
    hr = pDEX
        ? pDEX->InvokeEx(
        DISPID_VALUE,
        lcid,
        wFlags,
        pPassedParams,
        pvarResult,
        pexcepinfo,
        pSrvProvider)
        : pDisp->Invoke(
        DISPID_VALUE,
        riid,
        lcid,
        wFlags,
        pPassedParams,
        pvarResult,
        pexcepinfo,
        NULL);

    // If we had to pass the this parameter and we had more than
    // 1 parameter then copy the orginal back over (the args could have
    // been byref).
    if(pPassedParams!=pdispparams && pPassedParams->cArgs>1)
    {
        hr = dispParams.MoveArgsToDispParams(pdispparams, pdispparams->cArgs);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    ReleaseInterface(pDEX);
    ReleaseInterface(pUnk);

    RRETURN (hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::InvokeAA
//
//  Synopsis:   helper
//
//--------------------------------------------------------------------------
HRESULT CBase::InvokeAA(
                        DISPID              dispidMember,
                        CAttrValue::AATYPE  aaType,
                        LCID                lcid,
                        WORD                wFlags,
                        DISPPARAMS*         pdispparams,
                        VARIANT*            pvarResult,
                        EXCEPINFO*          pexcepinfo,
                        IServiceProvider*   pSrvProvider)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    LONG    cArgsActual = GetArgsActual(pdispparams);

    if((cArgsActual==0) && (wFlags&DISPATCH_PROPERTYGET))
    {
        if(pvarResult)
        {
            hr = GetVariantAt(FindAAIndex(dispidMember, aaType),
                pvarResult, FALSE);
            if(!hr)
            {
                goto Cleanup;
            }
        }
    }
    else if(wFlags & DISPATCH_PROPERTYPUT)
    {
        if(pdispparams && cArgsActual==1)
        {
            AAINDEX aaIdx = FindAAIndex(dispidMember, aaType);
            if(aaIdx == AA_IDX_UNKNOWN)
            {
                hr = AddVariant(dispidMember, pdispparams->rgvarg, aaType);
            }
            else
            {
                hr = ChangeVariantAt(aaIdx, pdispparams->rgvarg);
            }
            if(hr)
            {
                goto Cleanup;
            }
            hr = OnPropertyChange(dispidMember, 0);
            goto Cleanup;
        }
        else
        {
            // Missing value to set.
            hr = DISP_E_BADPARAMCOUNT;
            goto Cleanup;
        }
    }

    // If the wFlags was marked as a Get/Method and the get failed then try
    // this as a method invoke.  Of if the wFlags was only a method invoke
    // then the hr is set to DISP_E_MEMBERNOTFOUND at the entry ExpandoInvoke.
    if(hr && (wFlags&DISPATCH_METHOD))
    {
        AAINDEX aaidx = FindAAIndex(dispidMember, aaType);
        if(AA_IDX_UNKNOWN == aaidx)
        {
            if(pvarResult)
            {
                VariantClear(pvarResult);
                pvarResult->vt = VT_NULL;
                hr = S_OK;
            }
        }
        else
        {
            IDispatch* pDisp = NULL;

            hr = GetDispatchObjectAt(aaidx, &pDisp);
            if(!hr && pDisp)
            {
                hr = InvokeDispatchWithThis(
                    pDisp,
                    NULL,
                    IID_NULL,
                    lcid,
                    wFlags,
                    pdispparams,
                    pvarResult,
                    pexcepinfo,
                    pSrvProvider);
                ReleaseInterface(pDisp);
            }
        }
    }
    else
    {
        hr = DISP_E_MEMBERNOTFOUND;
    }

Cleanup:
    RRETURN (hr);
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::SetErrorInfo
//
//  Synopsis:
//
//
//----------------------------------------------------------------------------
HRESULT CBase::SetErrorInfo(HRESULT hr)
{
    PreSetErrorInfo();

    if(FAILED(hr))
    {
        ClearErrorInfo();
        CloseErrorInfo(hr);
    }

    return hr;
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::SetErrorInfoPGet
//
//----------------------------------------------------------------------------
HRESULT CBase::SetErrorInfoPGet(HRESULT hr, DISPID dispid)
{
    // No PreSetErrorInfo call needed on read-only operations.
    if(FAILED(hr))
    {
        ClearErrorInfo();
        CloseErrorInfoPGet(hr, dispid);
    }

    return hr;
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::SetErrorInfoPSet
//
//----------------------------------------------------------------------------
HRESULT CBase::SetErrorInfoPSet(HRESULT hr, DISPID dispid)
{
    PreSetErrorInfo();

    if(FAILED(hr))
    {
        if(hr == E_INVALIDARG)
        {
            hr = CTL_E_INVALIDPROPERTYVALUE;
        }
        ClearErrorInfo();
        CloseErrorInfoPSet(hr, dispid);
    }

    return hr;
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::SetErrorInfoInvalidArg
//
//----------------------------------------------------------------------------
HRESULT CBase::SetErrorInfoInvalidArg()
{
    PreSetErrorInfo();

    ClearErrorInfo();
    return CloseErrorInfo(E_INVALIDARG);
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::SetErrorInfoBadValue
//
//----------------------------------------------------------------------------
HRESULT CBase::SetErrorInfoBadValue(DISPID dispid, UINT ids, ...)
{
    PreSetErrorInfo();

    va_list arg;
    CErrorInfo* pEI;

    ClearErrorInfo();

    if(ids && (pEI=GetErrorInfo())!=NULL)
    {
        va_start(arg, ids);
        pEI->SetTextV(EPART_SOLUTION, ids, &arg);
        va_end(arg);
    }

    return CloseErrorInfoPSet(E_INVALIDARG, dispid);
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::SetErrorInfoPBadValue
//
//----------------------------------------------------------------------------
HRESULT CBase::SetErrorInfoPBadValue(DISPID dispid, UINT ids, ...)
{
    PreSetErrorInfo();

    va_list arg;
    CErrorInfo* pEI;

    ClearErrorInfo();

    if(ids && (pEI=GetErrorInfo())!=NULL)
    {
        va_start(arg, ids);
        pEI->SetTextV(EPART_SOLUTION, ids, &arg);
        va_end(arg);
    }

    return CloseErrorInfoPSet(CTL_E_INVALIDPROPERTYVALUE, dispid);
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::CloseErrorInfo
//
//  Specific method for automation calls
//
//----------------------------------------------------------------------------
HRESULT CBase::CloseErrorInfo(HRESULT hr, DISPID dispid, INVOKEKIND invkind)
{
    CErrorInfo* pEI;

    if(OK(hr))
    {
        return hr;
    }

    if((pEI=GetErrorInfo()) != NULL)
    {
        pEI->_invkind = invkind;
        pEI->_dispidInvoke = dispid;
        if(BaseDesc()->_piidDispinterface)
        {
            pEI->_iidInvoke = *BaseDesc()->_piidDispinterface;
        }
        else
        {
            // BUGBUG rgardner If we don't have an IID what do we
            // do????
            return S_FALSE;
        }
    }

    return CloseErrorInfo(hr);
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::CloseErrorInfo
//
//  Synopsis:
//
//
//----------------------------------------------------------------------------
HRESULT CBase::CloseErrorInfo(HRESULT hr)
{
    if(FAILED(hr))
    {
        Assert(BaseDesc()->_pclsid);
        ::CloseErrorInfo(hr, (BaseDesc()->_pclsid?*BaseDesc()->_pclsid:CLSID_NULL));
    }

    return hr;
}

//+---------------------------------------------------------------------------
//
//  Member:     CBase::SetCodeProperty
//
//----------------------------------------------------------------------------
HRESULT CBase::SetCodeProperty(DISPID dispidCodeProp, IDispatch* pDispCode, BOOL* pfAnyDeleted/*=NULL*/)
{
    HRESULT hr = S_OK;
    BOOL fAnyDelete;

    fAnyDelete = DidFindAAIndexAndDelete(dispidCodeProp, CAttrValue::AA_Attribute);
    fAnyDelete |= DidFindAAIndexAndDelete(dispidCodeProp, CAttrValue::AA_Internal);

    if(pDispCode)
    {
        hr = AddDispatchObject(dispidCodeProp, pDispCode, CAttrValue::AA_Internal);
    }

    // if this is an element, let it know that events can now fire
    OnPropertyChange(DISPID_EVPROP_ONATTACHEVENT, 0);

    if(pfAnyDeleted)
    {
        *pfAnyDeleted = fAnyDelete;
    }

    RRETURN(hr);
}

BOOL CBase::removeAttributeDispid(DISPID dispid, const PROPERTYDESC* pPropDesc/*=NULL*/)
{
    AAINDEX     aaIx;
    BOOL        fExpando;
    CVariant    varNULL(VT_NULL);

    fExpando = IsExpandoDISPID(dispid);

    if(dispid != STDPROPID_XOBJ_STYLE)
    {
        aaIx = FindAAIndex(dispid, fExpando ?CAttrValue::AA_Expando:CAttrValue::AA_Attribute);
    }
    else
    {
        // We cannot delete the inline style object. It is created "on demand".
        // We need to remove it's attrarray that stored on the element.
        aaIx = FindAAIndex(DISPID_INTERNAL_INLINESTYLEAA, CAttrValue::AA_AttrArray);
    }

    if(aaIx == AA_IDX_UNKNOWN)
    {
        return FALSE;
    }

    // BUGBUG: handle expando undo (jbeda)
    /*if(!fExpando && QueryCreateUndo(TRUE, FALSE))
    {
        HRESULT     hr;
        CVariant    varOld;
        IDispatch*  pDisp = NULL;

        hr = PunkOuter()->QueryInterface(IID_IDispatch, (LPVOID*)&pDisp);
        if(hr)
        {
            goto UndoCleanup;
        }

        hr = GetDispProp(
            pDisp,
            dispid,
            _afxGlobalData._lcidUserDefault,
            &varOld);
        if(hr)
        {
            goto UndoCleanup;
        }

        hr = CreatePropChangeUndo(dispid, &varOld, NULL);

UndoCleanup:
        ReleaseInterface(pDisp);
    } wlw note*/

    DeleteAt(aaIx);

    if(!fExpando)
    {
        CLock Lock(this); // For OnPropertyChange

        if(!pPropDesc)
        {
        }

        // Need to fire a property change event
        if(OnPropertyChange(dispid, pPropDesc->GetdwFlags()))
        {
            return FALSE;
        }
    }

    return TRUE;
}

//+------------------------------------------------------------------------
//
//  Member:     CBase::QueryService
//
//  Synopsis:   Get service from host.  Derived classes should override.
//
//  Arguments:  guidService     id of service
//              iid             id of interface on service
//              ppv             the service
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CBase::QueryService(REFGUID guidService, REFIID iid, void** ppv)
{
    if(ppv)
    {
        *ppv = NULL;
    }
    RRETURN(E_NOINTERFACE);
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::DoAdvise
//
//  Synopsis:   Implements IConnectionPoint::Advise.
//
//  Arguments:  [ppUnkSinkOut]    -- Pointer to outgoing sink ptr.
//              [iid]             -- Interface of the advisee.
//              [fEvent]          -- Whether it's an event iface advise
//              [pUnkSinkAdvise]  -- Advisee's sink.
//              [pdwCookie]       -- Cookie to identify connection.
//
//----------------------------------------------------------------------------
HRESULT CBase::DoAdvise(REFIID iid, DISPID dispidBase, IUnknown* pUnkSinkAdvise,
                        IUnknown** ppUnkSinkOut, DWORD* pdwCookie)
{
    HRESULT                 hr;
    IUnknown*               pUnk = NULL;
    CAttrValue::AAExtraBits wAAExtra = CAttrValue::AA_Extra_Empty;
    DWORD                   dwCookie;

    Assert(iid != IID_NULL);

    if(!pUnkSinkAdvise)
    {
        RRETURN(E_INVALIDARG);
    }

    hr = pUnkSinkAdvise->QueryInterface(iid, (void**)&pUnk);
    if(hr)
    {
        if(dispidBase == DISPID_A_EVENTSINK)
        {
            hr = pUnkSinkAdvise->QueryInterface(IID_IDispatch, (void**)&pUnk);
        }

        if(hr)
        {
            RRETURN1(CONNECT_E_CANNOTCONNECT, CONNECT_E_CANNOTCONNECT);
        }
    }

    hr = AddUnknownObjectMultiple(dispidBase, pUnk, CAttrValue::AA_Internal, wAAExtra);
    if(hr)
    {
        goto Cleanup;
    }

    if(ppUnkSinkOut)
    {
        *ppUnkSinkOut = pUnk;
        pUnk->AddRef();
    }

#ifdef _WIN64
    AAINDEX aaidx;
    aaidx = AA_IDX_UNKNOWN;
    Verify(FindAdviseIndex(dispidBase, 0, pUnk, &aaidx) == S_OK); // We just added it, so this better not fail
    dwCookie = (DWORD)InterlockedIncrement((LONG*)&g_dwCookieForWin64);
    if(aaidx != AA_IDX_UNKNOWN)
    {
        Verify(SetCookieAt(aaidx, dwCookie) == S_OK);
    }
#else
    dwCookie = (DWORD)pUnk;
#endif

    if(pdwCookie)
    {
        *pdwCookie = dwCookie;
    }

Cleanup:
    ReleaseInterface(pUnk);
    RRETURN1(hr, CONNECT_E_ADVISELIMIT);
}

//+---------------------------------------------------------------------------
//
//  Method:     CBase::DoUnadvise
//
//  Synopsis:   Implements IConnectionPoint::Unadvise.
//
//  Arguments:  [dwCookie]        -- Cookie handed out upon Advise.
//
//----------------------------------------------------------------------------
HRESULT CBase::DoUnadvise(DWORD dwCookie, DISPID dispidBase)
{
    HRESULT hr;
    AAINDEX aaidx;

    if(!dwCookie)
    {
        return S_OK;
    }

#ifdef _WIN64
    hr = FindAdviseIndex(dispidBase, dwCookie, NULL, &aaidx);
#else
    hr = FindAdviseIndex(dispidBase, 0, (IUnknown*)dwCookie, &aaidx);
#endif
    if(hr)
    {
        goto Cleanup;
    }

    DeleteAt(aaidx);

Cleanup:
    RRETURN1(hr, CONNECT_E_NOCONNECTION);
}

HRESULT CBase::FindAdviseIndex(DISPID dispidBase, DWORD dwCookie, IUnknown* pUnkCookie, AAINDEX* paaidx)
{
    HRESULT     hr       = S_OK;
    IUnknown*   pUnkSink = NULL;
    AAINDEX     aaidx    = AA_IDX_UNKNOWN;

    if(!*GetAttrArray())
    {
        hr = CONNECT_E_NOCONNECTION;
        goto Cleanup;
    }

    // Loop through the attr array indices looking for the dwCookie and/or
    // pUnkCookie as the value.
    for( ; ; )
    {
        aaidx = FindNextAAIndex(dispidBase, CAttrValue::AA_Internal, aaidx);
        if(aaidx == AA_IDX_UNKNOWN)
        {
            break;
        }

        if(pUnkCookie)
        {
            ClearInterface(&pUnkSink);
            hr = GetUnknownObjectAt(aaidx, &pUnkSink);
            if(hr || !pUnkSink)
            {
                hr = CONNECT_E_NOCONNECTION;
                goto Cleanup;
            }

            if(pUnkSink == pUnkCookie)
            {
                break;
            }
        }

#ifdef _WIN64
        if(dwCookie)
        {
            DWORD dwCookieSink;
            hr = GetCookieAt(aaidx, &dwCookieSink);
            if(dwCookie == dwCookieSink)
            {
                break;
            }
        }
#endif
    }

    if(aaidx == AA_IDX_UNKNOWN)
    {
        hr = CONNECT_E_NOCONNECTION;
        goto Cleanup;
    }

    *paaidx = aaidx;

Cleanup:
    ReleaseInterface(pUnkSink);
    RRETURN1(hr, CONNECT_E_NOCONNECTION);
}

HRESULT CBase::FindEventName(ITypeInfo* pTISrc, DISPID dispid, BSTR* pBSTR)
{
    int         ievt;
    int         cbEvents;
    UINT        cbstr;
    HRESULT     hr;
    FUNCDESC*   pFDesc = NULL;
    TYPEATTR*   pTAttr = NULL;

    Assert(pTISrc);
    Assert(pBSTR);

    *pBSTR = NULL;

    hr = pTISrc->GetTypeAttr(&pTAttr);
    if(hr)
    {
        goto Cleanup;
    }

    if(pTAttr->typekind != TKIND_DISPATCH)
    {
        hr = E_NOINTERFACE;
        goto Cleanup;
    }

    // Total events.
    cbEvents = pTAttr->cFuncs;
    if(cbEvents == 0)
    {
        hr = E_NOINTERFACE;
        goto Cleanup;
    }

    // Populate the event table.
    for(ievt=0; ievt<cbEvents; ++ievt)
    {
        hr = pTISrc->GetFuncDesc(ievt, &pFDesc);
        if(hr)
        {
            goto Cleanup;
        }

        // Did we find the event that fired?
        if(dispid == pFDesc->memid)
        {
            // Yes, so return the name.
            hr = pTISrc->GetNames(dispid, pBSTR, 1, &cbstr);
            goto Cleanup;
        }

        pTISrc->ReleaseFuncDesc(pFDesc);
        pFDesc = NULL;
    }

    hr = S_OK;

Cleanup:
    if(pFDesc)
    {
        pTISrc->ReleaseFuncDesc(pFDesc);
    }
    if(pTAttr)
    {
        pTISrc->ReleaseTypeAttr(pTAttr);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CBase::FireEventV, public
//
//  Synopsis:   Fires an event out the primary dispatch event connection point
//              and using the corresponding code pointers living in attr array
//
//  Arguments:  [dispidEvent]  -- DISPID of event to fire
//              [dispidProp]   -- dispid of property where event ptr is stored
//              [pvbRes]       -- Resultant boolean return value
//              [pbTypes]      -- Pointer to array giving the types of params
//              [va_list]      -- Parameters
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT CBase::FireEventV(DISPID dispidEvent, DISPID dispidProp, IDispatch* pEventObject,
                          VARIANT* pvRes, BYTE* pbTypes, va_list valParms)
{
    HRESULT     hr;
    DISPPARAMS  dp;
    VARIANTARG  av[EVENTPARAMS_MAX];
    CExcepInfo  ei;
    UINT        uArgErr;

    dp.rgvarg = av;
    CParamsToDispParams(&dp, pbTypes, valParms);

    hr = FireEvent(dispidEvent, dispidProp, pvRes, &dp, &ei, &uArgErr, NULL, pEventObject);
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CBase::FireEvent, public
//
//  Synopsis:   Fires an event out the primary dispatch event connection point
//              and using the corresponding code pointers living in attr array
//
//  Arguments:  dispidEvent     Dispid of the event to fire
//              dispidProp      Dispid of property where event func is stored
//              pvarResult      Where to store the return value
//              pdispparams     Parameters for Invoke
//              pexcepinfo      Any exception info
//              puArgErr        Error argument
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
enum EVENTTYPE
{
    EVENTTYPE_EMPTY = 0,
    EVENTTYPE_OLDSTYLE = 1,
    EVENTTYPE_VBHOOKUP = 2
};

enum EVENT_TYPE { VB_EVENTS, CPC_EVENTS };

HRESULT CBase::FireEvent(
         DISPID     dispidEvent, 
         DISPID     dispidProp, 
         VARIANT*   pvarResult, 
         DISPPARAMS*pdispparams,
         EXCEPINFO* pexcepinfo,
         UINT*      puArgErr,
         ITypeInfo* pTIEventSrc/*=NULL*/,
         IDispatch* pEventObject/*=NULL*/)
{
    AAINDEX                     aaidx;
    HRESULT                     hr = S_OK;
    IDispatch*                  pDisp = NULL;
    CVariant                    Var;
    CStackPtrAry<IDispatch*, 2> arySinks;
    CStackDataAry<WORD, 2>      arySinkType; 
    long                        i;
    VARIANT                     varEO;              // Don't release.
    VARIANT*                    pVarEventObject;
    CDispParams                 dispParams;
    BOOL                        f_CPCEvents = FALSE;
    EVENT_TYPE                  whichEventType;

    if(pvarResult)
    {
        VariantInit(pvarResult);
    }

    // First fire off any properties bound to this event.
    if(dispidProp != DISPID_UNKNOWN) // Skip property events
    {
        aaidx = FindAAIndex(dispidProp, CAttrValue::AA_Internal);
        if(AA_IDX_UNKNOWN != aaidx)
        {
            hr = GetDispatchObjectAt(aaidx, &pDisp);
            if(!hr && pDisp)
            {
                CAttrValue* pAV = GetAttrValueAt(aaidx);

                pVarEventObject = NULL;

                hr = InvokeDispatchWithThis(
                    pDisp,
                    pVarEventObject,
                    IID_NULL,
                    _afxGlobalData._lcidUserDefault,
                    DISPATCH_METHOD,
                    pdispparams,
                    &Var,
                    pexcepinfo,
                    NULL);

                ClearInterface(&pDisp);

                if(pvarResult)
                {
                    hr = VariantCopy(pvarResult, &Var);
                    if(hr)
                    {
                        goto Cleanup;
                    }
                }
            }
        }
    }

    // Now fire off to anyone listening through the normal connection
    // points.  Remember that these are also just stored in the attr
    // array with a special internal dispid.
    aaidx = AA_IDX_UNKNOWN;

    for(;;)
    {
        CAttrValue* pAV;

        aaidx = FindNextAAIndex(DISPID_A_EVENTSINK, CAttrValue::AA_Internal, aaidx);
        if(aaidx == AA_IDX_UNKNOWN)
        {
            break;
        }

        Assert(!pDisp);

        pAV = GetAttrValueAt(aaidx);
        if(pAV)
        {
            if(OK(FetchObject(pAV, VT_UNKNOWN, (void**)&pDisp)) && pDisp)
            {
                WORD wET;

                hr = arySinks.Append(pDisp);
                if(hr)
                {
                    goto Cleanup;
                }

                wET  = EVENTTYPE_OLDSTYLE;
                wET |= EVENTTYPE_EMPTY;

                hr = arySinkType.AppendIndirect(&wET, NULL);
                if(hr)
                {
                    goto Cleanup;
                }

                // Upon success arySinks takes over the ref of pDisp
                pDisp = NULL;
            }
        }
    }

    whichEventType = VB_EVENTS;

NextEventSet:
    for(i=0; i<arySinks.Size(); i++)
    {
        DISPPARAMS* pPassedParams;

        VariantClear(&Var);

        // If not old style event, and we have an eventObject and we're not firing the onError
        // event which already has arguments then prepare to pass the event object as an
        // argument.
        if((!(arySinkType[i]&EVENTTYPE_OLDSTYLE)) && pEventObject && dispidProp!=DISPID_EVPROP_ONERROR)
        {
            // Only compute this once.
            if(dispParams.cArgs == 0)
            {
                UINT uIdx;

                // Fix up pdispparam
                dispParams.Initialize(pdispparams->cArgs+1, pdispparams->cNamedArgs);

                hr = dispParams.Create(pdispparams);
                if(hr)
                {
                    goto Cleanup;
                }

                // Now dispParams.cArgs is > 0.
                uIdx = dispParams.cArgs - 1;

                V_VT(&varEO) = VT_DISPATCH;
                V_DISPATCH(&varEO) = pEventObject;

                memcpy(&(dispParams.rgvarg[uIdx]), &varEO, sizeof(VARIANT));
            }

            pPassedParams = &dispParams;
        }
        else
        {
            pPassedParams = pdispparams;
        }

        if(arySinkType[i] & EVENTTYPE_VBHOOKUP)
        {
            // Are we firing VB_EVENTS?
            if(whichEventType == VB_EVENTS)
            {
                // Yes, so fire the event.
                PROPERTYDESC*       pDesc;
                ITridentEventSink*  pTriSink = (ITridentEventSink*)(arySinks[i]);

                if(!FindPropDescFromDispID(dispidProp, (PROPERTYDESC**)&pDesc, NULL, NULL))
                {
                    LPCOLESTR pEventString = (LPCOLESTR)(pDesc->pstrExposedName
                        ? pDesc->pstrExposedName : pDesc->pstrName);
                    pTriSink->FireEvent(pEventString, pPassedParams, &Var, pexcepinfo);
                }
                else if(pTIEventSrc)
                {
                    BSTR eventName;

                    hr = FindEventName(pTIEventSrc, dispidEvent, &eventName);

                    if(hr==S_OK && eventName)
                    {
                        // ActiveX events don't pass the eventObject argument.
                        pTriSink->FireEvent(eventName, pdispparams, &Var, pexcepinfo);
                    }

                    FormsFreeString(eventName);
                }
            }
        }
        else
        {
            // Firing CPC events yet?
            if(whichEventType == CPC_EVENTS)
            {
                arySinks[i]->Invoke(
                    dispidEvent,
                    IID_NULL,
                    _afxGlobalData._lcidUserDefault,
                    DISPATCH_METHOD,
                    pPassedParams,
                    &Var,
                    pexcepinfo,
                    puArgErr);
            }
            else
            {
                // No, just flag we got some CPC events to fire.
                f_CPCEvents = TRUE;
            }
        }

        // If we had to pass the this parameter and we had more than
        // 1 parameter then copy the orginal back over (the args could have
        // been byref).
        if(pPassedParams!=pdispparams && pPassedParams->cArgs>1)
        {
            hr = dispParams.MoveArgsToDispParams(pdispparams, pdispparams->cArgs);
            if(hr)
            {
                goto Cleanup;
            }
        }

        // Copy return value into pvarResult if there was one.
        if(pvarResult && V_VT(&Var)!=VT_EMPTY)
        {
            hr = VariantCopy(pvarResult, &Var);
            if(hr)
            {
                goto Cleanup;
            }
        }
    }

    // Any CPC events?
    if(f_CPCEvents && whichEventType==VB_EVENTS)
    {
        // Yes, and we just finished VB Events so fire all CPC events.
        whichEventType = CPC_EVENTS;
        f_CPCEvents = FALSE;
        goto NextEventSet;
    }

    Assert((whichEventType==CPC_EVENTS && !f_CPCEvents) || whichEventType==VB_EVENTS);

Cleanup:
    ReleaseInterface(pDisp);
    arySinks.ReleaseAll();

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CBase::FirePropertyNotify, public
//
//----------------------------------------------------------------------------
HRESULT CBase::FirePropertyNotify(DISPID dispid, BOOL fOnChanged)
{
    IPropertyNotifySink*    pPNS = NULL;
    HRESULT                 hr = S_OK;
    AAINDEX                 aaidx = AA_IDX_UNKNOWN;
    CStackPtrAry<IPropertyNotifySink*, 2> arySinks;
    long                    i;

    // Take sinks local first, and then fire to avoid reentrancy problems.
    for(;;)
    {
        aaidx = FindNextAAIndex(DISPID_A_PROPNOTIFYSINK, CAttrValue::AA_Internal, aaidx);
        if(aaidx == AA_IDX_UNKNOWN)
        {
            break;
        }

        Assert(!pPNS);
        if(OK(GetUnknownObjectAt(aaidx, (IUnknown**)&pPNS))&&pPNS)
        {
            hr = arySinks.Append(pPNS);
            if(hr)
            {
                goto Cleanup;
            }

            pPNS = NULL; // arySinks has taken over ref.
        }
    }

    for(i=0; i<arySinks.Size(); i++)
    {
        if(fOnChanged)
        {
            arySinks[i]->OnChanged(dispid);
        }
        else
        {
            hr = arySinks[i]->OnRequestEdit(dispid);
            if(!hr)
            {
                // Somebody doesn't want the change.
                break;
            }
        }
    }

Cleanup:
    ReleaseInterface(pPNS);
    arySinks.ReleaseAll();
    RRETURN(hr);
}

#define GETMEMBER_CASE_SENSITIVE    0x00000001
#define GETMEMBER_AS_SOURCE         0x00000002
#define GETMEMBER_ABSOLUTE          0x00000004

//+-------------------------------------------------------------------------
//
//  Method:     CBase::AddExpando, helper
//
//  Synopsis:   Add an expando property to the attr array.
//
//  Results:    S_OK        - return dispid of found name
//              E_          - unable to add expando
//
//--------------------------------------------------------------------------
HRESULT CBase::AddExpando(LPOLESTR pszExpandoName, DISPID* pdispID)
{
    HRESULT     hr;
    CAtomTable* pat;
    BOOL        fExpando;

    pat = GetAtomTable(&fExpando);
    if(!pat || !fExpando)
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

    hr = pat->AddNameToAtomTable(pszExpandoName, pdispID);
    if(hr)
    {
        goto Cleanup;
    }

    // Make it an expando range.
    *pdispID += DISPID_EXPANDO_BASE;
    if(*pdispID > DISPID_EXPANDO_MAX)
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }

Cleanup:
    if(hr)
    {
        *pdispID = DISPID_UNKNOWN;
    }

    RRETURN(hr);
}

HRESULT CBase::GetExpandoName(DISPID expandoDISPID, LPCTSTR* lpPropName)
{
    HRESULT hr;

    Assert(expandoDISPID >= DISPID_EXPANDO_BASE);

    // Get the name associated with the expando DISPID
    CAtomTable* pat = GetAtomTable();

    if(!pat)
    {
        hr = DISP_E_MEMBERNOTFOUND;
        goto Cleanup;
    }
    // When we put the dispid in the attr array we offset it by the expando base
    hr = pat->GetNameFromAtom(expandoDISPID-DISPID_EXPANDO_BASE, lpPropName);
Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::SetExpando, helper
//
//  Synopsis:   Set an expando property to the attr array.
//
//--------------------------------------------------------------------------
HRESULT CBase::SetExpando(LPOLESTR pszExpandoName, VARIANT* pvarPropertyValue)
{
    HRESULT hr;
    DISPID dispIDExpando;
    AAINDEX aaIdx;

    hr = AddExpando(pszExpandoName, &dispIDExpando);
    if(hr)
    {
        goto Cleanup;
    }

    aaIdx = FindAAIndex(dispIDExpando, CAttrValue::AA_Expando);
    if(aaIdx == AA_IDX_UNKNOWN)
    {
        hr = AddVariant(dispIDExpando, pvarPropertyValue, CAttrValue::AA_Expando);
    }
    else
    {
        hr = ChangeVariantAt(aaIdx, pvarPropertyValue);
    }

Cleanup:
    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     CBase::NextProperty, helper
//
//  Synopsis:   Loop through all property entries in ITypeInfo then enumerate
//              through all expando properties.  Any DISPID handed out here
//              must stay constant (GetIDsOfNamesEx should hand out the
//              same).  This routines enumerates regular properties and expando.
//
//  Results:
//              S_OK        - returns next DISPID after currDispID
//              S_FALSE     - enumeration is done
//              E_nnnn      - call failed.
//--------------------------------------------------------------------------
HRESULT CBase::NextProperty(DISPID currDispID, BSTR* pStrNextName, DISPID* pNextDispID)
{
    HRESULT hr = S_FALSE;

    Assert(pNextDispID);

    // Initialize to not found.  If result is S_FALSE and pNextDispID is
    // currDispID then the last known item was a property found in this routine,
    // however there isn't another property after this one.  If pNextDispID is
    // DISPID_UNKNOWN then the last known property wasn't fetched from this
    // routine.  This is important for activex controls so we know when to ask
    // for first property versus next property.
    *pNextDispID = DISPID_UNKNOWN;

    if(IsExpandoDISPID(currDispID))
    {
        goto Cleanup;
    }

    // Enumerate thru attributes and properties.
    if(pStrNextName)
    {
        const VTABLEDESC*   pVTblArray;
        const PROPERTYDESC* ppropdesc;
        BOOL                fPropertyFound = FALSE;
        DWORD               dwppFlags;

        pVTblArray = GetVTableArray();
        if(!pVTblArray)
        {
            goto Cleanup;
        }

        while(pVTblArray->pPropDesc)
        {
            ppropdesc = pVTblArray->pPropDesc;

            dwppFlags = ppropdesc->GetPPFlags();

            // Only iterate over properties and items which are not hidden.
            if((dwppFlags&(PROPPARAM_INVOKEGet|PROPPARAM_INVOKESet))==0 || (dwppFlags&PROPPARAM_HIDDEN))
            {
                goto NextEntry;
            }

            // Get the first property.
            if(currDispID == DISPID_STARTENUM)
            {
                fPropertyFound = TRUE;
            }
            else if(currDispID == ppropdesc->GetDispid())
            {
                // Found a match get the next property.
                fPropertyFound = TRUE;
                *pNextDispID = currDispID;  // Signal we found last one now we we need to look for next property.
                goto NextEntry;
            }

            // Look at this propertydesc?
            if(fPropertyFound)
            {
                // Yes, we're either the first one or this is the dispid we're
                // looking for.
                const TCHAR* pstr = ppropdesc->pstrExposedName ? ppropdesc->pstrExposedName : ppropdesc->pstrName;

                // If property has a leading underscore then don't display.
                if(pstr[0] == _T('_'))
                {
                    goto NextEntry;
                }

                // allocate the result string
                *pStrNextName = SysAllocString(pstr);
                if(*pStrNextName == NULL)
                {
                    hr = E_OUTOFMEMORY;
                    goto Cleanup;
                }

                // We're returning this ID.
                *pNextDispID = ppropdesc->GetDispid();

                hr = S_OK;
                break;
            }

NextEntry:
            pVTblArray++;
        }
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

BOOL CBase::IsExpandoDISPID(DISPID dispid, DISPID* pOLESiteExpando/*=NULL*/)
{
    DISPID tmpDispID;

    if(!pOLESiteExpando)
    {
        pOLESiteExpando = &tmpDispID;
    }

    if(dispid>=DISPID_EXPANDO_BASE && dispid<=DISPID_EXPANDO_MAX)
    {
        *pOLESiteExpando = dispid;
    }
    else if(dispid>=DISPID_ACTIVEX_EXPANDO_BASE && dispid<=DISPID_ACTIVEX_EXPANDO_MAX)
    {
        *pOLESiteExpando = (dispid - DISPID_ACTIVEX_EXPANDO_BASE) + DISPID_EXPANDO_BASE;
    }
    else
    {
        *pOLESiteExpando = dispid;
        return FALSE;
    }

    return TRUE;
}

const VTABLEDESC* CBase::FindVTableEntryForName(LPCTSTR szName, BOOL fCaseSensitive, WORD* pVTblOffset)
{
    const VTABLEDESC* found = NULL;
    int nVTblEntries = GetVTableCount();

    if(nVTblEntries)
    {
        int r;
        const VTABLEDESC* pLow = GetVTableArray();
        const VTABLEDESC* pHigh;
        const VTABLEDESC* pMid;

        pHigh = pLow + nVTblEntries - 1;

        // Binary search for property name
        while(pHigh >= pLow)
        {
            if(pHigh != pLow)
            {
                pMid = pLow + (((pHigh-pLow)+1) >> 1);
                r = StrCmpIC(szName, pMid->pPropDesc->pstrExposedName ?
                    pMid->pPropDesc->pstrExposedName : pMid->pPropDesc->pstrName);
                if(r < 0)
                {
                    pHigh = pMid - 1;
                }
                else if(r > 0)
                {
                    pLow = pMid + 1;
                }
                else
                {
                    // If case sensitive then make sure it really matches.
                    if(fCaseSensitive)
                    {
                        if(StrCmpC(szName, pMid->pPropDesc->pstrExposedName ?
                            pMid->pPropDesc->pstrExposedName : pMid->pPropDesc->pstrName))
                        {
                            // Didn't match exactly, return not found.
                            found = NULL;
                            break;
                        }
                    }

                    found = pMid;
                    break;
                }
            }
            else 
            {
                STRINGCOMPAREFN pfnCompareString = fCaseSensitive ? StrCmpC : StrCmpIC;

                found = pfnCompareString(szName,
                    pLow->pPropDesc->pstrExposedName?
                    pLow->pPropDesc->pstrExposedName:pLow->pPropDesc->pstrName)
                    ? NULL : pLow;
                break;
            }
        }
    }

    return found;
}

//+------------------------------------------------------------------------
//
//  Member:     CBase::CLock::CLock
//
//  Synopsis:   Lock resources in CBase object.
//
//-------------------------------------------------------------------------
CBase::CLock::CLock(CBase* pBase)
{
    _pBase = pBase;
    pBase->PunkOuter()->AddRef();
}

//+------------------------------------------------------------------------
//
//  Member:     CServer::CLock::~CLock
//
//  Synopsis:   Unlock resources in CBase object.
//
//-------------------------------------------------------------------------
CBase::CLock::~CLock()
{
    _pBase->PunkOuter()->Release();
}

BOOL CBase::DidFindAAIndexAndDelete(DISPID dispID, CAttrValue::AATYPE aaType)
{
    BOOL fRetValue = TRUE;
    AAINDEX aaidx;

    aaidx = FindAAIndex(dispID, aaType);

    if(AA_IDX_UNKNOWN != aaidx)
    {
        DeleteAt(aaidx);
    }
    else
    {
        fRetValue = FALSE;
    }

    return fRetValue;
}

AAINDEX CBase::FindAAType(CAttrValue::AATYPE aaType, AAINDEX lastIndex)
{
    AAINDEX newIdx = AA_IDX_UNKNOWN;

    if(_pAA)
    {
        int nPos;

        if(lastIndex == AA_IDX_UNKNOWN)
        {
            nPos = 0;
        }
        else
        {
            nPos = ++lastIndex;
        }

        while(nPos < _pAA->Size())
        {
            CAttrValue* pAV = (CAttrValue*)(*_pAA) + nPos;
            // Logical & so we can find more than one type
            if(pAV->AAType() == aaType)
            {
                newIdx = (AAINDEX)nPos;
                break;
            }
            else
            {
                nPos++;
            }
        }
    }
    return newIdx;
}

//+------------------------------------------------------------------------
//
//  Member:     CBase::FindNextAAIndex
//
//  Synopsis:   Find next AAIndex with given dispid and aatype
//
//  Arguments:  dispid      Dispid to look for
//              aatype      AAType to look for
//              paaIdx      The aaindex of the last entry, if AA_IDX_UNKNOWN
//                          will return the 
//
//-------------------------------------------------------------------------
AAINDEX CBase::FindNextAAIndex(DISPID dispid, CAttrValue::AATYPE aaType, AAINDEX aaIndex)
{
    CAttrValue* pAV;

    if(AA_IDX_UNKNOWN == aaIndex)
    {
        return FindAAIndex(dispid, aaType, AA_IDX_UNKNOWN, TRUE);
    }

    aaIndex++;
    if(!_pAA || aaIndex>=(ULONG)_pAA->Size())
    {
        return AA_IDX_UNKNOWN;
    }

    pAV = _pAA->FindAt(aaIndex);
    if(!pAV || pAV->GetDISPID()!=dispid || pAV->GetAAType()!=aaType)
    {
        return AA_IDX_UNKNOWN;
    }

    return aaIndex;
}

HRESULT CBase::GetStringAt(AAINDEX aaIdx, LPCTSTR* ppStr)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    CAttrValue* pAV;

    Assert(ppStr);

    *ppStr = NULL;

    // Any attr array?
    if(_pAA)
    {
        pAV = _pAA->FindAt(aaIdx);
        // Found AttrValue?
        if(pAV)
        {
            *ppStr = pAV->GetLPWSTR();
            hr = S_OK;
        }
    }

    RRETURN(hr);
}

HRESULT CBase::GetIntoStringAt(AAINDEX aaIdx, BSTR* pbStr, LPCTSTR* ppStr)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    CAttrValue* pAV;

    Assert(ppStr);
    Assert(pbStr);

    *ppStr = NULL;
    *pbStr = NULL;

    // Any attr array?
    if(_pAA)
    {
        pAV = _pAA->FindAt(aaIdx);
        // Found AttrValue?
        if(pAV)
        {
            hr = pAV->GetIntoString(pbStr, ppStr);
        }
    }

    RRETURN(hr);
}

HRESULT CBase::GetIntoBSTRAt(AAINDEX aaIdx, BSTR* pbStr)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    CAttrValue* pAV;

    Assert(pbStr);

    *pbStr = NULL;

    // Any attr array?
    if(_pAA)
    {
        pAV = _pAA->FindAt(aaIdx);
        // Found AttrValue?
        if(pAV)
        {
            hr = S_OK;
            switch(pAV->GetAVType())
            {
            case VT_LPWSTR:
                hr = FormsAllocString(pAV->GetLPWSTR(), pbStr);
                break;

            default:
                {
                    VARIANT varDest;
                    VARIANT varSrc;
                    varDest.vt = VT_EMPTY;
                    pAV->GetAsVariantNC(&varSrc);
                    hr = VariantChangeTypeSpecial(&varDest, &varSrc, VT_BSTR);
                    if(hr)
                    {
                        VariantClear(&varDest);
                        if(hr == DISP_E_TYPEMISMATCH)
                        {
                            hr = S_FALSE;
                        }
                        goto Cleanup;
                    }
                    *pbStr = V_BSTR(&varDest);
                }
                break;
            }
        }
    }
Cleanup:
    RRETURN(hr);
}

HRESULT CBase::GetPointerAt(AAINDEX aaIdx, void** ppv)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    CAttrValue* pAV;

    Assert(ppv);

    *ppv = NULL;

    // Any attr array?
    if(_pAA)
    {
        pAV = _pAA->FindAt(aaIdx);
        // Found AttrValue?
        if(pAV)
        {
            *ppv = pAV->GetPointer();
            hr = S_OK;
        }
    }

    RRETURN(hr);
}

HRESULT CBase::GetVariantAt(AAINDEX aaIdx, VARIANT* pVar, BOOL fAllowNullVariant)
{
    HRESULT hr = S_OK;
    const CAttrValue* pAV;

    Assert(pVar);

    // Don't have an attr array or attr value will return
    // NULL variant.  Note, this will result in the return result
    // being a null.
    // Any attr array?
    if(_pAA && ((pAV=_pAA->FindAt(aaIdx))!=NULL))
    {
        hr = pAV->GetIntoVariant(pVar);
        if (hr && fAllowNullVariant)
        {
            pVar->vt = VT_NULL;
        }
    }
    else
    {
        if(fAllowNullVariant)
        {
            pVar->vt = VT_NULL;
        }
        else
        {
            hr = DISP_E_MEMBERNOTFOUND;
        }
    }

    RRETURN(hr);
}

UINT CBase::GetVariantTypeAt(AAINDEX aaIdx)
{
    if(_pAA)
    {
        CAttrValue* pAV;

        pAV = _pAA->FindAt(aaIdx);
        // Found AttrValue?
        if(pAV)
        {
            return pAV->GetAVType();
        }
    }

    return VT_EMPTY;
}

HRESULT CBase::AddSimple(DISPID dispID, DWORD dwSimple, CAttrValue::AATYPE aaType)
{
    VARIANT varNew;

    varNew.vt = VT_I4;
    varNew.lVal = (long)dwSimple;

    RRETURN(CAttrArray::Set(&_pAA, dispID, &varNew, NULL, aaType));
}

HRESULT CBase::AddPointer(DISPID dispID, void* pValue, CAttrValue::AATYPE aaType)
{
    VARIANT varNew;

    varNew.vt = VT_PTR;
    varNew.byref = pValue;

    RRETURN(CAttrArray::Set(&_pAA, dispID, &varNew, NULL, aaType));
}

HRESULT CBase::AddString(DISPID dispID, LPCTSTR pch, CAttrValue::AATYPE aaType)
{
    VARIANT varNew;

    varNew.vt = VT_LPWSTR;
    varNew.byref = (LPVOID)pch;

    RRETURN(CAttrArray::Set(&_pAA, dispID, &varNew, NULL, aaType));
}

HRESULT CBase::AddBSTR(DISPID dispID, LPCTSTR pch, CAttrValue::AATYPE aaType)
{
    VARIANT varNew;
    HRESULT hr;

    varNew.vt = VT_BSTR;
    hr = FormsAllocString(pch, &V_BSTR(&varNew));
    if(hr)
    {
        goto Cleanup;
    }

    hr = CAttrArray::Set(&_pAA, dispID, &varNew, NULL, aaType);
    if(hr)
    {
        FormsFreeString(V_BSTR(&varNew));
    }
Cleanup:
    RRETURN(hr);
}

HRESULT CBase::AddUnknownObject(DISPID dispID, IUnknown* pUnk, CAttrValue::AATYPE aaType)
{
    VARIANT varNew;

    varNew.vt = VT_UNKNOWN;
    varNew.punkVal = pUnk;

    RRETURN(CAttrArray::Set(&_pAA, dispID, &varNew, NULL, aaType));
}

HRESULT CBase::AddDispatchObject(DISPID dispID, IDispatch* pDisp, CAttrValue::AATYPE aaType)
{
    VARIANT varNew;

    varNew.vt = VT_DISPATCH;
    varNew.pdispVal = pDisp;

    RRETURN(CAttrArray::Set(&_pAA, dispID, &varNew, NULL, aaType));
}

HRESULT CBase::AddVariant(DISPID dispID, VARIANT* pVar, CAttrValue::AATYPE aaType)
{
    RRETURN(CAttrArray::Set(&_pAA, dispID, pVar, NULL, aaType));
}

HRESULT CBase::AddAttrArray(DISPID dispID, CAttrArray* pAttrArray, CAttrValue::AATYPE aaType)
{
    VARIANT varNew;

    varNew.vt = CAttrValue::VT_ATTRARRAY;
    varNew.byref = pAttrArray;

    RRETURN(CAttrArray::Set(&_pAA, dispID, &varNew, NULL, aaType));
}

//+---------------------------------------------------------------------------------
//
//  Member:     CBase::AddUnknownObjectMultiple
//
//  Synopsis:   Add an object to the attr array allowing for multiple
//              entries at this dispid.
//
//----------------------------------------------------------------------------------
HRESULT CBase::AddUnknownObjectMultiple(DISPID dispID, IUnknown* pUnk, 
                                        CAttrValue::AATYPE aaType,
                                        CAttrValue::AAExtraBits wFlags/*=CAttrValue::AA_Extra_Empty*/)
{
    VARIANT var;

    V_VT(&var) = VT_UNKNOWN;
    V_UNKNOWN(&var) = pUnk;

    RRETURN(CAttrArray::Set(&_pAA, dispID, &var, NULL, aaType, wFlags, TRUE));
}

//+---------------------------------------------------------------------------------
//
//  Member:     CBase::AddDispatchObjectMultiple
//
//  Synopsis:   Add an object to the attr array allowing for multiple
//              entries at this dispid.
//
//----------------------------------------------------------------------------------
HRESULT CBase::AddDispatchObjectMultiple(DISPID dispID, IDispatch* pDisp, CAttrValue::AATYPE aaType)
{
    VARIANT var;

    V_VT(&var) = VT_DISPATCH;
    V_UNKNOWN(&var) = pDisp;

    RRETURN(CAttrArray::Set(&_pAA, dispID, &var, NULL, aaType, 0, TRUE));
}

HRESULT CBase::ChangeSimpleAt(AAINDEX aaIdx, DWORD dwSimple)
{
    VARIANT varNew;

    varNew.vt = VT_I4;
    varNew.lVal = (long)dwSimple;

    RRETURN(_pAA?_pAA->SetAt(aaIdx, &varNew):DISP_E_MEMBERNOTFOUND);
}

HRESULT CBase::ChangeStringAt(AAINDEX aaIdx, LPCTSTR pch)
{
    VARIANT varNew;

    varNew.vt = VT_LPWSTR;
    varNew.byref = (LPVOID)pch;

    RRETURN(_pAA?_pAA->SetAt(aaIdx, &varNew):DISP_E_MEMBERNOTFOUND);
}

HRESULT CBase::ChangeUnknownObjectAt(AAINDEX aaIdx, IUnknown* pUnk)
{
    VARIANT varNew;

    varNew.vt = VT_UNKNOWN;
    varNew.punkVal = pUnk;

    RRETURN(_pAA?_pAA->SetAt(aaIdx, &varNew):DISP_E_MEMBERNOTFOUND);
}

HRESULT CBase::ChangeDispatchObjectAt(AAINDEX aaIdx, IDispatch* pDisp)
{
    VARIANT varNew;

    varNew.vt = VT_DISPATCH;
    varNew.pdispVal = pDisp;

    RRETURN(_pAA?_pAA->SetAt(aaIdx, &varNew):DISP_E_MEMBERNOTFOUND);
}

HRESULT CBase::ChangeVariantAt(AAINDEX aaIdx, VARIANT* pVar)
{
    RRETURN(_pAA?_pAA->SetAt(aaIdx, pVar):DISP_E_MEMBERNOTFOUND);
}

HRESULT CBase::ChangeAATypeAt(AAINDEX aaIdx, CAttrValue::AATYPE aaType)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;

    // Any attr array?
    if(_pAA)
    {
        CAttrValue* pAV;

        pAV = _pAA->FindAt(aaIdx);
        // Found AttrValue?
        if(pAV)
        {
            pAV->SetAAType(aaType);
            hr = S_OK;
        }
    }

    RRETURN(hr);
}

HRESULT CBase::FetchObject(CAttrValue* pAV, VARTYPE vt, void** ppvoid)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;

    Assert(ppvoid);

    // Found AttrValue?
    if(pAV)
    {
        if(pAV->GetAVType()==vt && vt==VT_UNKNOWN)
        {
            *ppvoid = (void*)pAV->GetUnknown();
            pAV->GetUnknown()->AddRef();
        }
        else if(pAV->GetAVType()==vt && vt==VT_DISPATCH)
        {
            *ppvoid = (void*)pAV->GetDispatch();
            pAV->GetDispatch()->AddRef();
        }

        hr = S_OK;
    }

    RRETURN(hr);
}

HRESULT CBase::GetObjectAt(AAINDEX aaIdx, VARTYPE vt, void** ppvoid)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    CAttrValue* pAV;

    Assert(ppvoid);

    // Don't have an attr array or attr value will return
    // NULL IUnknown *.
    *ppvoid = NULL;

    // Any attr array?
    if(_pAA)
    {
        pAV = _pAA->FindAt(aaIdx);
        hr = FetchObject(pAV, vt, ppvoid);
    }

    RRETURN(hr);
}

#ifdef _WIN64
HRESULT CBase::GetCookieAt(AAINDEX aaIdx, DWORD* pdwCookie)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    CAttrValue* pAV;

    Assert(pdwCookie);

    // Don't have an attr array or attr value will return
    // NULL IUnknown *.
    *pdwCookie = 0;

    // Any attr array?
    if(_pAA)
    {
        pAV = _pAA->FindAt(aaIdx);
        if(pAV)
        {
            *pdwCookie = pAV->GetCookie();
            hr = S_OK;
        }
    }

    RRETURN(hr);
}

HRESULT CBase::SetCookieAt(AAINDEX aaIdx, DWORD dwCookie)
{
    HRESULT hr = DISP_E_MEMBERNOTFOUND;
    CAttrValue* pAV;

    // Any attr array?
    if(_pAA)
    {
        pAV = _pAA->FindAt(aaIdx);
        if(pAV)
        {
            pAV->SetCookie(dwCookie);
            hr = S_OK;
        }
    }

    RRETURN(hr);
}
#endif



//+-------------------------------------------------------------------------
//
//  Method:     MatchExactGetIDsOfNames, (exported helper used by shdocvw)
//
//  Synopsis:   Loop through all property entries in ITypeInfo does a case
//              sensitive match (GetIDsofNames is case insensitive).
//
//  Results:    S_OK                - return dispid of found name
//              DISP_E_UNKNOWNNAME  - name not found
//
//--------------------------------------------------------------------------
STDAPI MatchExactGetIDsOfNames(ITypeInfo* pTI, REFIID riid, LPOLESTR* rgszNames,
                               UINT cNames, LCID lcid, DISPID* rgdispid, BOOL fCaseSensitive)
{
    HRESULT hr = S_OK;
    CTypeInfoNav tin;
    STRINGCOMPAREFN pfnCompareString;

    if(cNames == 0)
    {
        goto Cleanup;
    }

    if(!IsEqualIID(riid, IID_NULL) || !pTI || !rgszNames || !rgdispid)
    {
        RRETURN(E_INVALIDARG);
    }

    pfnCompareString = fCaseSensitive ? StrCmpC : StrCmpI;

    rgdispid[0] = DISPID_UNKNOWN;

    // Loop thru properties.
    hr = tin.InitITypeInfo(pTI);
    while((hr=tin.Next()) == S_OK)
    {
        VARDESC*    pVar;
        FUNCDESC*   pFunc;
        DISPID      memid = DISPID_UNKNOWN;

        if((pVar=tin.getVarD()) != NULL)
        {
            memid = pVar->memid;
        }
        else if((pFunc=tin.getFuncD()) != NULL)
        {
            memid = pFunc->memid;
        }

        // Got a property?
        if(memid != DISPID_UNKNOWN)
        {
            BSTR bstrName;
            unsigned int cNameRet;

            // Get the name.
            hr = pTI->GetNames(memid, &bstrName, 1, &cNameRet);
            if(hr)
            {
                break;
            }

            if(cNameRet == 1)
            {
                // Does it match?
                if(pfnCompareString(rgszNames[0], bstrName) == 0)
                {
                    rgdispid[0] = memid;
                    SysFreeString(bstrName);
                    break;
                }

                SysFreeString(bstrName);
            }
        }
    }

Cleanup:
    if(hr || rgdispid[0]==DISPID_UNKNOWN)
    {
        hr = DISP_E_UNKNOWNNAME;
    }

    RRETURN(hr);
}

//+-------------------------------------------------------------------------
//
//  Method:     attachEvent
//
//  Synopsis:   Adds an AA_AttachEvent entry to attrarray to support multi-
//              casting of onNNNNN events.
//
//--------------------------------------------------------------------------
HRESULT CBase::attachEvent(BSTR bstrEvent, IDispatch* pDisp, VARIANT_BOOL* pResult)
{
    HRESULT         hr = S_OK;
    DISPID          dispid;
    IDispatchEx*    pDEXMe = NULL;

    if(!pDisp || !bstrEvent)
    {
        goto Cleanup;
    }

    hr = PrivateQueryInterface(IID_IDispatchEx, (void**)&pDEXMe);
    if(hr)
    {
        goto Cleanup;
    }

    // BUGBUG: (anandra) Always being case sensitive here.  Need to
    // check for VBS and be insensitive then.
    hr = pDEXMe->GetDispID(bstrEvent, fdexNameCaseSensitive|fdexNameEnsure, &dispid);
    if(hr)
    {
        goto Cleanup;
    }

    hr = AddDispatchObjectMultiple(dispid, pDisp, CAttrValue::AA_AttachEvent);

    // let opp know that an event was attached.
    OnPropertyChange(DISPID_EVPROP_ONATTACHEVENT, 0);

Cleanup:
    if(pResult)
    {
        *pResult = hr ? VARIANT_FALSE : VARIANT_TRUE;
    }
    ReleaseInterface(pDEXMe);
    RRETURN(SetErrorInfo(hr));
}

//+-------------------------------------------------------------------------
//
//  Method:     detachEvent
//
//  Synopsis:   Loops through all AA_AttachEvent entries in the attrarray
//              and removes first entry who's COM identity is the same as
//              the pDisp passed in.
//
//--------------------------------------------------------------------------
HRESULT CBase::detachEvent(BSTR bstrEvent, IDispatch* pDisp)
{
    AAINDEX         aaidx = AA_IDX_UNKNOWN;
    DISPID          dispid;
    IDispatch*      pThisDisp = NULL;
    HRESULT         hr = S_OK;
    IDispatchEx*    pDEXMe = NULL;

    if(!pDisp || !bstrEvent)
    {
        goto Cleanup;
    }

    hr = PrivateQueryInterface(IID_IDispatchEx, (void**)&pDEXMe);
    if(hr)
    {
        goto Cleanup;
    }

    // BUGBUG: (anandra) Always being case sensitive here.  Need to
    // check for VBS and be insensitive then.
    hr = pDEXMe->GetDispID(bstrEvent, fdexNameCaseSensitive, &dispid);
    if(hr)
    {
        goto Cleanup;
    }

    // Find event that has this function pointer.
    for(;;)
    {
        aaidx = FindNextAAIndex(dispid, CAttrValue::AA_AttachEvent, aaidx);
        if(aaidx == AA_IDX_UNKNOWN)
        {
            break;
        }

        ClearInterface(&pThisDisp);
        if(GetDispatchObjectAt(aaidx, &pThisDisp))
        {
            continue;
        }

        if(IsSameObject(pDisp, pThisDisp))
        {
            break;
        }
    };

    // Found item to delete?
    if(aaidx != AA_IDX_UNKNOWN)
    {
        DeleteAt(aaidx);
    }

Cleanup:
    ReleaseInterface(pThisDisp);
    ReleaseInterface(pDEXMe);
    RRETURN(SetErrorInfo(S_OK));
}

HRESULT STDMETHODCALLTYPE CBase::getAttribute(BSTR bstrPropertyName, LONG lFlags, VARIANT* pvarPropertyValue)
{
    HRESULT hr;
    DISPID dispID;
    DISPPARAMS dp = _afxGlobalData._Zero.dispparams;
    CVariant varNULL(VT_NULL);
    PROPERTYDESC* propDesc = NULL;
    IDispatchEx* pDEX = NULL;

    if(!bstrPropertyName || !pvarPropertyValue)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    pvarPropertyValue->vt = VT_NULL;

    hr = PrivateQueryInterface(IID_IDispatchEx, (void**)&pDEX);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pDEX->GetDispID(bstrPropertyName,
        lFlags&GETMEMBER_CASE_SENSITIVE?fdexNameCaseSensitive:0,
        &dispID);
    if(hr)
    {
        goto Cleanup;
    }

    // If we're asked for the SaveAs value - our best guess is to grab the attr array value. This won't
    // always work - but it's the best we can do!
    if(lFlags&GETMEMBER_AS_SOURCE || lFlags&GETMEMBER_ABSOLUTE)
    {
        // Here we pretty much do what we do at Save time. The only difference is we don't do the GetBaseObject call
        // so certain re-directed properties won't work - e.g. the BODY onXXX properties
        if(IsExpandoDISPID(dispID, &dispID))
        {
            hr = GetIntoBSTRAt(FindAAIndex(dispID, CAttrValue::AA_Expando),
                &V_BSTR(pvarPropertyValue));
        }
        else
        {
            // Try the Unknown value first...
            hr = GetIntoBSTRAt(FindAAIndex(dispID, CAttrValue::AA_UnknownAttr),
                &V_BSTR(pvarPropertyValue));
            if(hr == DISP_E_MEMBERNOTFOUND)
            {
                // No Unknown - try getting the attribute as it would be saved
                hr = FindPropDescFromDispID(dispID, &propDesc, NULL, NULL);
                Assert(propDesc);
                if(!hr && propDesc->pfnHandleProperty)
                {
                    if(propDesc->pfnHandleProperty) 
                    {
                        hr = propDesc->HandleGetIntoAutomationVARIANT(this, pvarPropertyValue);

                        // This flag only works for URL properties we want to
                        // return the fully expanded URL.
                        if(lFlags&GETMEMBER_ABSOLUTE &&
                            propDesc->pfnHandleProperty==&PROPERTYDESC::HandleUrlProperty &&
                            !hr)
                        {
                            hr = ExpandedRelativeUrlInVariant(pvarPropertyValue);
                        }

                        goto Cleanup;
                    }
                    else
                    {
                        hr = DISP_E_UNKNOWNNAME;
                    }
                }
            }
        }

        if(!hr)
        {
            pvarPropertyValue->vt = VT_BSTR;
        }
    }
    else
    {
        // Need to check for an unknown first, because the regular get_'s will return
        // a default. No point in doing this if we're looking at an expando
        if(!IsExpandoDISPID(dispID))
        {
            hr = GetIntoBSTRAt(FindAAIndex(dispID, CAttrValue::AA_UnknownAttr),
                &V_BSTR(pvarPropertyValue));
            // If this worked - we're done.
            if(!hr)
            {
                pvarPropertyValue->vt = VT_BSTR;
                goto Cleanup;
            }
        } 

        // See if we can get a regular property or expando
        hr = pDEX->InvokeEx(
            dispID,
            _afxGlobalData._lcidUserDefault,
            DISPATCH_PROPERTYGET,
            &dp,
            pvarPropertyValue,
            NULL,
            NULL);
    }

Cleanup:
    ReleaseInterface(pDEX);
    if(hr==DISP_E_UNKNOWNNAME || hr==DISP_E_MEMBERNOTFOUND)
    {
        // Couldn't find property - return a Null rather than an error
        hr = S_OK;
    }
    RRETURN(SetErrorInfo(hr));
}

HRESULT STDMETHODCALLTYPE CBase::setAttribute(BSTR strPropertyName, VARIANT varPropertyValue, LONG lFlags)
{
    HRESULT hr;
    DISPID dispID;
    DISPPARAMS dp;
    EXCEPINFO excepinfo;
    UINT uArgError;
    DISPID dispidPut = DISPID_PROPERTYPUT; // Dispid of prop arg.
    CVariant varNULL(VT_NULL);
    IDispatchEx* pDEX = NULL;
    VARIANT* pVar;

    // Implementation leverages the existing Invoke mechanism
    InitEXCEPINFO(&excepinfo);

    if(!strPropertyName)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    hr = PrivateQueryInterface(IID_IDispatchEx, (void**)&pDEX);
    if(hr)
    {
        goto Cleanup;
    }

    hr = pDEX->GetDispID(strPropertyName,
        lFlags&GETMEMBER_CASE_SENSITIVE?fdexNameCaseSensitive|fdexNameEnsure:fdexNameEnsure,
        &dispID);
    if(hr)
    {
        goto Cleanup;
    }

    pVar = V_ISBYREF(&varPropertyValue) ? V_VARIANTREF(&varPropertyValue) : &varPropertyValue;

    dp.rgvarg = pVar;
    dp.rgdispidNamedArgs = &dispidPut;
    dp.cArgs = 1;
    dp.cNamedArgs = 1;

    // See if it's accepted as a regular property or expando
    hr = pDEX->Invoke(
        dispID,
        IID_NULL,
        _afxGlobalData._lcidUserDefault,
        DISPATCH_PROPERTYPUT,
        &dp,
        NULL,
        &excepinfo,
        &uArgError);

    if(hr)
    {
        // Failed to parse .. make an unknown
        CVariant varBSTR;
        const PROPERTYDESC* ppropdesc;
        const BASICPROPPARAMS* bpp;

        CLock Lock (this); // For OnPropertyChange

        ppropdesc = FindPropDescForName(strPropertyName);
        if(!ppropdesc)
        {
            goto Cleanup;
        }

        // It seems sensible to only allow string values for unknowns so we can always
        // persist them.
        // Try to see if it parses as a valid value
        // Coerce it to a string...
        hr = VariantChangeTypeSpecial(&varBSTR, pVar, VT_BSTR, NULL, VARIANT_NOVALUEPROP);
        if(hr)
        {
            goto Cleanup;
        }

        // Add it as an unknown (invalid) value
        // SetString with fIsUnkown set to TRUE
        hr = CAttrArray::SetString(&_pAA, ppropdesc, (LPTSTR)varBSTR.bstrVal, TRUE);
        if(hr)
        {
            goto Cleanup;
        }

        bpp = (const BASICPROPPARAMS*)(ppropdesc+1);

        // Need to fire a property change event
        hr = OnPropertyChange(bpp->dispid, bpp->dwFlags);
    }

Cleanup:
    ReleaseInterface(pDEX);
    FreeEXCEPINFO(&excepinfo);
    RRETURN(SetErrorInfo(hr));
}

HRESULT STDMETHODCALLTYPE CBase::removeAttribute(BSTR strPropertyName, LONG lFlags, VARIANT_BOOL* pfSuccess)
{
    DISPID dispID;
    IDispatchEx* pDEX = NULL;

    if(pfSuccess)
    {
        *pfSuccess = VB_FALSE;
    }

    if(PrivateQueryInterface(IID_IDispatchEx, (void**)&pDEX))
    {
        goto Cleanup;
    }

    if(pDEX->GetDispID(strPropertyName,
        lFlags&GETMEMBER_CASE_SENSITIVE?fdexNameCaseSensitive:0, &dispID))
    {
        goto Cleanup;
    }

    if(!removeAttributeDispid(dispID))
    {
        goto Cleanup;
    }

    if(pfSuccess)
    {
        *pfSuccess = VB_TRUE;
    }

Cleanup:
    ReleaseInterface(pDEX);

    RRETURN(SetErrorInfo(S_OK));
    return E_NOTIMPL;
}