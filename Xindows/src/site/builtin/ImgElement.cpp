
#include "stdafx.h"
#include "ImgElement.h"

#include "ImgHelper.h"

#define _cxx_
#include "../gen/img.hdl"

#include <Ratings.h>

extern NEWIMGTASKFN NewImgTaskArt;

#define DATE_STR_LENGTH 30
//+---------------------------------------------------------------------------
//
//  method : ConvertDateTimeToString
//
//  Synopsis:
//          Converts given date and time if requested to string using the local time
//          in a fixed format mm/dd/yyyy or mm/dd/yyy hh:mm:ss
//          
//----------------------------------------------------------------------------

// NB (cthrash) This function produces a standard(?) format date, of the form
// MM/DD/YY HH:MM:SS (in military time.) The date format *will not* be tailored
// for the locale.  This is for Netscape compatibility, and is a departure from
// how IE3 worked.  If you want the date in the format of the document's locale,
// you should use the Java Date object.
HRESULT ConvertDateTimeToString(FILETIME Time, BSTR* pBstr, BOOL fReturnTime)
{
    HRESULT     hr;
    SYSTEMTIME  SystemTime;
    TCHAR       pchDateStr[DATE_STR_LENGTH];
    FILETIME    ft;

    Assert(pBstr);

    // We want to return local time as Nav not GMT 
    if(!FileTimeToLocalFileTime(&Time, &ft))
    {
        hr  = GetLastWin32Error();
        goto Cleanup;
    }

    if(!FileTimeToSystemTime(&ft, &SystemTime))
    {
        hr  = GetLastWin32Error();
        goto Cleanup;
    }

    // We want Gregorian dates and 24-hour time.
    if(fReturnTime)
    {
        hr = Format(_afxGlobalData._hInstResource, 0, pchDateStr, ARRAYSIZE(pchDateStr),
            _T("<0d2>/<1d2>/<2d4> <3d2>:<4d2>:<5d2>"),
            SystemTime.wMonth,
            SystemTime.wDay,
            SystemTime.wYear,
            SystemTime.wHour,
            SystemTime.wMinute,
            SystemTime.wSecond);
    }
    else
    {
        hr = Format(_afxGlobalData._hInstResource, 0, pchDateStr, ARRAYSIZE(pchDateStr),
            _T("<0d2>/<1d2>/<2d4>"),
            SystemTime.wMonth,
            SystemTime.wDay,
            SystemTime.wYear);
    }

    if(hr)
    {
        goto Cleanup;
    }

    hr = FormsAllocString(pchDateStr, pBstr);

Cleanup:
    RRETURN(hr);
}




const CElement::CLASSDESC CImgElement::s_classdesc =
{
    {
        &CLSID_HTMLImg,                 // _pclsid
        s_acpi,                         // _pcpi
        ELEMENTDESC_CARETINS_SL |
        ELEMENTDESC_NEVERSCROLL |
        ELEMENTDESC_EXBORDRINMOV,       // _dwFlags
        &IID_IHTMLImgElement,           // _piidDispinterface
        &s_apHdlDescs,                  // _apHdlDesc
    },
    (void*)s_apfnpdIHTMLImgElement,     // _pfnTearOff
    NULL,                               // _pAccelsDesign
    NULL                                // _pAccelsRun
};

CImgElement::CImgElement (ELEMENT_TAG eTag, CDocument* pDoc) : super(eTag, pDoc)
{
    _fCanClickImage = FALSE;
}

//+------------------------------------------------------------------------
//
//  Member:     Init
//
//  Synopsis:   
//
//-------------------------------------------------------------------------
HRESULT CImgElement::Init()
{
    HRESULT hr = S_OK;

    _pImage = new CImgHelper(Doc(), this);

    if(!_pImage)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    hr = super::Init();

Cleanup:
    RRETURN(hr);
}

HRESULT CImgElement::CreateElement(CHtmTag* pht, CDocument* pDoc, CElement** ppElement)
{
    Assert(ppElement);

    *ppElement = new CImgElement(pht->GetTag(), pDoc);

    RRETURN((*ppElement) ? S_OK : E_OUTOFMEMORY);
}

//+------------------------------------------------------------------------
//
//  Member:     CImgElement::PrivateQueryInterface, IUnknown
//
//  Synopsis:   Private unknown QI.
//
//-------------------------------------------------------------------------
HRESULT CImgElement::PrivateQueryInterface(REFIID iid, void** ppv)
{
    *ppv = NULL;
    switch(iid.Data1)
    {
        QI_HTML_TEAROFF(this, IHTMLElement2, NULL)
    }

    if(*ppv)
    {
        ((IUnknown*)*ppv)->AddRef();
        RRETURN(S_OK);
    }

    RRETURN(super::PrivateQueryInterface(iid, ppv));
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgElement::Notify
//
//  Synopsis:   Receives notifications
//
//+---------------------------------------------------------------------------
void CImgElement::Notify(CNotification* pNF)
{
    HRESULT         hr = S_OK;
    CElement::CLock Lock(this);

    super::Notify(pNF);
    Assert(_pImage);

    _pImage->Notify(pNF);
    switch(pNF->Type())
    {
    case NTYPE_ELEMENT_QUERYFOCUSSABLE:
        if(!IsEditable(TRUE))
        {
            CQueryFocus* pQueryFocus = (CQueryFocus*)pNF->DataAsPtr();

            pQueryFocus->_fRetVal = FALSE;
        }
        break;
    case NTYPE_ELEMENT_QUERYTABBABLE:
        if(!IsEditable(TRUE))
        {
            CQueryFocus* pQueryFocus = (CQueryFocus*)pNF->DataAsPtr();

            // Assume that focussability is already checked for, and only make
            // sure that tabIndex is non-negative for subdivision
            Assert(IsFocussable(pQueryFocus->_lSubDivision));
            pQueryFocus->_fRetVal = TRUE;
        }
        break;
    case NTYPE_ELEMENT_ENTERTREE:
        EnterTree();
        break;

    case NTYPE_BASE_URL_CHANGE:
        OnPropertyChange(DISPID_CImgElement_src,    ((PROPERTYDESC*)&s_propdescCImgElementsrc)->GetdwFlags());
        OnPropertyChange(DISPID_CImgElement_dynsrc, ((PROPERTYDESC*)&s_propdescCImgElementdynsrc)->GetdwFlags());
        OnPropertyChange(DISPID_CImgElement_lowsrc, ((PROPERTYDESC*)&s_propdescCImgElementlowsrc)->GetdwFlags());
        break;
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     EnterTree
//
//+---------------------------------------------------------------------------
HRESULT CImgElement::EnterTree()
{
    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     HandleMessage
//
//  Synopsis:   Handle messages bubbling when the passed site is non null
//
//  Arguments:  [pMessage]  -- message
//              [pChild]    -- pointer to child when bubbling allowed
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CImgElement::HandleMessage(CMessage* pMessage)
{
    HRESULT hr              = S_FALSE;
    BOOL    fInBrowseMode   = !IsEditable(TRUE);
    TCHAR   cBuf[pdlUrlLen];
    TCHAR*  pchExpandedUrl  = cBuf;
    IUniformResourceLocator* pURLToDrag = NULL;

    if(!hr)
    {
        goto Cleanup;
    }

    // WM_CONTEXTMENU message should always be handled.
    if(pMessage->message == WM_CONTEXTMENU)
    {
        Assert(_pImage);
        hr = _pImage->ShowImgContextMenu(pMessage);
    }

    // And process the message if it hasn't been already.
    if(hr == S_FALSE)
    {
        switch(pMessage->message)
        {
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
            if(fInBrowseMode)
            {
                // if it was not NULL it was handled before
                Doc()->SetMouseCapture(
                    MOUSECAPTURE_METHOD(CImgElement, HandleCaptureMessageForImage, handlecapturemessageforimage),
                    (void*)this, TRUE);
                hr = S_OK;
            }
            break;

        case WM_MENUSELECT:
        case WM_INITMENUPOPUP:
            hr = S_FALSE;
            break;

        case WM_SETCURSOR:
            if(!IsEditable())
            {
                SetCursorStyle(IDC_ARROW);
            }
            else
            {
                SetCursorStyle(IDC_SIZEALL);
            }
            hr = S_OK;
            break;
        }

        if(hr == S_FALSE)
        {
            hr = super::HandleMessage(pMessage);
        }
    }

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     HandleCaptureMessageForImage
//
//  Synopsis:   Tracks mouse while user is clicking on an IMG in an A
//
//  Arguments:  [pMessage]  -- message
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CImgElement::HandleCaptureMessageForImage(CMessage* pMessage)
{
    HRESULT hr = S_OK;

    switch(pMessage->message)
    {
    case WM_LBUTTONUP:
        if(_fCanClickImage)
        {
            pMessage->SetNodeClk(GetFirstBranch());
        }
        // Fall thru

    case WM_MBUTTONUP:
    case WM_RBUTTONUP:
        Doc()->SetMouseCapture(NULL, NULL);
        if(pMessage->message == WM_RBUTTONUP)
        {
            hr = S_FALSE;
        }
        break;

    case WM_MOUSEMOVE:
        {
            // If the user moves the mouse outside the wobble zone,
            // show the no-entry , plus disallow a subsequent OnClick
            POINT ptCursor = { LOWORD(pMessage->lParam), HIWORD(pMessage->lParam) };
            CDocument* pDoc = Doc();

            if(_fCanClickImage && !PtInRect(&_rcWobbleZone, ptCursor))
            {
                _fCanClickImage = FALSE;
            }

            // initiate drag-drop
            if(!_fCanClickImage && !pDoc->_fIsDragDropSrc)
            {
                Assert(!pDoc->_pDragDropSrcInfo);
                if(!pDoc->_pElementOMCapture)
                {
                    DragElement(GetCurLayout(), pMessage->dwKeyState, NULL, -1);
                }
            }
            // Intentional drop through to WM_SETCURSOR - WM_SETCURSOR is NOT sent
            // while the Capture is set
        }

    case WM_SETCURSOR:
        {
            LPCTSTR idc;
            CRect rc;

            GetCurLayout()->GetClientRect(&rc);

            idc = IDC_ARROW;

            SetCursorStyle(idc);
            hr = S_OK;
        }
        break;
    }

    if(hr == S_FALSE)
    {
        hr = super::HandleMessage(pMessage);
    }

    RRETURN1(hr, S_FALSE);
}

//+---------------------------------------------------------------------------
//
//  Member:     HandleCaptureMessageForArea
//
//  Synopsis:   Tracks mouse while user is clicking on an IMG with an AREA
//
//  Arguments:  [pMessage]  -- message
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT STDMETHODCALLTYPE CImgElement::HandleCaptureMessageForArea(CMessage* pMessage)
{
    RRETURN1(HandleMessage(pMessage), S_FALSE);
}

HRESULT CImgElement::ApplyDefaultFormat(CFormatInfo* pCFI)
{
    HRESULT         hr          = S_OK;
    CUnitValue      uvBorder    = GetAAborder();
    int             i;

    _fBelowAnchor = FALSE;

    if(uvBorder.IsNull())
    {
        // check if the image is inside an anchor
        if(_fBelowAnchor)
        {
            uvBorder.SetValue(2, CUnitValue::UNIT_PIXELS);
        }
    }

    // Set the anchor border
    if(!uvBorder.IsNull())
    {
        COLORREF crColor;
        DWORD dwRawValue = uvBorder.GetRawValue();

        crColor = 0x00000000;

        pCFI->PrepareFancyFormat();

        for(i=BORDER_TOP; i<=BORDER_LEFT; i++)
        {
            pCFI->_ff()._bBorderStyles[i] = fmBorderStyleSingle;
            pCFI->_ff()._ccvBorderColors[i].SetValue(crColor, FALSE);
            pCFI->_ff()._cuvBorderWidths[i].SetRawValue(dwRawValue);
        }

        pCFI->UnprepareForDebug();

    }

    hr = super::ApplyDefaultFormat(pCFI);
    if(hr || !_pImage)
    {
        goto Cleanup;
    }

    _pImage->SetImgAnim(pCFI->_pcf->IsDisplayNone() || pCFI->_pcf->IsVisibilityHidden());

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     CBlockElement::Save
//
//  Synopsis:   Save the tag to the specified stream.
//
//-------------------------------------------------------------------------
HRESULT CImgElement::Save(CStreamWriteBuff* pStmWrBuff, BOOL fEnd)
{
    if(fEnd)
    {
        return S_OK; // No end IMG tag
    }

    CElement* pelAnchorClose = NULL;
    HRESULT hr;

    if(pStmWrBuff->TestFlag(WBF_FOR_RTF_CONV))
    {
        // RichEdit2.0 crashes when it gets rtf with nested field
        // tags (easily generated by the rtf to html converter for
        // <a><img></a> (images in anchors).  To work around this,
        // we close any anchors this image may be in before writing
        // the image tag, and reopen them immediately after.
        pelAnchorClose = GetFirstBranch()->SearchBranchToFlowLayoutForTag(ETAG_A)->SafeElement();
    }

    if(pelAnchorClose)
    {
        hr = pelAnchorClose->WriteTag(pStmWrBuff, TRUE, TRUE);
        if(hr)
        {
            goto Cleanup;
        }
    }

    hr = super::Save(pStmWrBuff, fEnd);
    if(hr)
    {
        goto Cleanup;
    }

    if(pelAnchorClose)
    {
        hr = pelAnchorClose->WriteTag(pStmWrBuff, FALSE, TRUE);
        if(hr)
        {
            goto Cleanup;
        }
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------------
//
//  Member:     OnPropertyChange
//
//  Note:       Called after a property has changed to notify derived classes
//              of the change.  All properties (except those managed by the
//              derived class) are consistent before this call.
//
//              Also, fires a property change notification to the site.
//
//-------------------------------------------------------------------------
HRESULT CImgElement::OnPropertyChange(DISPID dispid, DWORD dwFlags)
{
    HRESULT hr = S_OK;

    switch(dispid)
    {
    case DISPID_CImgElement_src:
        if(_pImage)
        {
            hr = _pImage->SetImgSrc(IMGF_REQUEST_RESIZE|IMGF_INVALIDATE_FRAME);
        }
        break;
    case DISPID_CImgElement_lowsrc:
        if(_pImage)
        {
            LPCTSTR szUrl = GetAAsrc();

            if(!szUrl)
            {
                Assert(_pImage);
                hr = _pImage->FetchAndSetImgCtx(GetAAlowsrc(), IMGF_REQUEST_RESIZE|IMGF_INVALIDATE_FRAME);
            }

        }
        break;

    case DISPID_CImgElement_dynsrc:
        if(_pImage)
        {
            hr = _pImage->SetImgDynsrc();
        }
        break;
    }

    if(OK(hr))
    {
        hr = super::OnPropertyChange(dispid, dwFlags);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgElement::QueryStatus, public
//
//  Synopsis:   Implements QueryStatus for CImgElement
//
//  Notes:      This override of CImgBase::QueryStatus allows special
//              handling of hyperlink context menu entries for images
//              with active areas in client side maps
//
//----------------------------------------------------------------------------
HRESULT CImgElement::QueryStatus(
         GUID*          pguidCmdGroup,
         ULONG          cCmds,
         MSOCMD         rgCmds[],
         MSOCMDTEXT*    pcmdtext)
{
    int idm;

    Trace0("CImgElement::QueryStatus\n");

    Assert(CBase::IsCmdGroupSupported(pguidCmdGroup));
    Assert(cCmds == 1);

    MSOCMD*         pCmd    = &rgCmds[0];
    HRESULT         hr      = S_OK;

    Assert(!pCmd->cmdf);

    idm = CBase::IDMFromCmdID(pguidCmdGroup, pCmd->cmdID);
    switch(idm)
    {
    case IDM_FOLLOWLINKC:
    case IDM_FOLLOWLINKN:
    case IDM_PRINTTARGET:
    case IDM_SAVETARGET:
        // Plug a ratings security hole.
        if((idm==IDM_PRINTTARGET || idm==IDM_SAVETARGET) &&
            S_OK==RatingEnabledQuery())
        {
            pCmd->cmdf = MSOCMDSTATE_DISABLED;
            break;
        }
        break;

    case IDM_IMAGE:
        // When a single image is selected, allow to bring up an insert image dialog
        pCmd->cmdf = MSOCMDSTATE_UP;
    }

    if(!pCmd->cmdf)
    {
        Assert(_pImage);
        hr = _pImage->QueryStatus(
            pguidCmdGroup,
            1,
            pCmd,
            pcmdtext);
        if(!pCmd->cmdf)
        {
            hr = super::QueryStatus(pguidCmdGroup, 1, pCmd, pcmdtext);
        }
    }

Cleanup:
    RRETURN_NOTRACE(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgElement::Exec, public
//
//  Synopsis:   Executes a command on the CImgElement
//
//  Notes:      This override of CImgBase::Exec allows special
//              handling of hyperlink context menu entries for images
//              with active areas in client side maps
//
//----------------------------------------------------------------------------
HRESULT CImgElement::Exec(
                  GUID*         pguidCmdGroup,
                  DWORD         nCmdID,
                  DWORD         nCmdexecopt,
                  VARIANTARG*   pvarargIn,
                  VARIANTARG*   pvarargOut)
{
    Trace0("CImgElement::Exec\n");

    Assert(CBase::IsCmdGroupSupported(pguidCmdGroup));

    int             idm         = CBase::IDMFromCmdID(pguidCmdGroup, nCmdID);
    HRESULT         hr          = MSOCMDERR_E_NOTSUPPORTED;

    if(hr==MSOCMDERR_E_NOTSUPPORTED && _pImage)
    {
        hr = _pImage->Exec(
            pguidCmdGroup,
            nCmdID,
            nCmdexecopt,
            pvarargIn,
            pvarargOut);
    }

    if(hr==MSOCMDERR_E_NOTSUPPORTED)
    {
        hr = super::Exec(
            pguidCmdGroup,
            nCmdID,
            nCmdexecopt,
            pvarargIn,
            pvarargOut);
    }

Cleanup:
    RRETURN_NOTRACE(hr);
}

HRESULT CImgElement::ClickAction(CMessage* pMessage)
{
    return S_OK;
}

const CImageElementFactory::CLASSDESC CImageElementFactory::s_classdesc =
{
    {
        &CLSID_HTMLImageElementFactory,     // _pclsid
        NULL,                               // _pcpi
        0,                                  // _dwFlags
        &IID_IHTMLImageElementFactory,      // _piidDispinterface
        &s_apHdlDescs,                      // _apHdlDesc
    },
    (void*)s_apfnIHTMLImageElementFactory,  // _apfnTearOff
};

HRESULT STDMETHODCALLTYPE CImageElementFactory::create(VARIANT varWidth, VARIANT varHeight, IHTMLImgElement** ppnewElem)
{
    HRESULT         hr;
    CElement*       pElement = NULL;
    CImgElement*    pImgElem;
    CVariant        varI4Width;
    CVariant        varI4Height;

    // We must return into a ptr else there's no-one holding onto a ref!
    if(!ppnewElem)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *ppnewElem = NULL;

    // actualy ( [ long width, long height ] )
    // Create an Image element parented to the root site

    // This is some temproary unfinished code just to test the call
    hr = _pDoc->CreateElement(ETAG_IMG, &pElement);
    if(hr)
    {
        goto Cleanup;
    }

    pImgElem = DYNCAST(CImgElement, pElement);

    // Set the flag that indicates we're parented invisibly to the resize
    // this stops unpleasantness with ResizeElement.
    // BUGBUG: Is this necessary any longer (it's a RequestResize carry-over) (brendand)
    pImgElem->_pImage->_fCreatedWithNew = TRUE;

    hr = pImgElem->QueryInterface(IID_IHTMLImgElement, (void**)ppnewElem);

    // BUGBUG rgardner this doesn't work - it causes a risize request - & GPF!
    // Image code needs to be smart enough to spot images parented to the root site

    // Set the width & height if supplied
    if(varWidth.vt!=VT_EMPTY && varWidth.vt!=VT_ERROR)
    {
        pImgElem->_pImage->_fSizeInCtor = TRUE;
    }

    hr = varI4Width.CoerceVariantArg(&varWidth, VT_I4);
    if(hr == S_OK)
    {
        hr = pImgElem->putWidth(V_I4(&varI4Width));
    }
    if(!OK(hr))
    {
        goto Cleanup;
    }

    hr = varI4Height.CoerceVariantArg(&varHeight, VT_I4);
    if(hr == S_OK)
    {
        hr = pImgElem->putHeight(V_I4(&varI4Height));
    }
    if(!OK(hr))
    {
        goto Cleanup;
    }

Cleanup:
    if(OK(hr))
    {
        hr = S_OK; // not to propagate possible S_FALSE
    }
    else
    {
        ReleaseInterface(*(IUnknown**)ppnewElem);
    }

    CElement::ClearPtr(&pElement);

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
// Member: get_mimeType
//
//----------------------------------------------------------------------------
extern BSTR GetFileTypeInfo(TCHAR* pszFileName);

STDMETHODIMP CImgElement::get_mimeType(BSTR* pMimeType)
{
    HRESULT hr = S_OK;
    TCHAR* pchCachedFile = NULL;

    if(!pMimeType)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *pMimeType = NULL;

    hr = _pImage->GetFile(&pchCachedFile);

    if(!hr && pchCachedFile)
    {
        *pMimeType = GetFileTypeInfo(pchCachedFile);
    }

    MemFreeString(pchCachedFile);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
// Member: get_fileSize
//
//----------------------------------------------------------------------------
STDMETHODIMP CImgElement::get_fileSize(BSTR* pFileSize)
{
    HRESULT hr = S_OK;
    TCHAR   szBuf[64];
    TCHAR*  pchCachedFile = NULL;

    if(pFileSize == NULL)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pFileSize = NULL;

    Assert(_pImage);
    hr = _pImage->GetFile(&pchCachedFile);
    if(!hr && pchCachedFile)
    {
        WIN32_FIND_DATA wfd;
        HANDLE hFind = FindFirstFile(pchCachedFile, &wfd);
        if(hFind != INVALID_HANDLE_VALUE)
        {
            FindClose(hFind);

            Format(_afxGlobalData._hInstResource, 0, szBuf, ARRAYSIZE(szBuf), _T("<0d>"), (long)wfd.nFileSizeLow);
            *pFileSize = SysAllocString(szBuf);
        }
    }
    else
    {
        *pFileSize = SysAllocString(_T("-1"));
        hr = S_OK;
    }

Cleanup:
    MemFreeString(pchCachedFile);
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
// Member: get_fileCreatedDate
//
//-----------------------------------------------------------------------------
STDMETHODIMP CImgElement::get_fileCreatedDate(BSTR* pFileCreatedDate)
{
    HRESULT hr = S_OK;
    TCHAR* pchCachedFile = NULL;

    if(pFileCreatedDate == NULL)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pFileCreatedDate = NULL;

    Assert(_pImage);
    hr = _pImage->GetFile(&pchCachedFile);
    if(!hr && pchCachedFile)
    {
        WIN32_FIND_DATA wfd;
        HANDLE hFind = FindFirstFile(pchCachedFile, &wfd);
        if(hFind != INVALID_HANDLE_VALUE)
        {
            FindClose(hFind);
            // we always return the local time in a fixed format mm/dd/yyyy to make it possible to parse
            // FALSE means we do not want the time
            hr = ConvertDateTimeToString(wfd.ftCreationTime, pFileCreatedDate, FALSE);
        }
    }

Cleanup:
    MemFreeString(pchCachedFile);
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
// Member: get_fileModifiedDate
//
//----------------------------------------------------------------------------
STDMETHODIMP CImgElement::get_fileModifiedDate(BSTR* pFileModifiedDate)
{
    HRESULT hr = S_OK;
    TCHAR* pchCachedFile = NULL;

    if(pFileModifiedDate == NULL)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pFileModifiedDate = NULL;

    Assert(_pImage);
    hr = _pImage->GetFile(&pchCachedFile);
    if(!hr && pchCachedFile)
    {
        WIN32_FIND_DATA wfd;
        HANDLE hFind = FindFirstFile(pchCachedFile, &wfd);
        if (hFind != INVALID_HANDLE_VALUE)
        {
            FindClose(hFind);
            // we always return the local time in a fixed format mm/dd/yyyy to make it possible to parse
            // FALSE means we do not want the time
            hr = ConvertDateTimeToString(wfd.ftLastWriteTime, pFileModifiedDate, FALSE);
        }
    }

Cleanup:
    MemFreeString(pchCachedFile);
    RRETURN(SetErrorInfo(hr));
}

// BUGBUG (lmollico): get_fileUpdatedDate won't work if src=file://image
//+---------------------------------------------------------------------------
//
// Member: get_fileUpdatedDate
//
//----------------------------------------------------------------------------
STDMETHODIMP CImgElement::get_fileUpdatedDate(BSTR* pFileUpdatedDate)
{
    HRESULT hr = S_OK;
    TCHAR   cBuf[pdlUrlLen];
    TCHAR*  pszUrl = cBuf;

    if(pFileUpdatedDate == NULL)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pFileUpdatedDate = NULL;

    hr = Doc()->ExpandUrl(GetAAsrc(), ARRAYSIZE(cBuf), pszUrl, this);
    if(hr)
    {
        goto Cleanup;
    }

    if(pszUrl)
    {
        BYTE                        buf[MAX_CACHE_ENTRY_INFO_SIZE];
        INTERNET_CACHE_ENTRY_INFO*  pInfo = (INTERNET_CACHE_ENTRY_INFO*)buf;
        DWORD                       cInfo = sizeof(buf);

        if(RetrieveUrlCacheEntryFile(pszUrl, pInfo, &cInfo, 0))
        {
            // we always return the local time in a fixed format mm/dd/yyyy to make it possible to parse
            // FALSE means we do not want the time
            hr = ConvertDateTimeToString(pInfo->LastModifiedTime, pFileUpdatedDate, FALSE);
            DoUnlockUrlCacheEntryFile(pszUrl, 0);
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
// Member: get_href
//
//----------------------------------------------------------------------------
STDMETHODIMP CImgElement::get_href(BSTR* pHref)
{
    HRESULT hr = S_OK;
    LPCTSTR pchUrl = NULL;

    if(pHref == NULL)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pHref = NULL;

    if(_pImage && _pImage->_pBitsCtx)
    {
        pchUrl = _pImage->_pBitsCtx->GetUrl();
    }
    else if(_pImage && _pImage->_pImgCtx)
    {
        pchUrl = _pImage->_pImgCtx->GetUrl();
    }

    if(pchUrl)
    {
        *pHref = SysAllocString(pchUrl);
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//
// Member: get_protocol
//
//----------------------------------------------------------------------------
extern TCHAR* ProtocolFriendlyName(TCHAR* szUrl);

STDMETHODIMP CImgElement::get_protocol(BSTR* pProtocol)
{
    HRESULT hr      = S_OK;
    LPCTSTR pchUrl  = NULL;
    TCHAR*  pResult;

    if(pProtocol == NULL)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pProtocol = NULL;

    if(_pImage && _pImage->_pBitsCtx)
    {
        pchUrl = _pImage->_pBitsCtx->GetUrl();
    }
    else if(_pImage && _pImage->_pImgCtx)
    {
        pchUrl = _pImage->_pImgCtx->GetUrl();
    }

    if(pchUrl)
    {
        pResult = ProtocolFriendlyName((TCHAR*)pchUrl);
        if(pResult)
        {
            int z = (_tcsncmp(pResult, 4, _T("URL:"), -1)==0) ? (4) : (0);
            * pProtocol = SysAllocString(pResult+z);
            SysFreeString(pResult);
        }
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
// Member: get_nameProp
//
//-----------------------------------------------------------------------------
STDMETHODIMP CImgElement::get_nameProp(BSTR* pName)
{
    *pName = NULL;

    TCHAR   cBuf[pdlUrlLen];
    TCHAR*  pszUrl  = cBuf;
    TCHAR*  pszName = NULL;
    HRESULT hr      = Doc()->ExpandUrl(GetAAsrc(), ARRAYSIZE(cBuf), pszUrl, this);
    if(!hr)
    {
        pszName = _tcsrchr(pszUrl, _T('/'));
        if(!pszName)
        {
            pszName = pszUrl;
        }
        else
        {
            pszName ++;
        }

        *pName = SysAllocString(pszName);
    }

    RRETURN(SetErrorInfo(hr));
}

//+----------------------------------------------------------------------------
//
//  Method:     CImgElem::GetSubdivisionCount
//
//  Synopsis:   returns the count of subdivisions
//
//-----------------------------------------------------------------------------
HRESULT CImgElement::GetSubDivisionCount(long* pc)
{
    HRESULT hr = S_OK;

    hr = super::GetSubDivisionCount(pc);
    goto Cleanup;

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Method:     CImgElem::GetSubdivisionTabs
//
//  Synopsis:   returns the subdivisions tabindices
//
//-----------------------------------------------------------------------------
HRESULT CImgElement::GetSubDivisionTabs(long* pTabs, long c)
{
    HRESULT hr = S_OK;

    if(!c)
    {
        goto Cleanup;
    }

    goto Cleanup;

Cleanup:
    RRETURN(hr);
}

//+----------------------------------------------------------------------------
//
//  Method:     CImgElem::SubDivisionFromPt
//
//  Synopsis:   returns the subdivisions tabindices
//
//-----------------------------------------------------------------------------
HRESULT CImgElement::SubDivisionFromPt(POINT pt, long* plSub)
{
    HRESULT     hr = S_OK;

    Assert(GetFirstBranch());

    *plSub = -1;

    hr = E_FAIL;
    goto Cleanup;

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgElement::GetFocusShape
//
//  Synopsis:   Returns the shape of the focus outline that needs to be drawn
//              when this element has focus. This function creates a new
//              CShape-derived object. It is the caller's responsibility to
//              release it.
//
//----------------------------------------------------------------------------
HRESULT CImgElement::GetFocusShape(long lSubDivision, CDocInfo* pdci, CShape** ppShape)
{
    HRESULT             hr = S_FALSE;
    const CParaFormat*  pPF = GetFirstBranch()->GetParaFormat();
    BOOL                fRTL = pPF->HasRTL(FALSE);
    CLayout*            pLayout = Layout();
    CSize               size = _afxGlobalData._Zero.size;
    if(fRTL)
    {
        // we are only interested in adjusting the x positioning
        // for RTL direction.
        CRect rcClient;
        pLayout->GetClientRect(&rcClient);
        size.cx = rcClient.Width();
    }

    *ppShape = NULL;

    Assert(GetAAtabIndex() >= 0);

    hr = super::GetFocusShape(lSubDivision, pdci, ppShape);
    goto Cleanup;

Cleanup:
    RRETURN1(hr, S_FALSE);
}

//+----------------------------------------------------------------------------
//
//  Function:   DoSubDivisionEvents
//
//  Synopsis:   Fire the specified event on the given subdivision.
//
//  Arguments:  [dispidEvent]   -- dispid of the event to fire.
//              [dispidProp]    -- dispid of prop containing event func.
//              [pvb]           -- Boolean return value
//              [pbTypes]       -- Pointer to array giving the types of parms
//              [...]           -- Parameters
//
//-----------------------------------------------------------------------------
HRESULT CImgElement::DoSubDivisionEvents(
                                 long       lSubDivision,
                                 DISPID     dispidEvent,
                                 DISPID     dispidProp,
                                 VARIANT*   pvb,
                                 BYTE*      pbTypes, ...)
{
    if(lSubDivision < 0)
    {
        return S_OK;
    }

    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgElement::ShowTooltip
//
//  Synopsis:   Show the tooltip for this element.
//
//----------------------------------------------------------------------------
HRESULT CImgElement::ShowTooltip(CMessage* pmsg, POINT pt)
{
    HRESULT         hr = S_OK;
    CString         strTitle;
    BOOL            fRTL = FALSE;
    CDocument*      pDoc = Doc();

    if(pDoc->_pInPlace == NULL)
    {
        goto Cleanup;
    }

    // check to see if tooltip should display the title property
    hr = super::ShowTooltip(pmsg, pt);
    if(hr == S_OK)
    {
        goto Cleanup;
    }

    Assert(_pImage);
    hr = _pImage->ShowTooltip(pmsg, pt);

Cleanup:
    RRETURN1 (hr, S_FALSE);
}

HRESULT CImgElement::GetHeight(long* pl)
{
    VARIANT v;
    HRESULT hr;

    hr = s_propdescCImgElementheight.a.HandleUnitValueProperty(
        HANDLEPROP_VALUE|(PROPTYPE_VARIANT<<16),
        &v,
        this,
        CVOID_CAST(GetAttrArray()));
    if(hr)
    {
        goto Cleanup;
    }

    Assert(V_VT(&v) == VT_I4);
    Assert(pl);

    *pl = V_I4(&v);

Cleanup:
    RRETURN(hr);
}

HRESULT CImgElement::putHeight(long l)
{
    VARIANT v;

    if(l < 0)
    {
        l = 0;
    }

    V_VT(&v) = VT_I4;
    V_I4(&v) = l;

    RRETURN(s_propdescCImgElementheight.a.HandleUnitValueProperty(
        HANDLEPROP_SET|HANDLEPROP_AUTOMATION|HANDLEPROP_DONTVALIDATE|(PROPTYPE_VARIANT<<16),
        &v,
        this,
        CVOID_CAST(GetAttrArray())));
}

HRESULT CImgElement::GetWidth(long* pl)
{
    VARIANT v;
    HRESULT hr;

    hr = s_propdescCImgElementwidth.a.HandleUnitValueProperty(
        HANDLEPROP_VALUE | (PROPTYPE_VARIANT << 16),
        &v,
        this,
        CVOID_CAST(GetAttrArray()));
    if(hr)
    {
        goto Cleanup;
    }

    Assert(V_VT(&v) == VT_I4);
    Assert(pl);

    *pl = V_I4(&v);

Cleanup:
    RRETURN(hr);
}

HRESULT CImgElement::putWidth(long l)
{
    VARIANT v;

    if(l < 0)
    {
        l = 0;
    }

    V_VT(&v) = VT_I4;
    V_I4(&v) = l;

    RRETURN(s_propdescCImgElementwidth.a.HandleUnitValueProperty(
        HANDLEPROP_SET|HANDLEPROP_DONTVALIDATE|HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v,
        this,
        CVOID_CAST(GetAttrArray())));
}

STDMETHODIMP CImgElement::put_height(long l)
{
    RRETURN(SetErrorInfoPSet(putHeight(l), DISPID_CImgElement_width));
}

STDMETHODIMP CImgElement::get_height(long* p)
{
    Assert(_pImage);
    RRETURN(_pImage->get_height(p));
}

STDMETHODIMP CImgElement::put_width(long l)
{
    RRETURN(SetErrorInfoPSet(putWidth(l), DISPID_CImgElement_width));
}

STDMETHODIMP CImgElement::get_width(long* p)
{
    Assert(_pImage);
    RRETURN(_pImage->get_width(p));
}

STDMETHODIMP CImgElement::get_src(BSTR* pstrFullSrc)
{
    HRESULT hr;
    Assert(_pImage);
    hr = _pImage->get_src(pstrFullSrc);
    RRETURN(SetErrorInfo(hr));
}

//+------------------------------------------------------------------
//
//  member : put_src
//
//  sysnopsis : impementation of the interface src property set
//          since this is a URL property we want the crlf striped out
//
//-------------------------------------------------------------------
STDMETHODIMP CImgElement::put_src(BSTR bstrSrc)
{
    RRETURN(SetErrorInfo(s_propdescCImgElementsrc.b.SetUrlProperty(bstrSrc,
        this, (CVoid*)(void*)(GetAttrArray()))));
}

//+----------------------------------------------------------------------------
//
// Methods:     get/set_hspace
//
// Synopsis:    hspace for aligned images is 3 pixels by default, so we need
//              a method to identify if a default value is specified.
//
//-----------------------------------------------------------------------------
STDMETHODIMP CImgElement::put_hspace(long v)
{
    return s_propdescCImgElementhspace.b.SetNumberProperty(v, this, CVOID_CAST(GetAttrArray()));
}

STDMETHODIMP CImgElement::get_hspace(long * p)
{
    HRESULT hr = s_propdescCImgElementhspace.b.GetNumberProperty(p, this, CVOID_CAST(GetAttrArray()));

    if(!hr)
    {
        *p = *p==-1 ? 0 : *p;
    }

    return hr;
}

//+----------------------------------------------------------------------------
//
//  Member : [get_/put_] onload
//
//  synopsis : store in this element's propdesc
//
//+----------------------------------------------------------------------------
HRESULT CImgElement:: put_onload(VARIANT v)
{
    HRESULT hr = s_propdescCImgElementonload .a.HandleCodeProperty(
        HANDLEPROP_SET|HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        &v,
        this,
        CVOID_CAST(GetAttrArray()));

    RRETURN( SetErrorInfo( hr ));
}

HRESULT CImgElement:: get_onload(VARIANT* p)
{
    HRESULT hr = s_propdescCImgElementonload.a.HandleCodeProperty(
        HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16),
        p,
        this,
        CVOID_CAST(GetAttrArray()));

    RRETURN(SetErrorInfo(hr));
}

//+---------------------------------------------------------------------------
//  Member :    CImgElement::IsHSpaceDefined
//
//  Synopsis:   if hspace is defined on the image
//
//+---------------------------------------------------------------------------
BOOL CImgElement::IsHSpaceDefined() const
{
    DWORD v;
    CAttrArray::FindSimple(*GetAttrArray(), &s_propdescCImgElementhspace.a, &v);

    return v != -1;
}

HRESULT CImgElement::get_readyState(BSTR* p)
{
    HRESULT hr = S_OK;

    if(!p)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    hr = s_enumdeschtmlReadyState.StringFromEnum(_pImage->_readyStateFired, p);

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CImgElement::get_readyState(VARIANT* pVarRes)
{
    HRESULT hr = S_OK;

    if(!pVarRes)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    hr = get_readyState(&V_BSTR(pVarRes));
    if(!hr)
    {
        V_VT(pVarRes) = VT_BSTR;
    }

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

HRESULT CImgElement::get_readyStateValue(long* plRetValue)
{
    HRESULT hr = S_OK;

    if(!plRetValue)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    *plRetValue = _pImage->_readyStateFired;

Cleanup:
    RRETURN(SetErrorInfo(hr));
}

void CImgElement::Passivate()
{
    if(_pImage)
    {
        _pImage->Passivate();
        delete _pImage;
        _pImage = NULL;
    }
    super::Passivate();
}