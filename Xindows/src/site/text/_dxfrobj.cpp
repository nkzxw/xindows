
#include "stdafx.h"
#include "_dxfrobj.h"

#include <ShlObj.h>

// If you change g_rgFETC[], change g_rgDOI[] and enum FETCINDEX and CFETC in
// _dxfrobj.h accordingly, and register nonstandard clipboard formats in
// RegisterFETCs(). Order entries in order of most desirable to least, e.g.,
// RTF in front of plain text.
FORMATETC g_rgFETC[] =
{
    {0,                 NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL}, // CF_HTML
    {0,                 NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL}, // CF_RTF
    {CF_UNICODETEXT,    NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL},
    {CF_TEXT,           NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL},
    {0,                 NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL}, // CF_RTFASTEXT
    {0,                 NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL}, // CF_FILEDESCRIPTORA
    {0,                 NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL}, // CF_FILEDESCRIPTORW
    {0,                 NULL, DVASPECT_CONTENT,  0, TYMED_HGLOBAL}, // CF_FILECONTENTS
    {0,                 NULL, DVASPECT_CONTENT,  0, TYMED_HGLOBAL}, // CF_UNIFORMRESOURCELOCATOR
};

int CFETC = ARRAYSIZE(g_rgFETC);

//TODO v-richa check the added members for correctness
DWORD g_rgDOI[] =
{
    DOI_CANPASTEPLAIN,
    DOI_CANPASTEPLAIN,
    DOI_CANPASTEPLAIN,
    DOI_CANPASTEPLAIN,
    DOI_CANPASTEPLAIN,
    DOI_NONE,
    DOI_NONE,
    DOI_NONE,
    DOI_NONE,
};

/*
*  RegisterFETCs()
*
*  @func
*      Register nonstandard format ETCs.  Called when DLL is loaded
*
*  @todo
*      Register other RTF formats (and add places for them in g_rgFETC[])
*/
void RegisterFETCs()
{
    if(!g_rgFETC[iHTML].cfFormat)
    {
        g_rgFETC[iHTML].cfFormat
            = (CLIPFORMAT)RegisterClipboardFormatA("HTML Format");
    }

    if(!g_rgFETC[iRtfFETC].cfFormat)
    {
        g_rgFETC[iRtfFETC].cfFormat
            = (CLIPFORMAT)RegisterClipboardFormatA("Rich Text Format");
    }

    if(!g_rgFETC[iRtfAsTextFETC].cfFormat)
    {
        g_rgFETC[iRtfAsTextFETC].cfFormat
            = (CLIPFORMAT)RegisterClipboardFormatA("RTF As Text");
    }

    if(!g_rgFETC[iFileDescA].cfFormat)
    {
        g_rgFETC[iFileDescA].cfFormat
            = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_FILEDESCRIPTORA);
    }

    if(!g_rgFETC[iFileDescW].cfFormat)
    {
        g_rgFETC[iFileDescW].cfFormat
            = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_FILEDESCRIPTORW);
    }

    if(!g_rgFETC[iFileContents].cfFormat)
    {
        g_rgFETC[iFileContents].cfFormat
            = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_FILECONTENTS);
    }

    if(!g_rgFETC[iUniformResourceLocator].cfFormat)
    {
        g_rgFETC[iUniformResourceLocator].cfFormat
            = (CLIPFORMAT)RegisterClipboardFormat(CFSTR_SHELLURL);
    }
}



/*
*  CTextXBag::EnumFormatEtc (dwDirection, ppenumFormatEtc)
*
*  @mfunc
*      returns an enumerator which lists all of the available formats in
*      this data transfer object
*
*  @rdesc
*      HRESULT
*
*  @devnote
*      we have no 'set' formats for this object
*/
STDMETHODIMP CTextXBag::EnumFormatEtc(
          DWORD dwDirection,                  // @parm DATADIR_GET/SET
          IEnumFORMATETC** ppenumFormatEtc)   // @parm out parm for enum FETC interface
{
    HRESULT hr = NOERROR;

    *ppenumFormatEtc = NULL;

    if(_pLinkDataObj && (_cTotal<_cFormatMax))
    {
        _prgFormats[_cTotal++] = g_rgFETC[iFileDescA];
        _prgFormats[_cTotal++] = g_rgFETC[iFileDescW];
        _prgFormats[_cTotal++] = g_rgFETC[iFileContents];
        _prgFormats[_cTotal++] = g_rgFETC[iUniformResourceLocator];
    }

    if(dwDirection == DATADIR_GET)
    {
        hr = CEnumFormatEtc::Create(_prgFormats, _cTotal, ppenumFormatEtc);
    }

    return hr;
}

/*
*  CTextXBag::GetData (pformatetcIn, pmedium)
*
*  @mfunc
*      retrieves data of the specified format
*
*  @rdesc
*      HRESULT
*
*  @devnote
*      The three formats currently supported are CF_UNICODETEXT on
*      an hglobal, CF_TEXT on an hglobal, and CF_RTF on an hglobal
*
*  @todo (alexgo): handle all of the other formats as well
*/
STDMETHODIMP CTextXBag::GetData(FORMATETC* pformatetcIn, STGMEDIUM* pmedium)
{
    CLIPFORMAT  cf = pformatetcIn->cfFormat;
    HRESULT     hr = DV_E_FORMATETC;
    HGLOBAL     hGlobal;

    if(!(pformatetcIn->tymed & TYMED_HGLOBAL))
    {
        goto Cleanup;
    }

    memset(pmedium, '\0', sizeof(STGMEDIUM));
    pmedium->tymed = TYMED_NULL;

    if(cf == cf_HTML)
    {
        hGlobal = _hCFHTMLText;
    }
    else if(cf == CF_UNICODETEXT)
    {
        hGlobal = _hUnicodeText;
    }
    else if(cf == CF_TEXT)
    {
        hGlobal = _hText;
    }
    else if(cf==cf_RTF || cf==cf_RTFASTEXT)
    {
        hGlobal = _hRTFText;
    }
    else
    {
        goto Cleanup;
    }

    if(hGlobal)
    {
        pmedium->tymed   = TYMED_HGLOBAL;
        pmedium->hGlobal = DuplicateHGlobal(hGlobal);
        if(!pmedium->hGlobal)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }

        hr = S_OK;
    }

Cleanup:
    if(hr && _pLinkDataObj)
    {
        hr = _pLinkDataObj->GetData(pformatetcIn, pmedium);
    }
    RRETURN(hr);
}

/*
*  CTextXBag::QueryGetData (pformatetc)
*
*  @mfunc
*      Queries whether the given format is available in this data object
*
*  @rdesc
*      HRESULT
*/
STDMETHODIMP CTextXBag::QueryGetData(FORMATETC* pformatetc) // @parm FETC to look for
{
    DWORD cFETC = _cTotal;
    CLIPFORMAT cf = pformatetc->cfFormat;

    while(cFETC--) // Maybe faster to search from start
    {
        if(cf==_prgFormats[cFETC].cfFormat && pformatetc->tymed&TYMED_HGLOBAL)
        {
            // Check for RTF handle even if we claim to support the format
            if(_hRTFText || (cf!=cf_RTF && cf!=cf_RTFASTEXT))
            {
                return NOERROR;
            }
        }
    }

    if(_pLinkDataObj)
    {
        return _pLinkDataObj->QueryGetData(pformatetc);
    }

    return DV_E_FORMATETC;
}

STDMETHODIMP CTextXBag::SetData(LPFORMATETC pformatetc, STGMEDIUM FAR* pmedium, BOOL fRelease)
{
    _hUnicodeText = pmedium->hGlobal;
    if(!pmedium->hGlobal)
    {
        _hText = NULL;
    }

    return S_OK;
}

STDMETHODIMP CTextXBag::QueryInterface(REFIID iid, LPVOID* ppv)
{
    if(_pSelDragDropSrcInfo && (iid==IID_IUnknown))
    {
        return _pSelDragDropSrcInfo->QueryInterface(iid, ppv);
    }
    else
    {
        return super::QueryInterface(iid, ppv);
    }
}

STDMETHODIMP_(ULONG) CTextXBag::AddRef()
{
    return _pSelDragDropSrcInfo?_pSelDragDropSrcInfo->AddRef():super::AddRef();
}

STDMETHODIMP_(ULONG) CTextXBag::Release()
{
    return _pSelDragDropSrcInfo?_pSelDragDropSrcInfo->Release():super::Release();
}

//+------------------------------------------------------------------------
//
//  Member:     CTextXBag::Create
//
//  Synopsis:   Static creator of text xbags
//
//  Arguments:  pMarkup         The markup that owns the selection
//              fSupportsHTML   Can the selection be treated as HTML?
//              ppRange         Array of ptrs to ranges
//              cRange          Number of items in above array
//              ppTextXBag      Returned xbag.
//
//-------------------------------------------------------------------------
HRESULT CTextXBag::Create(
          CMarkup*              pMarkup,
          BOOL                  fSupportsHTML,
          ISegmentList*         pSegmentList,
          BOOL                  fDragDrop,
          CTextXBag**           ppTextXBag,
          CSelDragDropSrcInfo*  pSelDragDropSrcInfo/*=NULL*/)
{
    HRESULT hr;

    CTextXBag* pTextXBag = new CTextXBag();

    if(!pTextXBag)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pTextXBag->_pSelDragDropSrcInfo = pSelDragDropSrcInfo;

    if(fDragDrop)
    {
        pTextXBag->_pDoc = pMarkup->Doc();
    }

    hr = pTextXBag->SetKeyState();
    if(hr)
    {
        goto Error;
    }

    hr = pTextXBag->FillWithFormats(pMarkup, fSupportsHTML, pSegmentList);
    if(hr)
    {
        goto Error;
    }

Cleanup:
    *ppTextXBag = pTextXBag;

    RRETURN(hr);

Error:
    delete pTextXBag;
    pTextXBag = NULL;
    goto Cleanup;
}

//+---------------------------------------------------------------------------
//
//  Member:     CTextXBag::CTextXBag
//
//  Synopsis:   Private ctor
//
//----------------------------------------------------------------------------
CTextXBag::CTextXBag()
{
    _ulRefs       = 1;
    _cTotal       = CFETC;
    _cFormatMax   = 1;
    _prgFormats   = g_rgFETC;
    _hText        = NULL;
    _hUnicodeText = NULL;
    _hRTFText     = NULL;
    _hCFHTMLText  = NULL;
    _pLinkDataObj = NULL;
}

//+---------------------------------------------------------------------------
//
//  Member:     CTextXBag::~CTextXBag
//
//  Synopsis:   Private dtor
//
//----------------------------------------------------------------------------
CTextXBag::~CTextXBag()
{
    if(_prgFormats && _prgFormats!=g_rgFETC)
    {
        delete[] _prgFormats;
    }

    if(_hText)
    {
        GlobalFree(_hText);
    }

    if(_hUnicodeText)
    {
        GlobalFree(_hUnicodeText);
    }

    if(_hRTFText)
    {
        GlobalFree(_hRTFText);
    }

    if(_hCFHTMLText)
    {
        GlobalFree(_hCFHTMLText);
    }

    TLS(pDataClip) = NULL;
}

//+------------------------------------------------------------------------/
//
//  Member:     CTextXBag::SetKeyState
//
//  Synopsis:   Sets the _dwButton member of the CTextXBag
//
//-------------------------------------------------------------------------
HRESULT CTextXBag::SetKeyState()
{
    static int vk[] =
    {
        VK_LBUTTON, // MK_LBUTTON = 0x0001
        VK_RBUTTON, // MK_RBUTTON = 0x0002
        VK_SHIFT,   // MK_SHIFT   = 0x0004
        VK_CONTROL, // MK_CONTROL = 0x0008
        VK_MBUTTON, // MK_MBUTTON = 0x0010
        VK_MENU,    // MK_ALT     = 0x0020
    };

    int     i;
    DWORD   dwKeyState = 0;

    for(i=0; i<ARRAYSIZE(vk); i++)
    {
        if(GetKeyState(vk[i]) & 0x8000)
        {
            dwKeyState |= (1 << i);
        }
    }

    _dwButton = dwKeyState & (MK_LBUTTON|MK_MBUTTON|MK_RBUTTON);

    return S_OK;
}

//+------------------------------------------------------------------------/
//
//  Member:     CTextXBag::FillWithFormats
//
//  Synopsis:   Fills the text bag with the formats it supports
//
//-------------------------------------------------------------------------
HRESULT CTextXBag::FillWithFormats(
           CMarkup*      pMarkup,
           BOOL          fSupportsHTML,
           ISegmentList* pSegmentList)
{
    typedef HRESULT (CTextXBag::*FnSet)(CMarkup*, BOOL, ISegmentList*);

    static FnSet aFnSet[] =
    {
        &CTextXBag::SetText,
        &CTextXBag::SetUnicodeText,
        &CTextXBag::SetCFHTMLText,
    };

    HRESULT hr = S_OK;
    int     n;

    // Allocate our _prgFormats array
    _cFormatMax = ARRAYSIZE(aFnSet) + 3;
    _cTotal     = 0;
    _prgFormats = new FORMATETC[_cFormatMax];
    if(NULL == _prgFormats)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    for(n=0; n<ARRAYSIZE(aFnSet); ++n)
    {
        // If one format fails to copy, do not abort all formats.
        CALL_METHOD(this, aFnSet[n], (pMarkup, fSupportsHTML, pSegmentList));
    }

Cleanup:
    RRETURN(hr);
}

HRESULT  CTextXBag::GetDataObjectInfo(
          IDataObject*  pdo,        // @parm data object to get info on
          DWORD*        pDOIFlags)  // @parm out parm for info
{
    DWORD       i;
    FORMATETC*  pfetc = g_rgFETC;

    for(i=0; i<DWORD(CFETC); i++,pfetc++)
    {
        if(pdo->QueryGetData(pfetc) == NOERROR)
        {
            *pDOIFlags |= g_rgDOI[i];
        }
    }
    return NOERROR;
}

//+---------------------------------------------------------------------------
//
//  Member:     CTextXBag::GetHTMLText
//
//  Synopsis:   Converts the SegmentList into text stored in an hglobal
//
//----------------------------------------------------------------------------
HRESULT CTextXBag::GetHTMLText(
           HGLOBAL*         phGlobal, 
           ISegmentList*    pSegmentList,
           CMarkup*         pMarkup, 
           DWORD            dwSaveHtmlMode,
           CODEPAGE         codepage, 
           DWORD            dwStrWrBuffFlags)
{
    HGLOBAL     hGlobal  = NULL;   // Global memory handle
    LPSTREAM    pIStream = NULL;   // IStream pointer
    HRESULT     hr;
    int         iSegmentCount;
    CDocument*  pDoc;

    // Do the prep work
    hr = CreateStreamOnHGlobal(NULL, FALSE, &pIStream);
    if(hr)
    {
        goto Error;
    }

    pDoc = pMarkup->Doc();
    Assert(pDoc);


    // Use a scope to clean up the StreamWriteBuff
    {
        CMarkupPointer      mpStart( pDoc ), mpEnd( pDoc );
        CStreamWriteBuff    StreamWriteBuff(pIStream, codepage);
        long                i;

        StreamWriteBuff.SetFlags(dwStrWrBuffFlags);

        hr = pSegmentList->GetSegmentCount(&iSegmentCount, NULL);
        if(hr)
        {
            goto Cleanup;
        }

        // Save the segments using the range saver
        for(i=0; i<iSegmentCount; i++)
        {
            hr = pSegmentList->MovePointersToSegment(i, &mpStart, &mpEnd);        

            CRangeSaver rs(&mpStart, &mpEnd, dwSaveHtmlMode, &StreamWriteBuff, mpStart.Markup());

            hr = rs.Save();
            if(hr)
            {
                goto Error;
            }
        }

        StreamWriteBuff.Terminate(); // appends a null character
    }

    // Wrap it up
    hr = GetHGlobalFromStream(pIStream, &hGlobal);
    if(hr)
    {
        goto Error;
    }

Cleanup:
    ReleaseInterface(pIStream);
    *phGlobal = hGlobal;
    RRETURN(hr);

Error:
    if(hGlobal)
    {
        GlobalFree(hGlobal);
        hGlobal = NULL;
    }
    goto Cleanup;
}

//+------------------------------------------------------------------------
//
//  Member:     CTextXBag::SetTextHelper
//
//  Synopsis:   Gets text in a variety of formats
//
//  Argument:   pTxtSite:           The text site under which to get the text from
//              ppRanges:           Text ranges to save text from
//              cRanges:            Count of ppRanges
//              dwSaveHtmlFlags:    format to save in
//              cp:                 codepage to save in
//              dwStmWrBuffFlags:   stream write buffer flags
//              phGlobalText:       hGlobal to get back
//              iFETCIndex:         _prgFormat index, or -1 to not set it
//
//-------------------------------------------------------------------------
HRESULT CTextXBag::SetTextHelper(
         CMarkup*       pMarkup,
         ISegmentList*  pSegmentList,
         DWORD          dwSaveHtmlFlags,
         CODEPAGE       cp,
         DWORD          dwStmWrBuffFlags,
         HGLOBAL*       phGlobalText,
         int            iFETCIndex)
{
    HRESULT hr;
    HGLOBAL hText = NULL;

    Assert(_cTotal < _cFormatMax);

    // Make sure not to crash if we are out of space for this format.
    if(_cTotal >= _cFormatMax)
    {
        return S_OK;
    }

    {
        hr = GetHTMLText( 
            &hText, pSegmentList, pMarkup, 
            dwSaveHtmlFlags, cp, dwStmWrBuffFlags);
        if(hr)
        {
            goto Error;
        }
    }

    // if the text length is zero, pretend as if the format is
    // unavailable (see bug #52407)
    if(iUnicodeFETC==iFETCIndex || iAnsiFETC==iFETCIndex)
    {
        BOOL fEmpty;

        Assert(hText);

        LPVOID pText= GlobalLock(hText);
        if(iAnsiFETC == iFETCIndex) 
        {
            fEmpty= *((char*)pText) == 0;
        }
        else
        {
            // Please don't use strlen on unicode strings.
            fEmpty= *((TCHAR*)pText) == 0;
        }
        GlobalUnlock(hText);
        if(fEmpty)
        {
            GlobalFree(hText);
            hText = NULL;
            goto Cleanup;
        }
    }

    if(iFETCIndex != -1)
    {
        _prgFormats[_cTotal++] = g_rgFETC[iFETCIndex];
    }

Cleanup:
    *phGlobalText = hText;
    RRETURN(hr);

Error:
    if(hText)
    {
        GlobalFree(hText);
        hText = NULL;
    }

    goto Cleanup;
}

//+------------------------------------------------------------------------
//
//  Member:     CTextXBag::SetText
//
//  Synopsis:   Gets ansi plaintext in CP_ACP
//
//-------------------------------------------------------------------------
HRESULT CTextXBag::SetText(
       CMarkup*         pMarkup,
       BOOL             fSupportsHTML,
       ISegmentList*    pSegmentList)
{
    DWORD dwFlags = WBF_SAVE_PLAINTEXT|WBF_NO_WRAP|WBF_FORMATTED_PLAINTEXT;

    if(!fSupportsHTML)
    {
        dwFlags |= WBF_KEEP_BREAKS;
    }

    RRETURN(SetTextHelper(pMarkup, pSegmentList,
        RSF_SELECTION|RSF_NO_ENTITIZE_UNKNOWN,
        _afxGlobalData._cpDefault, dwFlags, &_hText, iAnsiFETC));
}

//+------------------------------------------------------------------------
//
//  Member:     CTextXBag::SetUnicodeText
//
//  Synopsis:   Gets unicode plaintext
//
//-------------------------------------------------------------------------
HRESULT CTextXBag::SetUnicodeText(
          CMarkup*      pMarkup,
          BOOL          fSupportsHTML,
          ISegmentList* pSegmentList)
{
    DWORD dwFlags = WBF_SAVE_PLAINTEXT|WBF_NO_WRAP|WBF_FORMATTED_PLAINTEXT;

    if(!fSupportsHTML)
    {
        dwFlags |= WBF_KEEP_BREAKS;
    }

    RRETURN(SetTextHelper(pMarkup, pSegmentList,
        RSF_SELECTION, CP_UCS_2,
        dwFlags, &_hUnicodeText, iUnicodeFETC));
}

//+------------------------------------------------------------------------
//
//  Member:     CTextXBag::SetCFHTMLText
//
//  Synopsis:   Gets HTML with CF_HTML header in UTF-8
//
//-------------------------------------------------------------------------
HRESULT CTextXBag::SetCFHTMLText(
         CMarkup*       pMarkup,
         BOOL           fSupportsHTML,
         ISegmentList*  pSegmentList)
{
    HRESULT hr = S_OK;

    if(fSupportsHTML)
    {
        hr = SetTextHelper(pMarkup, pSegmentList,
            RSF_CFHTML, CP_UTF_8, WBF_NO_NAMED_ENTITIES, &_hCFHTMLText, iHTML);
    }

    RRETURN(hr);
}
