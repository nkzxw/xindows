
#include "stdafx.h"
#include "ImgHelper.h"

#include <ShlObj.h>

extern BOOL g_fScreenReader;
extern BYTE g_bJGJitState;

struct CACHE_ENTRY
{
    DWORD dwHash;
    WORD  wWidth;
    WORD  wHeight;
};

struct CACHE_FILE
{
    DWORD dwMagic;
    DWORD dwVersion;
    DWORD dwHit;
    DWORD dwMiss;
    CACHE_ENTRY aCacheEntry[2046];
};

static HANDLE s_hCacheFileMapping;
static CACHE_FILE* s_pCacheFile;
static BOOL   s_fInitializationTried;

#define MAGIC           0xCAC8EF17
#define OBJ_NAME(x)     "MSIMGSIZECache" #x
#define FILE_NAME       "MSIMGSIZ.DAT"

//+------------------------------------------------------------------------
//
//  VERSION     Increment this number with each file format change.
//
//-------------------------------------------------------------------------
#define VERSION 1

//+------------------------------------------------------------------------
//
//  Function:   InitImageSizeCache
//
//  Synopsis:   Open the image size cache file.  Create and initialize it
//              if required. 
//
//-------------------------------------------------------------------------
BOOL InitImageSizeCache()
{
    BOOL   fInitialize = FALSE;
    HANDLE hMutex = NULL;
    HANDLE hFile = INVALID_HANDLE_VALUE;
    CACHE_FILE* pCacheFile;

    if(s_fInitializationTried)
    {
        return s_pCacheFile != NULL;
    }

    s_fInitializationTried = TRUE;

    hMutex = CreateMutexA(NULL, FALSE, OBJ_NAME(Mutex));
    if(!hMutex)
    {
        goto Cleanup;
    }

    if(WaitForSingleObject(hMutex, 4000) != WAIT_OBJECT_0)
    {
        goto Cleanup;
    }

    // Did another thread initialize when we were not looking?
    if(s_pCacheFile)
    {
        goto Cleanup;
    }

    // Open or create new file mapping.
    s_hCacheFileMapping = OpenFileMappingA(FILE_MAP_ALL_ACCESS, FALSE, OBJ_NAME(Map));
    if(s_hCacheFileMapping)
    {
        BOOL fVersionMismatch = FALSE;
        pCacheFile = (CACHE_FILE*)MapViewOfFile(
            s_hCacheFileMapping, 
            FILE_MAP_ALL_ACCESS,            // access
            0,                              // offset low
            0,                              // offset high
            sizeof(CACHE_FILE));            // number of bytes to map

        if(!pCacheFile)
        {
            goto Cleanup;
        }

        // Punt if some other version of Trident has the file open.
        __try
        {
            fVersionMismatch = pCacheFile->dwMagic!=MAGIC || pCacheFile->dwVersion!=VERSION;
        }
        __except(EXCEPTION_EXECUTE_HANDLER)
        {
            goto Cleanup;
        }

        if(fVersionMismatch)
        {
            goto Cleanup;
        }
    }
    else
    {
        // File mapping does not exist.  Open or create file.
        HRESULT hr = S_OK;
        char achBuf[MAX_PATH+sizeof(FILE_NAME)+1];
        int  cch;
        hr = SHGetFolderPathA(NULL, CSIDL_APPDATA, NULL, 0, achBuf);
        if(hr)
        {
            goto Cleanup;
        }

        strcat(achBuf, "\\Microsoft");
        if(!PathFileExistsA(achBuf) && !CreateDirectoryA(achBuf, NULL))
        {
            goto Cleanup;
        }
        strcat(achBuf, "\\Internet Explorer");
        if(!PathFileExistsA(achBuf) && !CreateDirectoryA(achBuf, NULL))
        {
            goto Cleanup;
        }

        cch = strlen(achBuf);

        hFile = CreateFileA(
            achBuf,
            GENERIC_READ|GENERIC_WRITE,         // access mode
            FILE_SHARE_READ|FILE_SHARE_WRITE,   // share mode
            NULL,                               // security
            OPEN_ALWAYS,                        // disposition
            0,                                  // flags and attributes
            NULL);                              // template file

        if(hFile == INVALID_HANDLE_VALUE)
        {
            goto Cleanup;
        }

        // Do we need to initialize the file?
        if(GetLastError() != ERROR_ALREADY_EXISTS)
        {
            fInitialize = TRUE;
        }

        if(GetFileSize(hFile, NULL) != sizeof(CACHE_FILE))
        {
            fInitialize = TRUE;
            SetFilePointer(hFile, sizeof(CACHE_FILE), NULL, FILE_BEGIN);
            SetEndOfFile(hFile);
        }

        // Create the mapping.
        s_hCacheFileMapping = CreateFileMappingA(
            hFile,          // file
            NULL,           // security
            PAGE_READWRITE, // protect
            0,              // size low
            0,              // size high
            OBJ_NAME(Map)); // name

        if(!s_hCacheFileMapping)
        {
            goto Cleanup;
        }

        // Map view of file.
        pCacheFile = (CACHE_FILE*)MapViewOfFile(
            s_hCacheFileMapping, 
            FILE_MAP_ALL_ACCESS,    // access
            0,                      // offset low
            0,                      // offset high
            sizeof(CACHE_FILE));    // number of bytes to map

        if(!pCacheFile)
        {
            goto Cleanup;
        }

        // Reset the contents of the file if we are opening it for the
        // first time or if the file contents do not look right.
        __try
        {
            if(fInitialize || pCacheFile->dwMagic!=MAGIC ||
                pCacheFile->dwVersion!=VERSION)
            {
                memset(pCacheFile, 0, sizeof(CACHE_FILE));
                pCacheFile->dwMagic = MAGIC;
                pCacheFile->dwVersion = VERSION;
                FlushViewOfFile(pCacheFile, sizeof(CACHE_FILE));
            }
        }
        __except(EXCEPTION_EXECUTE_HANDLER)
        {
        }

    }

    s_pCacheFile = pCacheFile;

Cleanup:
    if(!s_pCacheFile && s_hCacheFileMapping)
    {
        CloseHandle(s_hCacheFileMapping);
        s_hCacheFileMapping = NULL;
    }

    if(hMutex)
    {
        ReleaseMutex(hMutex);
        CloseHandle(hMutex);
    }  

    if(hFile != INVALID_HANDLE_VALUE)
    {
        CloseHandle(hFile);
    }

    return s_pCacheFile != NULL;
}

//+------------------------------------------------------------------------
//
//  Function:   DeinitImageSizeCache
//
//  Synopsis:   Undo action of InitImageSizeCache 
//
//-------------------------------------------------------------------------
void DeinitImageSizeCache()
{
    if(s_pCacheFile)
    {
        FlushViewOfFile(s_pCacheFile, sizeof(CACHE_FILE));
    }
    if(s_hCacheFileMapping)
    {
        CloseHandle(s_hCacheFileMapping);
    }

    s_pCacheFile = NULL;
    s_hCacheFileMapping = NULL;
}

//+------------------------------------------------------------------------
//
//  Function:   _HashData
//
//  Synopsis:   Compute hash of bytes.
//
//-------------------------------------------------------------------------
// BUGBUG (garybu) Hash function is duplicated. Should export common fn from shlwapi.
const static BYTE Translate[256] =
{
    1, 14,110, 25, 97,174,132,119,138,170,125,118, 27,233,140, 51,
    87,197,177,107,234,169, 56, 68, 30,  7,173, 73,188, 40, 36, 65,
    49,213,104,190, 57,211,148,223, 48,115, 15,  2, 67,186,210, 28,
    12,181,103, 70, 22, 58, 75, 78,183,167,238,157,124,147,172,144,
    176,161,141, 86, 60, 66,128, 83,156,241, 79, 46,168,198, 41,254,
    178, 85,253,237,250,154,133, 88, 35,206, 95,116,252,192, 54,221,
    102,218,255,240, 82,106,158,201, 61,  3, 89,  9, 42,155,159, 93,
    166, 80, 50, 34,175,195,100, 99, 26,150, 16,145,  4, 33,  8,189,
    121, 64, 77, 72,208,245,130,122,143, 55,105,134, 29,164,185,194,
    193,239,101,242,  5,171,126, 11, 74, 59,137,228,108,191,232,139,
    6, 24, 81, 20,127, 17, 91, 92,251,151,225,207, 21, 98,113,112,
    84,226, 18,214,199,187, 13, 32, 94,220,224,212,247,204,196, 43,
    249,236, 45,244,111,182,153,136,129, 90,217,202, 19,165,231, 71,
    230,142, 96,227, 62,179,246,114,162, 53,160,215,205,180, 47,109,
    44, 38, 31,149,135,  0,216, 52, 63, 23, 37, 69, 39,117,146,184,
    163,200,222,235,248,243,219, 10,152,131,123,229,203, 76,120,209
};

void _HashData(LPBYTE pbData, DWORD cbData, LPBYTE pbHash, DWORD cbHash)
{
    DWORD i, j;
    //  seed the hash
    for(i = cbHash; i-- > 0;)
    {
        pbHash[i] = (BYTE) i;
    }

    //  do the hash
    for (j = cbData; j-- > 0;)
    {
        for (i = cbHash; i-- > 0;)
            pbHash[i] = Translate[pbHash[i] ^ pbData[j]];
    }
}

//+------------------------------------------------------------------------
//
//  Function:   GetCachedImageSize
//
//  Synopsis:   Get cached image size, returns FALSE if not found. 
//
//-------------------------------------------------------------------------
BOOL GetCachedImageSize(LPCTSTR pchURL, SIZE* psize)
{
    CACHE_ENTRY* pCacheEntry;

    struct { DWORD dw; WORD  w; } Hash;

    if(!InitImageSizeCache())
    {
        return FALSE;
    }

    _HashData((BYTE*)pchURL, _tcslen(pchURL)*sizeof(TCHAR), (BYTE*)&Hash, sizeof(Hash));
    if(Hash.dw == 0) Hash.dw = 1;

    pCacheEntry = &s_pCacheFile->aCacheEntry[Hash.w % ARRAYSIZE(s_pCacheFile->aCacheEntry)];

    // The following chunk of code can read bad values because the cache memory 
    // is shared by multiple processes.  We don't bother with the expense
    // of a mutex because we can tolerate fetching a bad image size.

    // read / write to/from Maped file can raise exceptions
    __try
    {
        if(pCacheEntry->dwHash == Hash.dw)
        {
            psize->cx = pCacheEntry->wWidth;
            psize->cy = pCacheEntry->wHeight;
            return TRUE;
        }
    }
    __except(EXCEPTION_EXECUTE_HANDLER)
    {
    }

    return FALSE;
}


//+------------------------------------------------------------------------
//
//  Function:   SetCachedImageSize
//
//  Synopsis:   Save image size for future use. 
//
//-------------------------------------------------------------------------
void SetCachedImageSize(LPCTSTR pchURL, SIZE size)
{
    CACHE_ENTRY* pCacheEntry;

    struct { DWORD dw; WORD  w; } Hash;

    if(!InitImageSizeCache())
    {
        return;
    }

    _HashData((BYTE*)pchURL, _tcslen(pchURL)*sizeof(TCHAR), (BYTE*)&Hash, sizeof(Hash));
    if(Hash.dw == 0) Hash.dw = 1;
    pCacheEntry = &s_pCacheFile->aCacheEntry[Hash.w%ARRAYSIZE(s_pCacheFile->aCacheEntry)];

    // The following chunk of code can confuse GetCachedImageSize because
    // the cache memory is shared by multiple processes.  We don't 
    // bother with the expense of a mutex because we can tolerate fetching a 
    // bad image size.

    // read / write to/from Maped file can raise exceptions
    __try
    {
        pCacheEntry->wWidth = size.cx;
        pCacheEntry->wHeight = size.cy;
        pCacheEntry->dwHash = Hash.dw;
    }
    __except(EXCEPTION_EXECUTE_HANDLER)
    {
    }
}

extern NEWIMGTASKFN NewImgTaskArt;

extern HRESULT CreateImgDataObject(
           CDocument*   pDoc,
           CImgCtx*     pImgCtx,
           CBitsCtx*    pBitsCtx,
           CElement*    pElement,
           CGenDataObject** ppImgDO);

DWORD g_dwImgIdInc = 0x80000000;

static ATOM s_atomActiveMovieWndClass = NULL;
LRESULT CALLBACK ActiveMovieWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

static void CopyDibToClipboard(CImgCtx* pImgCtx, CGenDataObject* pDO)
{
    HRESULT hr;
    HGLOBAL hgDib = NULL;
    IStream* pStm = NULL;

    Assert(pImgCtx && pDO);

    hr = CreateStreamOnHGlobal(NULL, FALSE, &pStm);
    if(hr)
    {
        goto Cleanup;
    }
    hr = pImgCtx->SaveAsBmp(pStm, FALSE);

    GetHGlobalFromStream(pStm, &hgDib);

    if(!hr && hgDib)
    {
        pDO->AppendFormatData(CF_DIB, hgDib);
        hgDib = NULL;
    }

Cleanup:
    if(hgDib)
    {
        GlobalFree(hgDib);
    }
    ReleaseInterface(pStm);
}

static void CopyHtmlToClipboard(CElement* pElement, CGenDataObject* pDO)
{
    HRESULT     hr;
    CDocument*  pDoc;
    HGLOBAL     hgHTML = NULL;
    IStream*    pStm = NULL;

    Assert(pElement && pDO);
    pDoc = pElement->Doc();

    CMarkupPointer mpStart(pDoc), mpEnd(pDoc);

    hr = CreateStreamOnHGlobal(NULL, FALSE, &pStm);
    if(hr)
    {
        goto Cleanup;
    }

    hr = mpStart.MoveAdjacentToElement(pElement, ELEM_ADJ_BeforeBegin);
    if(hr)
    {
        goto Cleanup;
    }
    hr = mpEnd.MoveAdjacentToElement(pElement, ELEM_ADJ_AfterEnd);
    if(hr)
    {
        goto Cleanup;
    }

    {
        CStreamWriteBuff StreamWriteBuff(pStm, CP_UTF_8);
        CRangeSaver rs(
            &mpStart,
            &mpEnd,
            RSF_CFHTML,
            &StreamWriteBuff,
            mpStart.Markup());

        StreamWriteBuff.SetFlags(WBF_NO_NAMED_ENTITIES);
        hr = rs.Save();
        if(hr)
        {
            goto Cleanup;
        }
        StreamWriteBuff.Terminate(); // appends a null character
    }

    GetHGlobalFromStream(pStm, &hgHTML);

    if(hgHTML)
    {
        pDO->AppendFormatData(cf_HTML, hgHTML);
    }

Cleanup:
    ReleaseInterface(pStm);
}

//+---------------------------------------------------------------------------
//
//  Function:   CopyFileToClipboard
//
//  Synopsis:   Supports CF_HDROP format
//
//  Arguments:  [pchPath]   Fully expanded file path (input)
//              [pDO]       Pointer to the data object (input)
//
//  Notes:      Mostly copied from IE3.
//
//----------------------------------------------------------------------------
HRESULT CopyFileToClipboard(const TCHAR* pchPath, CGenDataObject* pDO)
{
    HRESULT     hr = S_OK;
    ULONG       cchPath, cchAnsi;
    HGLOBAL     hgDropFiles = NULL;
    DROPFILES*  pdf;
    char*       pchDropPath;

    if(!pchPath || !*pchPath || !pDO)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    cchPath = _tcslen(pchPath);

    // Convert file path to ANSI. While NT shell can handle both
    // UNICODE and ANSI, Win95 shell can handle only ANSI.
    cchAnsi = WideCharToMultiByte(CP_ACP, 0, pchPath, cchPath+1, NULL, 0, NULL, NULL);
    if(cchAnsi <= 0)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    // (+ 2) for double null terminator.
    // GPTR does zero-init for us
    hgDropFiles = GlobalAlloc(GPTR, sizeof(DROPFILES)+cchAnsi+1);
    if(!hgDropFiles)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    pdf = (DROPFILES*)hgDropFiles;
    if(!pdf)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }
    pdf->pFiles = sizeof(DROPFILES);

    pchDropPath = (char*)(pdf + 1);
    WideCharToMultiByte(CP_ACP, 0, pchPath, cchPath+1, pchDropPath, cchAnsi+1, NULL, NULL);

    hr = pDO->AppendFormatData(CF_HDROP, hgDropFiles);
    if(hr)
    {
        goto Cleanup;
    }

    // Don't free hgDropFiles on subsequent error now that
    // hgDropFiles has been added to the data object.
    // IDataObject::Release() will.
    hgDropFiles = NULL;

Cleanup:
    if(hgDropFiles)
    {
        GlobalFree(hgDropFiles);
    }
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Function:   CreateImgDataObject
//
//  Synopsis:   Create an image data object (supports CF_HDROP, CF_HTML & CF_DIB formats)
//
//----------------------------------------------------------------------------
HRESULT CreateImgDataObject(
        CDocument*  pDoc,
        CImgCtx*    pImgCtx,
        CBitsCtx*   pBitsCtx,
        CElement*   pElement,
        CGenDataObject** ppImgDO)
{
    HRESULT         hr;
    TCHAR*          pchPath = NULL;
    CGenDataObject* pImgDO = NULL;
    IDataObject*    pDO = NULL;

    if(!pImgCtx && !pBitsCtx)
    {
        hr = E_INVALIDARG;
        goto Cleanup;
    }

    pImgDO = new CGenDataObject(pDoc);
    if(!pImgDO)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    // Add CF_HDROP format
    if(pBitsCtx)
    {
        hr = pBitsCtx->GetFile(&pchPath);
    }
    else
    {
        hr = pImgCtx->GetFile(&pchPath);
    }
    if(!hr)
    {
        CopyFileToClipboard(pchPath, pImgDO);
    }

    // Add CF_DIB format
    if(pImgCtx && !pBitsCtx)
    {
        CopyDibToClipboard(pImgCtx, pImgDO);
    }

    // Add CF_HTML format
    if(pElement)
    {
        CopyHtmlToClipboard(pElement, pImgDO);
    }

    if(ppImgDO)
    {
        *ppImgDO = pImgDO;
        pImgDO = NULL;
        hr = S_OK;
    }
    else
    {
        // Tell the shell that we want a copy, not a move
        pImgDO->SetPreferredEffect(DROPEFFECT_COPY);

        hr = pImgDO->QueryInterface(IID_IDataObject, (void**)&pDO);
        if(hr)
        {
            goto Cleanup;
        }

        hr = pDoc->SetClipboard(pDO);
    }

Cleanup:
    // cast to disambiguate between IUnknown's of IDataObject and IDropSource
    ReleaseInterface((IDataObject*)pImgDO);
    ReleaseInterface(pDO);
    MemFreeString(pchPath);
    RRETURN(hr);
}

#define IMP_IMG_FIRE_BOOL(eventName)                                    \
    BOOL CImgHelper::Fire_##eventName()                                 \
    { \
        Assert(_pOwner);                                                \
        return DYNCAST(CImgElement, _pOwner)->Fire_##eventName();       \
    }

#define IMP_IMG_FIRE_VOID(eventName)                                    \
    void CImgHelper::Fire_##eventName()                                 \
    {                                                                   \
        Assert(_pOwner);                                                \
        DYNCAST(CImgElement, _pOwner)->Fire_##eventName();              \
    }

#define IMP_IMG_GETAA(returnType, propName)                             \
    returnType CImgHelper::GetAA##propName () const                     \
    {                                                                   \
        Assert(_pOwner);                                                \
        return DYNCAST(CImgElement, _pOwner)->GetAA##propName();        \
    }

#define IMP_IMG_SETAA(paraType, propName)                               \
    HRESULT CImgHelper::SetAA##propName (paraType pv)                   \
    {                                                                   \
        Assert(_pOwner);                                                \
        return DYNCAST(CImgElement, _pOwner)->SetAA##propName(pv);      \
    }

IMP_IMG_GETAA(LPCTSTR, alt);

IMP_IMG_GETAA(LPCTSTR, src);
IMP_IMG_SETAA(LPCTSTR, src);

IMP_IMG_GETAA(CUnitValue, border);
IMP_IMG_GETAA(long, vspace);
IMP_IMG_GETAA(long, hspace);
IMP_IMG_GETAA(LPCTSTR, lowsrc);
IMP_IMG_GETAA(LPCTSTR, dynsrc);
IMP_IMG_GETAA(htmlStart, start);

IMP_IMG_GETAA(VARIANT_BOOL, complete);
IMP_IMG_SETAA(VARIANT_BOOL, complete);

IMP_IMG_GETAA(long, loop);
IMP_IMG_GETAA(LPCTSTR, onload);
IMP_IMG_GETAA(LPCTSTR, onerror);
IMP_IMG_GETAA(LPCTSTR, onabort);
IMP_IMG_GETAA(LPCTSTR, name);
IMP_IMG_GETAA(LPCTSTR, title);
IMP_IMG_GETAA(VARIANT_BOOL, cache);

IMP_IMG_FIRE_VOID(onerror);
IMP_IMG_FIRE_VOID(onload);
IMP_IMG_FIRE_VOID(onabort);
IMP_IMG_FIRE_VOID(onreadystatechange);
IMP_IMG_FIRE_BOOL(onbeforecopy);
IMP_IMG_FIRE_BOOL(onbeforepaste);
IMP_IMG_FIRE_BOOL(onbeforecut);
IMP_IMG_FIRE_BOOL(oncut);
IMP_IMG_FIRE_BOOL(oncopy);
IMP_IMG_FIRE_BOOL(onpaste);

BOOL CImgHelper::IsHSpaceDefined() const
{
    Assert(_pOwner);
    return DYNCAST(CImgElement, _pOwner)->IsHSpaceDefined();
}

CImageLayout* CImgHelper::Layout()
{
    Assert(_pOwner);
    return DYNCAST(CImageLayout ,DYNCAST(CImgElement, _pOwner)->Layout());
}

void CImgHelper::SetImgAnim(BOOL fDisplayNoneOrHidden)
{
    if(!!_fDisplayNoneOrHidden != !!fDisplayNoneOrHidden)
    {
        _fDisplayNoneOrHidden = fDisplayNoneOrHidden;

        if(_fAnimated)
        {
            if(_fDisplayNoneOrHidden)
            {
                if(_lCookie)
                {
                    CImgAnim* pImgAnim = GetImgAnim();

                    if(pImgAnim)
                    {
                        pImgAnim->UnregisterForAnim(this, _lCookie);
                        _lCookie = 0;
                    }
                }
            }
            else
            {
                CImgAnim* pImgAnim = CreateImgAnim();

                if(pImgAnim)
                {
                    if(!_lCookie)
                    {
                        pImgAnim->RegisterForAnim(
                            this,
                            (DWORD_PTR)Doc(),
                            g_dwImgIdInc++,
                            OnAnimSyncCallback,
                            NULL,
                            &_lCookie);
                    }

                    if(_lCookie)
                    {
                        pImgAnim->ProgAnim(_lCookie);
                    }
                }
            }
        }
    }
}

CImgHelper::CImgHelper(CDocument* pDoc, CSite* pSite)
{
    _pOwner = pSite;
    _fIsInPlace = TRUE;
    _fExpandAltText = (pDoc && pDoc->_pOptionSettings) ?
        (g_fScreenReader || pDoc->_pOptionSettings->fExpandAltText)
        &&!pDoc->_pOptionSettings->fShowImages : FALSE;
}

//+------------------------------------------------------------------------
//
//  Member:     EnterTree
//
//-------------------------------------------------------------------------
HRESULT CImgHelper::EnterTree()
{
    HRESULT     hr = S_OK;
    TCHAR       cBuf[pdlUrlLen];
    TCHAR*      pchNewUrl = cBuf;
    CDocument*  pDoc = Doc();

    _fIsInPlace = TRUE;

    if(_readyStateImg == READYSTATE_COMPLETE)
    {
        goto Cleanup;
    }

    _fCache = GetAAcache();

    // Ignore HR for the following two cases:
    // if load fails for security or other reason, it's not a problem for Init2
    // of the img element.
    SetImgDynsrc();
    SetImgSrc(0); // Don't request resize; don't invalidate frame

Cleanup:
    RRETURN(hr);
}

void CImgHelper::ExitTree(CNotification* pnf)
{
    CImgCtx* pImgCtx = GetImgCtx();

    if((_pVideoObj || _hwnd) && !pnf->IsSecondChance())
    {
        pnf->SetSecondChanceRequested();
        return;
    }

    if(_fIsInPlace)
    {
        _fIsInPlace = FALSE;
        SetActivity();
    }

    SetVideo();
    if(_pVideoObj)
    {
        _pVideoObj->Release();
        _pVideoObj = NULL;
    }
    if(_hwnd)
    {
        DestroyWindow(_hwnd);
        _hwnd = NULL;
    }

    if(pImgCtx)
    {
        pImgCtx->Disconnect();
    }

    if(_pImgCtx && _lCookie)
    {
        CImgAnim* pImgAnim = GetImgAnim();

        if(pImgAnim)
        {
            pImgAnim->UnregisterForAnim(this, _lCookie);
            _lCookie = 0;
        }
    }

    if(_hbmCache)
    {
        DeleteObject(_hbmCache);
        _hbmCache = NULL;
    }
}

//+------------------------------------------------------------------------
//
//  Member:     FetchAndSetImgCtx
//
//-------------------------------------------------------------------------
HRESULT CImgHelper::FetchAndSetImgCtx(const TCHAR* pchUrl, DWORD dwSetFlags)
{
    CImgCtx *pImgCtx;
    HRESULT hr;

    hr = Doc()->NewDwnCtx(DWNCTX_IMG, pchUrl, _pOwner, (CDwnCtx**)&pImgCtx);

    if(hr == S_OK)
    {
        SetImgCtx(pImgCtx, dwSetFlags);

        if(pImgCtx)
        {
            pImgCtx->Release();
        }
    }

    RRETURN(hr);
}

HRESULT CImgHelper::SetImgSrc(DWORD dwSetFlags)
{
    HRESULT hr;
    LPCTSTR szUrl = GetAAsrc();

    if(szUrl)
    {
        hr = FetchAndSetImgCtx(szUrl, dwSetFlags);
    }
    else
    {
        hr = FetchAndSetImgCtx(GetAAlowsrc(), dwSetFlags);
    }

    RRETURN(hr);
}

HRESULT CImgHelper::SetImgDynsrc()
{
    HRESULT hr = S_OK;
    LPCTSTR szDynSrcUrl;

    if(!(Doc()->_dwLoadf  & DLCTL_VIDEOS))
    {
        goto Cleanup;
    }

    szDynSrcUrl = GetAAdynsrc();
    if(szDynSrcUrl)
    {
        CBitsCtx* pBitsCtx;

        hr = Doc()->NewDwnCtx(DWNCTX_FILE, szDynSrcUrl, _pOwner, (CDwnCtx**)&pBitsCtx);

        if(hr == S_OK)
        {
            SetBitsCtx(pBitsCtx);

            if(pBitsCtx)
            {
                pBitsCtx->Release();
            }
        }
    }

Cleanup:
    RRETURN(hr);
}

void CImgHelper::InvalidateFrame()
{
    CImgElementLayout* pLayout = DYNCAST(CImgElementLayout, Layout());
    pLayout->InvalidateFrame();
}

void CImgHelper::SetImgCtx(CImgCtx* pImgCtx, DWORD dwSetFlags)
{
    if(_pImgCtx && _lCookie)
    {
        CImgAnim* pImgAnim = GetImgAnim();

        if(pImgAnim)
        {
            pImgAnim->UnregisterForAnim(this, _lCookie);
            _lCookie = 0;
        }
    }
    _fAnimated = FALSE;

    if(pImgCtx)
    {
        ULONG ulChgOn = IMGCHG_COMPLETE;
        CDocument* pDoc = Doc();

        if(_fIsActive)
        {
            ulChgOn |= IMGCHG_VIEW;
        }
        if(pDoc->_pOptionSettings->fPlayAnimations)
        {
            ulChgOn |= IMGCHG_ANIMATE;
        }

        SIZE sizeImg;
        BOOL fSelectSizeChange = TRUE;
        if((_fCreatedWithNew || !IsInMarkup()) && !_fSizeInCtor && GetCachedImageSize(pImgCtx->GetUrl(), &sizeImg))
        {
            putWidth(sizeImg.cx);
            putHeight(sizeImg.cy);
            fSelectSizeChange = FALSE;
        }

        ulChgOn |= (((_fCreatedWithNew || !IsInMarkup()) && !_fSizeInCtor && fSelectSizeChange) ? IMGCHG_SIZE : 0);

        // also need size if we've gone through CalcSize without an ImgCtx already
        ulChgOn |= (_fNeedSize ? IMGCHG_SIZE : 0);

        SetReadyStateImg(READYSTATE_UNINITIALIZED);

        pImgCtx->SetCallback(OnDwnChanCallback, this);
        pImgCtx->SelectChanges(ulChgOn, 0, TRUE);
        if(!(pImgCtx->GetState(FALSE, NULL) & (IMGLOAD_COMPLETE|IMGLOAD_STOPPED|IMGLOAD_ERROR)))
        {
            pImgCtx->SetProgSink(pDoc->GetProgSink());
        }

        pImgCtx->AddRef(); // addref the new before releasing the old
    }

    if(_pImgCtx)
    {
        _pImgCtx->SetProgSink(NULL);
        _pImgCtx->Disconnect();
        _pImgCtx->SetLoad(FALSE, NULL, FALSE);
        if(!pImgCtx)
        {
            _pImgCtx->Release();
            _pImgCtx = NULL;

            if((dwSetFlags&IMGF_REQUEST_RESIZE) && HasLayout())
            {
                ResizeElement(NFLAGS_FORCE);
            }
        }

        if(_pImgCtxPending)
        {
            _pImgCtxPending->SetLoad(FALSE, NULL, FALSE);
            _pImgCtxPending->Disconnect();
            _pImgCtxPending->Release();
        }

        _pImgCtxPending = pImgCtx;
    }
    else
    {
        _pImgCtx = pImgCtx;

        if((dwSetFlags&(IMGF_REQUEST_RESIZE|IMGF_INVALIDATE_FRAME)) && HasLayout())
        {
            if(dwSetFlags & IMGF_REQUEST_RESIZE)
            {
                ResizeElement(NFLAGS_FORCE);
            }
            else
            {
                InvalidateFrame();
            }
        }
    }

    if(HasLayout())
    {
        // If there is no pending call to resize the element,
        // post a layout request so that CLayout::EnsureDispNode gets a chance to
        // execute and adjust the opacity of the display node
        // BUGBUG: This should eventually be replaced by code that tracks when the
        //         display node itself is possible dirty and can schedule a fix-up
        //         (brendand)
        if(!Layout()->_fSizeThis)
        {
            Layout()->PostLayoutRequest(LAYOUT_MEASURE);
        }
    }

    _fSizeChange = FALSE;
    _fNeedSize = FALSE;

    if(pImgCtx)
    {
        OnDwnChan(pImgCtx);
    }
}

void CImgHelper::SetBitsCtx(CBitsCtx* pBitsCtx)
{
    if(_pBitsCtx)
    {
        _pBitsCtx->SetProgSink(NULL);
        _pBitsCtx->Disconnect();
        _pBitsCtx->Release();
    }

    _pBitsCtx = pBitsCtx;

    if(pBitsCtx)
    {
        pBitsCtx->AddRef(); // addref then new before releasing the old

        _fStopped = FALSE;

        if(pBitsCtx->GetState() & (DWNLOAD_COMPLETE|DWNLOAD_ERROR|DWNLOAD_STOPPED))
        {
            OnDwnChan(pBitsCtx);
        }
        else
        {
            pBitsCtx->SetProgSink(Doc()->GetProgSink());
            pBitsCtx->SetCallback(OnDwnChanCallback, this);
            pBitsCtx->SelectChanges(DWNCHG_COMPLETE, 0, TRUE);
        }
    }
}

HRESULT CImgHelper::putHeight(long l)
{
    Assert(_pOwner);
    RRETURN(DYNCAST(CImgElement, _pOwner)->putHeight(l));
}

HRESULT CImgHelper::GetHeight(long* pl)
{
    Assert(_pOwner);
    RRETURN(DYNCAST(CImgElement, _pOwner)->GetHeight(pl));
}

STDMETHODIMP CImgHelper::get_height(long* p)
{
    HRESULT hr = S_OK;

    if(p)
    {
        if(!IsInMarkup())
        {
            hr = GetHeight(p);
        }
        else if(_fCreatedWithNew)
        {
            *p = GetFirstBranch()->GetCascadedheight().GetUnitValue();
        }
        else
        {
            _pOwner->SendNotification(NTYPE_ELEMENT_ENSURERECALC);

            *p = Layout()->GetContentHeight();
        }
    }
    else
    {
        hr = E_INVALIDARG;
    }

    return hr;
}

HRESULT CImgHelper::putWidth(long l)
{
    Assert(_pOwner);
    RRETURN(DYNCAST(CImgElement, _pOwner)->putWidth(l));
}

HRESULT CImgHelper::GetWidth(long* pl)
{
    Assert(_pOwner);
    RRETURN(DYNCAST(CImgElement, _pOwner)->GetWidth(pl));
}

STDMETHODIMP CImgHelper::get_width(long* p)
{
    HRESULT hr = S_OK;

    if(p)
    {
        if(!IsInMarkup())
        {
            hr = GetWidth(p);
        }
        else if(_fCreatedWithNew)
        {
            *p = GetFirstBranch()->GetCascadedwidth().GetUnitValue();
        }
        else
        {
            _pOwner->SendNotification(NTYPE_ELEMENT_ENSURERECALC);

            *p = Layout()->GetContentWidth();
        }
    }
    else
    {
        hr = E_INVALIDARG;
    }

    return hr;
}

//+------------------------------------------------------------------------
//
//  Method:     CImgHelper::OnDwnChan
//
//-------------------------------------------------------------------------
void CImgHelper::OnDwnChan(CDwnChan* pDwnChan)
{
    BOOL fPending = FALSE;

    if(!Doc()->_pPrimaryMarkup)
    {
        return;
    }

    Assert(pDwnChan==_pImgCtxPending || pDwnChan==_pImgCtx || pDwnChan==_pBitsCtx);

    if(pDwnChan == _pImgCtxPending)
    {
        if(_pImgCtxPending->GetState() & (IMGCHG_VIEW|IMGCHG_COMPLETE))
        {
            _pImgCtx->Release();
            _pImgCtx = _pImgCtxPending;
            _pImgCtxPending = NULL;
            fPending = TRUE;
        }
    }

    if(pDwnChan==_pImgCtx && !(_pBitsCtx&&_pVideoObj))
    {
        ULONG ulState;
        SIZE  size;

        Trace2("[%08lX] OnDwnChan (enter) '%ls'\n", this, GetAAsrc());

        ulState = _pImgCtx->GetState(TRUE, &size);

        if((ulState&IMGCHG_SIZE) || fPending)
        {
            if(IsInMarkup())
            {
                CTreeNode* pNode;
                CTreeNode::CLock  Lock(GetFirstBranch());

                SetReadyStateImg(READYSTATE_LOADING);
                pNode = GetFirstBranch();
                if(!pNode)
                {
                    return;
                }

                CUnitValue uvWidth = pNode->GetCascadedwidth();
                CUnitValue uvHeight = pNode->GetCascadedheight();

                Trace1("[%08lX] OnChan IMGCHG_SIZE\n", this);

                if(uvWidth.IsNull() || uvHeight.IsNull() || _fExpandAltText)
                {
                    if(_fCreatedWithNew)
                    {
                        putWidth(size.cx);
                        putHeight(size.cy);
                    }
                    else
                    {
                        Trace1("[%08lX] OnChan ResizeElement\n", this);
                        CRect rc;

                        GetRectImg(&rc);
                        if(rc.right-rc.left!=size.cx
                            || rc.bottom-rc.top!=size.cy
                            || _fExpandAltText)
                        {
                            ResizeElement();
                        }
                    }

                    if(!(ulState&IMGLOAD_ERROR) && size.cx && size.cy)
                    {
                        SetCachedImageSize(_pImgCtx->GetUrl(), size);
                    }
                }
            }
            else
            {
                putWidth(size.cx);
                putHeight(size.cy);

                if(!(ulState&IMGLOAD_ERROR) && size.cx && size.cy)
                {
                    SetCachedImageSize(_pImgCtx->GetUrl(), size);
                }
            }
        }

        if(ulState & IMGCHG_VIEW)
        {
            if(_fIsActive)
            {
                long nrc;
                RECT prc[2];
                CRect rectImg;

                Trace1("[%08lX] OnChan IMGCHG_VIEW\n", this);

                GetRectImg(&rectImg);
                _pImgCtx->GetUpdateRects(prc, &rectImg, &nrc);

                Layout()->Invalidate(prc, nrc);
            }
        }
        else if(fPending)
        {
            InvalidateFrame();
        }

        if(ulState & IMGCHG_ANIMATE)
        {
            CImgAnim* pImgAnim = CreateImgAnim();

            _fAnimated = TRUE;

            if(pImgAnim)
            {
                if(!_lCookie && !_fDisplayNoneOrHidden)
                {
                    pImgAnim->RegisterForAnim(
                        this,
                        (DWORD_PTR)Doc(),
                        _pImgCtx->GetImgId(),
                        OnAnimSyncCallback,
                        NULL, 
                        &_lCookie);
                }

                if(_lCookie)
                {
                    pImgAnim->ProgAnim(_lCookie);
                }
            }
        }

        if(ulState & IMGCHG_COMPLETE)
        {
            Trace1("[%08lX] OnChan IMGCHG_COMPLETE\n", this);
            Assert(_pOwner);
            CElement::CLock  Lock(_pOwner);

            if(ulState & (IMGLOAD_ERROR|IMGLOAD_STOPPED))
            {
                if((_fCreatedWithNew || !IsInMarkup()) && !_fSizeInCtor)
                {
                    GetPlaceHolderBitmapSize(TRUE, &size);
                    putWidth(size.cx+GRABSIZE*2);
                    putHeight(size.cy+GRABSIZE*2);
                }

                if(ulState & IMGLOAD_STOPPED)
                {
                    Fire_onabort();
                }
                else
                {
                    Fire_onerror();

                    if(g_bJGJitState == JIT_NEED_JIT)
                    {
                        g_bJGJitState = JIT_PENDING;
                        GWPostMethodCall(Doc(), ONCALL_METHOD(CDocument, FaultInJG, faultinjg), 0, FALSE, "CDocument::FaultInJG");
                    }
                }

            }

            if(ulState & IMGLOAD_COMPLETE)
            {
                SetReadyStateImg(READYSTATE_COMPLETE);
                SetAAcomplete(TRUE);
            }

            _pImgCtx->SetProgSink(NULL);
            if(!GetFirstBranch())
            {
                Trace1("[%08lX] OnChan (leave)\n", this);
                return;
            }
        }

        Trace1("[%08lX] OnChan (leave)\n", this);
    }
    else if(pDwnChan == _pBitsCtx)
    {
        ULONG ulState = _pBitsCtx->GetState();

        if(ulState & DWNLOAD_COMPLETE)
        {
            CDocument* pDoc = Doc();

            // Create the video object if it doesn't exist
            if(!_pVideoObj)
            {
                _pVideoObj = (CIEMediaPlayer*)new CIEMediaPlayer;
                pDoc->_fBroadcastInteraction = TRUE;
                pDoc->_fBroadcastStop = TRUE;
            }

            if(_pVideoObj)
            {
                TCHAR* pchFile = NULL;

                if(OK(_pBitsCtx->GetFile(&pchFile)) &&
                    OK(_pVideoObj->SetURL(pchFile))) // Initialize & RenderFile
                {
                    _pVideoObj->SetLoopCount(GetAAloop());
                }
                else
                {
                    _pVideoObj->Release();
                    _pVideoObj = NULL;
                }
                MemFreeString(pchFile);
            }

            ResizeElement();

            SetVideo();

        }
        else if(ulState & DWNLOAD_ERROR)
        {
            if(_pVideoObj)
            {
                _pVideoObj->Release();
                _pVideoObj = NULL;
            }
            if(_hwnd)
            {
                DestroyWindow(_hwnd);
                _hwnd = NULL;
            }
        }

        // BUGBUG:
        // else we should fire an event to signal the error
        if(ulState & (DWNLOAD_COMPLETE|DWNLOAD_STOPPED|DWNLOAD_ERROR))
        {
            _pBitsCtx->SetProgSink(NULL);
        }
    }
}

//--------------------------------------------------------------------------
//
//  Method:     CImgHelper::ImgAnimCallback
//
//  Synopsis:   Called by the CImgAnim when certain events take place.
//
//--------------------------------------------------------------------------
void CImgHelper::OnAnimSync(DWORD dwReason, void* pvParam, void** ppvDataOut, IMGANIMSTATE* pImgAnimState)
{
    switch(dwReason)
    {
    case ANIMSYNC_GETIMGCTX:
        *(CImgCtx**) ppvDataOut = _pImgCtx;
        break;

    case ANIMSYNC_GETHWND:
        {
            CDocument* pDoc = Doc();

            *(HWND*)ppvDataOut = pDoc->_pInPlace ? pDoc->_pInPlace->_hwnd : NULL;
        }
        break;

    case ANIMSYNC_TIMER:
        if(_fIsActive)
        {
            InvalidateFrame();
            *(BOOL*) ppvDataOut = TRUE;
        }
        else
        {
            *(BOOL*) ppvDataOut = FALSE;
        }

        if(pImgAnimState->fLoop)
        {
            Fire_onload();
        }
        break;

    case ANIMSYNC_INVALIDATE:
        if(_fIsActive)
        {
            InvalidateFrame();
            *(BOOL*) ppvDataOut = TRUE;
        }
        else
        {
            *(BOOL*) ppvDataOut = FALSE;
        }
        break;

    default:
        Assert(FALSE);
    }
}

//+-------------------------------------------------------------------------
//
//  Method      CImgHelper::Cleanup
//
//  Synopsis    Shutdown main object by releasing references to
//              other objects and generally cleaning up.
//
//              Release any event connections held by the form.
//
//--------------------------------------------------------------------------
void CImgHelper::CleanupImage ( )
{
    if(_pVideoObj)
    {
        _pVideoObj->Release();
        _pVideoObj = NULL;
    }

    SetImgCtx(NULL, 0);
    SetBitsCtx(NULL);

    if(_hwnd)
    {
        DestroyWindow(_hwnd);
        _hwnd = NULL;
    }

    if(_hbmCache)
    {
        DeleteObject(_hbmCache);
        _hbmCache = NULL;
    }
}

//--------------------------------------------------------------------------
//
//  Method:     CImgHelper::Passivate
//
//  Synopsis:   This function is called when the main reference count
//              goes to zero.  The destructor is called when
//              the reference count for the main object and all
//              embedded sub-objects goes to zero.
//
//--------------------------------------------------------------------------
void CImgHelper::Passivate ( )
{
    Assert(!IsInMarkup());

    CleanupImage();
}

//+-------------------------------------------------------------------------
//
//  Function:   ActiveMovieWndProc
//
//--------------------------------------------------------------------------
LRESULT CALLBACK ActiveMovieWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    CImgHelper* pImgHelper;

    switch(msg)
    {
    case WM_NCCREATE:
        pImgHelper = (CImgHelper*)((LPCREATESTRUCTW) lParam)->lpCreateParams;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR) pImgHelper);
        return TRUE;
        break;

    case WM_NCDESTROY:
        pImgHelper = (CImgHelper*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
        if(pImgHelper)
        {
            pImgHelper->_hwnd = NULL;
        }
        SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
        return TRUE;
        break;

    case WM_MOUSEMOVE:
        pImgHelper = (CImgHelper*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
        if(pImgHelper && pImgHelper->GetAAstart()==htmlStartmouseover)
        {
            pImgHelper->Replay();
        }
        break;

    case WM_LBUTTONDOWN:
    case WM_RBUTTONDOWN:
    case WM_MBUTTONDOWN:
    case WM_LBUTTONUP:
    case WM_RBUTTONUP:
    case WM_MBUTTONUP:
        POINT ptMap;
        HWND hwndParent = GetParent(hwnd);
        ptMap.x = LOWORD(lParam);
        ptMap.y = HIWORD(lParam);
        MapWindowPoints(hwnd, hwndParent, &ptMap, 1);
        lParam = MAKELPARAM(ptMap.x, ptMap.y);
        return SendMessage(hwndParent, msg, wParam, lParam);
    }

    return DefWindowProc(hwnd, msg, wParam, lParam);
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgHelper::SetVideo
//
//  Synopsis:   Checks if we should animate the image
//
//----------------------------------------------------------------------------
void CImgHelper::SetVideo()
{
    HRESULT hr;
    BOOL fEnableInteraction;
    TCHAR* pszWndClass;

    if(!_pVideoObj)
    {
        return;
    }

    fEnableInteraction = Doc()->_fEnableInteraction;

    if(_fVideoPositioned && _fIsInPlace && fEnableInteraction && !_fStopped)
    {
        CRect rcImg;
        CDocument* pDoc = Doc();

        Layout()->GetClientRect(&rcImg, COORDSYS_GLOBAL);

        if(!s_atomActiveMovieWndClass)
        {
            hr = RegisterWindowClass(
                _T("ActiveMovie"),
                ActiveMovieWndProc,
                0,
                NULL, NULL,
                &s_atomActiveMovieWndClass);
            if(hr)
            {
                return;
            }
        }

        pszWndClass = (TCHAR*)(DWORD_PTR)s_atomActiveMovieWndClass;

        if((_hwnd==NULL) && !_pVideoObj->IsAudio())
        {
            _hwnd = CreateWindow(
                pszWndClass,
                NULL,
                WS_CHILD|WS_VISIBLE,
                rcImg.left, rcImg.top,
                rcImg.right - rcImg.left,
                rcImg.bottom - rcImg.top,
                pDoc->GetHWND(),
                NULL,
                _afxGlobalData._hInstThisModule,
                this);

            if(_hwnd == NULL)
            {
                return;
            }
        }

        OffsetRect(&rcImg, -rcImg.left, -rcImg.top);
        _pVideoObj->SetVideoWindow(_hwnd);

        _pVideoObj->SetNotifyWindow(pDoc->GetHWND(), WM_ACTIVEMOVIE, (LONG_PTR)this);
        _pVideoObj->SetWindowPosition(&rcImg);
        _pVideoObj->SetVisible(TRUE);

        if(GetAAstart() == htmlStartfileopen)
        {
            _pVideoObj->Play();
        }
    }
    else if(!_fIsInPlace || !fEnableInteraction)
    {

        // remove us from notifications
        _pVideoObj->SetNotifyWindow(NULL, WM_ACTIVEMOVIE, (LONG_PTR)this);

        // Stop the video
        _pVideoObj->Stop();
    }

    return;
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgHelper::Replay
//
//----------------------------------------------------------------------------
void CImgHelper::Replay()
{
    if(_pVideoObj)
    {
        if(_pVideoObj->GetStatus() == CIEMediaPlayer::IEMM_Completed)
        {
            _pVideoObj->Seek(0);
            _pVideoObj->SetLoopCount(GetAAloop());
        }

        _pVideoObj->Play();
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgHelper::SetActivity
//
//  Synopsis:   Turns activity on or off depending on visibility and
//              in-place activation.
//
//----------------------------------------------------------------------------
void CImgHelper::SetActivity()
{
    if(!!_fIsActive != !!_fIsInPlace)
    {
        CImgCtx* pImgCtx = GetImgCtx();

        _fIsActive = !_fIsActive;

        if(pImgCtx)
        {
            pImgCtx->SelectChanges(
                _fIsActive?IMGCHG_VIEW:0,
                !_fIsActive?IMGCHG_VIEW:0,
                FALSE);
        }
    }
}

BOOL CImgHelper::IsOpaque()
{
    if(!GetFirstBranch() || !HasLayout())
    {
        return FALSE;
    }

    BOOL fOpaque = _pImgCtx ? (_pImgCtx->GetState()&IMGTRANS_OPAQUE) : FALSE;

    if(fOpaque)
    {
        fOpaque = (GetAAvspace()==0) &&
            (IsHSpaceDefined() ? (GetAAHspace()==0) : !IsAligned());
    }

    return fOpaque;
}

//+---------------------------------------------------------------------------
//  Member :     CImgHelper::GetRectImg
//
//  Synopsis   : gets rectImg
//
//+---------------------------------------------------------------------------
void CImgHelper::GetRectImg(CRect* prectImg)
{
    Layout()->GetClientRect(prectImg);
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgHelper::OnReadyStateChange
//
//----------------------------------------------------------------------------
void CImgHelper::OnReadyStateChange()
{
    // do not call super::OnReadyStateChange here - we handle firing the event ourselves
    SetReadyStateImg(_readyStateImg);
}

//+------------------------------------------------------------------------
//
//  Member:     CImgHelper::SetReadyStateImg
//
//  Synopsis:   Use this to set the ready state;
//              it fires OnReadyStateChange if needed.
//
//  Returns:    HRESULT
//
//-------------------------------------------------------------------------
HRESULT CImgHelper::SetReadyStateImg(long readyStateImg)
{
    long readyState;

    _readyStateImg = readyStateImg;

    readyState = min((long)_readyStateImg, _pOwner->GetReadyState());

    if((long)_readyStateFired != readyState)
    {
        _readyStateFired = readyState;

        Fire_onreadystatechange();

        if(_readyStateImg == READYSTATE_COMPLETE)
        {
            Fire_onload();
        }
    }

    return S_OK;
}

//+------------------------------------------------------------------------
//
//  Member:     CImgHelper::Notify
//
//  Synopsis:   Listen for inplace (de)activation so we can turn on/off
//              animation
//
//-------------------------------------------------------------------------
void CImgHelper::Notify(CNotification* pNF)
{
    switch(pNF->Type())
    {
    case NTYPE_DOC_STATE_CHANGE_1:
        if(_pVideoObj && !!_fIsInPlace!=TRUE)
        {
            pNF->SetSecondChanceRequested();
            break;
        }

        // fall through
    case NTYPE_DOC_STATE_CHANGE_2:
        {
            CDocument* pDoc = Doc();

            if(!!_fIsInPlace != TRUE)
            {
                CImgCtx* pImgCtx = GetImgCtx();

                DWNLOADINFO dli;

                _fIsInPlace = !_fIsInPlace;
                SetActivity();

                SetVideo();

                pDoc->InitDownloadInfo(&dli);

                if(pImgCtx && _fIsInPlace && (pDoc->_dwLoadf&DLCTL_DLIMAGES))
                {
                    pImgCtx->SetLoad(TRUE, &dli, FALSE);
                }

                if(_pBitsCtx && _fIsInPlace)
                {
                    _pBitsCtx->SetLoad(TRUE, &dli, FALSE);
                }
            }
        }
        break;

    case NTYPE_STOP_1:
        {
            CImgCtx* pImgCtx = GetImgCtx();

            if(pImgCtx)
            {
                pImgCtx->SetProgSink(NULL);
                pImgCtx->SetLoad(FALSE, NULL, FALSE);

                if(pImgCtx->GetArtPlayer())
                {
                    pNF->SetSecondChanceRequested();
                }
            }

            if(_pBitsCtx)
            {
                _pBitsCtx->SetProgSink(NULL);
                _pBitsCtx->SetLoad(FALSE, NULL, FALSE);
            }

            if(_pVideoObj)
            {
                pNF->SetSecondChanceRequested();
            }

            break;
        }

    case NTYPE_STOP_2:
        {
            CImgCtx* pImgCtx = GetImgCtx();
            if(_pVideoObj)
            {
                _pVideoObj->Stop();
                _fStopped = TRUE;
            }

            if(pImgCtx)
            {
                CArtPlayer* pArtPlayer = pImgCtx->GetArtPlayer();

                if(pArtPlayer)
                {
                    pArtPlayer->DoPlayCommand(IDM_IMGARTSTOP);
                }
            }
        }
        break;

    case NTYPE_ENABLE_INTERACTION_1:
        if(_pVideoObj)
        {
            pNF->SetSecondChanceRequested();
        }
        break;

    case NTYPE_ENABLE_INTERACTION_2:
        SetVideo();
        break;

    case NTYPE_ACTIVE_MOVIE:
        {
            void* pv;
            pNF->Data(&pv);
            if(_pVideoObj && (pv==this))
            {
                _pVideoObj->NotifyEvent(); // Let the video object know something happened
            }
        }
        break;

    case NTYPE_ELEMENT_EXITTREE_1:
    case NTYPE_ELEMENT_EXITTREE_2:
        ExitTree(pNF);
        break;

    case NTYPE_ELEMENT_ENTERTREE:
        EnterTree();
        break;
    }
}

//+-------------------------------------------------------------------------
//
//  Method:     CImgHelper::ShowTooltip
//
//  Synopsis:   Displays the tooltip for the site.
//
//  Arguments:  [pt]    Mouse position in container window coordinates
//              msg     Message passed to tooltip for Processing
//
//--------------------------------------------------------------------------
HRESULT CImgHelper::ShowTooltip(CMessage* pmsg, POINT pt)
{
    HRESULT hr = S_FALSE;
    BOOL fRTL = FALSE;
    CDocument* pDoc = Doc();

    //  Check if we can display alt property as the tooltip.
    TCHAR*  pchString;
    CRect   rc;


    pchString = (LPTSTR)GetAAalt();
    if(pchString == NULL)
    {
        goto Cleanup;
    }

    {
        // Ignore spurious WM_ERASEBACKGROUNDs generated by tooltips
        CServer::CLock Lock(pDoc, SERVERLOCK_IGNOREERASEBKGND);

        Assert(GetFirstBranch());

        // Complex Text - determine if element is right to left for tooltip style setting
        fRTL = GetFirstBranch()->GetCharFormat()->_fRTL;
        Layout()->GetRect(&rc, COORDSYS_GLOBAL);
        FormsShowTooltip(pchString, pDoc->_pInPlace->_hwnd, *pmsg, &rc, (DWORD_PTR)this, fRTL);
        hr = S_OK;
    }

Cleanup:
    return hr;
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgHelper::QueryStatus, public
//
//  Synopsis:   Implements QueryStatus for CImgHelper
//
//----------------------------------------------------------------------------
HRESULT CImgHelper::QueryStatus(
        GUID*       pguidCmdGroup,
        ULONG       cCmds,
        MSOCMD      rgCmds[],
        MSOCMDTEXT* pcmdtext)
{
    int idm;

    Trace0("CImgHelper::QueryStatus\n");

    Assert(CBase::IsCmdGroupSupported(pguidCmdGroup));
    Assert(cCmds == 1);

    MSOCMD* pCmd = &rgCmds[0];
    HRESULT hr = S_OK;

    Assert(!pCmd->cmdf);

    idm = CBase::IDMFromCmdID(pguidCmdGroup, pCmd->cmdID);
    switch(idm)
    {
    case IDM_ADDFAVORITES:
        {
            TCHAR* pchUrl = NULL;

            pCmd->cmdf = (S_OK==GetUrl(&pchUrl)) ? MSOCMDSTATE_UP :
                MSOCMDSTATE_DISABLED;

            MemFreeString(pchUrl);
            break;
        }

    case IDM_CUT:
        // Enable if script wants to handle it, otherwise leave it to default
        if(!Fire_onbeforecut())
        {
            pCmd->cmdf = MSOCMDSTATE_UP;
        }
        break;

    case IDM_COPY:
        {
            TCHAR* pchPath = NULL;

            // Enable if script wants to handle it or we know there is something to copy,
            // otherwise leave it to default
            if(!Fire_onbeforecopy()
                || (_pBitsCtx && (S_OK==_pBitsCtx->GetFile(&pchPath)))
                || (_pImgCtx && !_pBitsCtx))
            {
                pCmd->cmdf = MSOCMDSTATE_UP;
            }
            MemFreeString(pchPath);
            break;
        }
    case IDM_PASTE:
        // Enable if script wants to handle it, otherwise leave it to default
        if(!Fire_onbeforepaste())
        {
            pCmd->cmdf = MSOCMDSTATE_UP;
        }
        break;

    case IDM_SHOWPICTURE:
        {
            CImgCtx* pImgCtx = GetImgCtx();
            ULONG ulState = pImgCtx ? pImgCtx->GetState()
                : (_pBitsCtx?_pBitsCtx->GetState():IMGLOAD_NOTLOADED);

            pCmd->cmdf = (ulState&IMGLOAD_COMPLETE) ? MSOCMDSTATE_DISABLED : MSOCMDSTATE_UP;
            break;
        }
    case IDM_SAVEPICTURE:
        {
            if((_pBitsCtx&&(_pBitsCtx->GetState()&DWNLOAD_COMPLETE)) ||
                (_pImgCtx&&(_pImgCtx->GetState()&IMGLOAD_COMPLETE)))
            {
                pCmd->cmdf = MSOCMDSTATE_UP;
            }
            else
            {
                pCmd->cmdf = MSOCMDSTATE_DISABLED;
            }
            break;
        }
    case IDM_SETDESKTOPITEM:
        {
            ULONG ulState = _pImgCtx ? _pImgCtx->GetState() : IMGLOAD_NOTLOADED;

            pCmd->cmdf = (ulState&IMGLOAD_COMPLETE) ? MSOCMDSTATE_UP : MSOCMDSTATE_DISABLED;
            break;
        }

    case IDM_DYNSRCPLAY:
        pCmd->cmdf = (_pVideoObj&&(_pVideoObj->CanPlay())) ? MSOCMDSTATE_UP : MSOCMDSTATE_DISABLED;
        break;

    case IDM_DYNSRCSTOP:
        pCmd->cmdf = (_pVideoObj&&_pVideoObj->CanStop()) ? MSOCMDSTATE_UP : MSOCMDSTATE_DISABLED;
        break;

    case IDM_IMGARTPLAY:
    case IDM_IMGARTSTOP:
    case IDM_IMGARTREWIND:
        pCmd->cmdf = MSOCMDSTATE_DISABLED;
        if(_pImgCtx && Doc()->_pOptionSettings->fPlayAnimations)
        {
            CArtPlayer* pArtPlayer = _pImgCtx->GetArtPlayer();

            if(pArtPlayer)
            {
                pCmd->cmdf = pArtPlayer->QueryPlayState(idm) ? MSOCMDSTATE_UP : MSOCMDSTATE_DISABLED;
            }
        }
        break;
    }

    RRETURN_NOTRACE(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgHelper::Exec, public
//
//  Synopsis:   Executes a command on the CImgHelper
//
//----------------------------------------------------------------------------
HRESULT CImgHelper::Exec(
         GUID*          pguidCmdGroup,
         DWORD          nCmdID,
         DWORD          nCmdexecopt,
         VARIANTARG*    pvarargIn,
         VARIANTARG*    pvarargOut)
{
    Trace0("CImgHelper::Exec\n");

    Assert(CBase::IsCmdGroupSupported(pguidCmdGroup));

    int     idm = CBase::IDMFromCmdID(pguidCmdGroup, nCmdID);
    HRESULT hr = MSOCMDERR_E_NOTSUPPORTED;

    switch(idm)
    {
    case IDM_CUT:
        Fire_oncut();
        // marka - we normally want to not allow a cut of an image in BrowseMode
        // but if we're in edit mode - we want to allow it.
        //
        // this bug was introduced as now a site selected image will be current
        // we can probably remove this once we make the currency work again as before
        if(!IsEditable(TRUE))
        {
            hr = S_OK;
        }
        break;
    case IDM_COPY:
        hr = CreateImgDataObject(Doc(), _pImgCtx, _pBitsCtx, _pOwner, NULL);
        Fire_oncopy();
        break;
    case IDM_PASTE:
        Fire_onpaste();

        // MarkA's comments above for IDM_CUT apply here also.
        if(!IsEditable(TRUE))
        {
            hr = S_OK;
        }
        break;
    case IDM_SHOWPICTURE:
        {
            DWNLOADINFO dli;
            CImgCtx* pImgCtx = GetImgCtx();

            Doc()->InitDownloadInfo(&dli);

            if(pImgCtx)
            {
                pImgCtx->SetLoad(TRUE, &dli, TRUE); // Reload on error
            }
            else
            {
                _pOwner->OnPropertyChange(DISPID_CImgElement_src, 0);
            }

            hr = S_OK;
            break;
        }

    case IDM_SAVEPICTURE:
        hr = PromptSaveAs();
        break;

    case IDM_DYNSRCPLAY:
        Replay();
        _fStopped = FALSE;
        hr = S_OK;
        break;
    case IDM_DYNSRCSTOP:
        if (_pVideoObj)
        {
            _pVideoObj->Stop();
            _fStopped = TRUE;
        }
        hr = S_OK;
        break;

    case IDM_IMGARTPLAY:
    case IDM_IMGARTSTOP:
    case IDM_IMGARTREWIND:
        if(_lCookie && _pImgCtx)
        {
            CImgAnim* pImgAnim = GetImgAnim();
            CArtPlayer* pArtPlayer = _pImgCtx->GetArtPlayer();

            if(pImgAnim && pArtPlayer)
            {
                if(idm == IDM_IMGARTPLAY)
                {
                    pImgAnim->StartAnim(_lCookie);
                    pArtPlayer->DoPlayCommand(IDM_IMGARTPLAY);
                }
                else if (idm == IDM_IMGARTSTOP)
                {
                    pArtPlayer->DoPlayCommand(IDM_IMGARTSTOP);
                    pImgAnim->StopAnim(_lCookie);
                }
                else // idm == IDM_IMGARTREWIND
                {
                    pArtPlayer->DoPlayCommand(IDM_IMGARTREWIND);
                    pImgAnim->StartAnim(_lCookie);
                    pArtPlayer->_fRewind = TRUE;
                }
            }
        }
        break;
    }

    RRETURN_NOTRACE(hr);
}

//+-------------------------------------------------------------------
//
//  Member  : ShowImgContextMenu
//
//  synopsis   : Implementation of interface src property get. this
//      should return the expanded src (e.g.  file://c:/temp/foo.jpg
//      rather than foo.jpg)
//
//------------------------------------------------------------------
HRESULT CImgHelper::ShowImgContextMenu(CMessage* pMessage)
{
    HRESULT hr;

    Assert(pMessage);
    Assert(WM_CONTEXTMENU == pMessage->message);

    if(_pImgCtx && _pImgCtx->GetMimeInfo() && _pImgCtx->GetMimeInfo()->pfnImg==NewImgTaskArt)
    {
        hr = _pOwner->OnContextMenu(
            MAKEPOINTS(pMessage->lParam).x,
            MAKEPOINTS(pMessage->lParam).y,
            (!IsEditable(TRUE) && (_pImgCtx->GetArtPlayer())) ?
            (CONTEXT_MENU_IMGART) : (CONTEXT_MENU_IMAGE));
    }
    else
    {
        LPCTSTR pchdynSrc = GetAAdynsrc();
        hr = _pOwner->OnContextMenu(
            MAKEPOINTS(pMessage->lParam).x,
            MAKEPOINTS(pMessage->lParam).y,
            (!IsEditable(TRUE) && (pchdynSrc && pchdynSrc[0]) ) ?
            (CONTEXT_MENU_IMGDYNSRC) : (CONTEXT_MENU_IMAGE));
    }
    RRETURN(hr);
}


//+-------------------------------------------------------------------
//
//  Member  : get_src
//
//  synopsis   : Implementation of interface src property get. this
//      should return the expanded src (e.g.  file://c:/temp/foo.jpg
//      rather than foo.jpg)
//
//------------------------------------------------------------------
STDMETHODIMP CImgHelper::get_src(BSTR* pstrFullSrc)
{
    HRESULT hr;
    TCHAR   cBuf[pdlUrlLen];
    TCHAR*  pchNewUrl = cBuf;

    if(!pstrFullSrc)
    {
        hr = E_POINTER;
        goto Cleanup;
    }

    *pstrFullSrc = NULL;

    hr = Doc()->ExpandUrl(GetAAsrc(), ARRAYSIZE(cBuf), pchNewUrl, _pOwner);
    if(hr || (pchNewUrl==NULL))
    {
        goto Cleanup;
    }

    *pstrFullSrc = SysAllocString(pchNewUrl);
    if(!*pstrFullSrc)
    {
        hr = E_OUTOFMEMORY;
    }

Cleanup:
    RRETURN(hr);
}

//+------------------------------------------------------------------
//
//  member : CImgHelper::GetFile
//
//-------------------------------------------------------------------
HRESULT CImgHelper::GetFile(TCHAR** ppchFile)
{
    HRESULT hr = E_FAIL;

    Assert(ppchFile);

    *ppchFile = NULL;

    if(_pBitsCtx)
    {
        if(_pBitsCtx->GetState() & DWNLOAD_COMPLETE)
        {
            hr = _pBitsCtx->GetFile(ppchFile);
        }
        else
        {
            hr = S_OK;
        }
    }
    else if(_pImgCtx)
    {
        if(_pImgCtx->GetState() & DWNLOAD_COMPLETE)
        {
            hr = _pImgCtx->GetFile(ppchFile);
        }
        else
        {
            hr = S_OK;
        }
    }
    else
    {
        // BUGBUG DOM May be we could return anything more informative
    }

    RRETURN(hr);
}

//+------------------------------------------------------------------
//
//  member : CImgHelper::GetUrl
//
//-------------------------------------------------------------------
HRESULT CImgHelper::GetUrl(TCHAR** ppchUrl)
{
    HRESULT hr = E_FAIL;

    Assert(ppchUrl);

    *ppchUrl = NULL;

    if(_pBitsCtx)
    {
        hr = MemAllocString(_pBitsCtx->GetUrl(), ppchUrl);
    }
    else if(_pImgCtx)
    {
        hr = MemAllocString(_pImgCtx->GetUrl(), ppchUrl);
    }

    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CImgHelper::PromptSaveAs
//
//  Synopsis:   Brings up the 'Save As' dialog and saves the image.
//
//  Arguments:  pchPath     Returns the file to whic the image was saved.
//
//----------------------------------------------------------------------------
HRESULT CImgHelper::PromptSaveAs(TCHAR* pchFileName/*=NULL*/, int cchFileName/*=0*/)
{
    /*if(_pBitsCtx && (_pBitsCtx->GetState()&DWNLOAD_COMPLETE))
    {
        Doc()->SaveImgCtxAs(NULL, _pBitsCtx, IDM_SAVEPICTURE, pchFileName, cchFileName);
    }
    else if(_pImgCtx && (_pImgCtx->GetState()&DWNLOAD_COMPLETE))
    {
        Doc()->SaveImgCtxAs(_pImgCtx, NULL, IDM_SAVEPICTURE, pchFileName, cchFileName);
    } wlw note*/
    RRETURN(S_OK);
}

long CImgHelper::GetAAHspace()
{
    long lHSpace = GetAAhspace();

    if(lHSpace == -1)
    {
        lHSpace = 0;
    }

    return lHSpace;
}

void CImgHelper::GetMarginInfo(
       CParentInfo* ppri,
       LONG* plLMargin,
       LONG* plTMargin,
       LONG* plRMargin,
       LONG* plBMargin)
{
    long lhSpace = 0;
    long lvSpace = 0;
    BOOL fUseDefMargin = _pOwner->IsAligned() && !_pOwner->IsAbsolute() && !IsHSpaceDefined();

    if(plLMargin || plRMargin)
    {
        long lhMargin=0;

        lhSpace = GetAAhspace();
        if(lhSpace < 0)
        {
            lhSpace = 0;
        }
        if(plLMargin)
        {
            *plLMargin += lhSpace;
            lhMargin = *plLMargin;
        }
        if(plRMargin)
        {
            *plRMargin += lhSpace;
            lhMargin += *plRMargin;
        }

        if(lhMargin > 0)
        {
            fUseDefMargin = FALSE;
        }
    }

    if(plTMargin || plBMargin)
    {
        long lvMargin=0;

        lvSpace = GetAAvspace();
        if(lvSpace < 0)
        {
            lvSpace = 0;
        }
        if(plTMargin)
        {
            lvMargin = *plTMargin;
            *plTMargin += lvSpace;
        }
        if(plBMargin)
        {
            lvMargin += *plBMargin;
            *plBMargin += lvSpace;
        }

        // if vertical margins are defined,
        // Netscape compatibility should goes away.
        // but if just vspace is defined,
        // we still need to preserve this compatibility
        if(lvMargin > 0)
        {
            fUseDefMargin = FALSE;
        }
    }

    if(fUseDefMargin && (plLMargin || plRMargin))
    {
        // (srinib/yinxie) - netscape compatibility, aligned images have
        // a 3 pixel hspace by default
        Assert(ppri);
        lhSpace = ppri->DocPixelsFromWindowX(DEF_HSPACE);
        if(plLMargin)
        {
            *plLMargin += lhSpace;
        }
        if(plRMargin)
        {
            *plRMargin += lhSpace;
        }
    }
}

void CImgHelper::CalcSize(CCalcInfo* pci, SIZE* psize)
{
    SIZE sizeImg;
    CBorderInfo bi;

    sizeImg = _afxGlobalData._Zero.size;

    Assert(pci);
    Assert(psize);

    SIZE sizeBorderHVSpace2;
    CUnitValue uvWidth  = GetFirstBranch()->GetCascadedwidth();
    CUnitValue uvHeight = GetFirstBranch()->GetCascadedheight();
    long lParentWidth   = psize->cx;
    long lParentHeight  = pci->_sizeParent.cy;
    BOOL fMaxWidth      = (uvWidth.GetUnitType()==CUnitValue::UNIT_PERCENT) &&
        ((pci->_smMode == SIZEMODE_MMWIDTH)
        ||(pci->_smMode == SIZEMODE_MINWIDTH));
    BOOL fMaxHeight     = (uvHeight.GetUnitType()==CUnitValue::UNIT_PERCENT) &&
        (  (pci->_smMode == SIZEMODE_MMWIDTH)
        || (pci->_smMode == SIZEMODE_MINWIDTH));
    SIZE sizeNew = { 0, 0 }, sizePlaceHolder;
    BOOL fHasSize;
    ULONG ulState;
    TCHAR* pchAlt = NULL;
    RECT rcText;

    _pOwner->GetBorderInfo(pci, &bi, FALSE);

    sizeBorderHVSpace2.cx = pci->DocPixelsFromWindowX(bi.aiWidths[BORDER_LEFT])
        + pci->DocPixelsFromWindowX(bi.aiWidths[BORDER_RIGHT]);
    sizeBorderHVSpace2.cy = pci->DocPixelsFromWindowY(bi.aiWidths[BORDER_TOP])
        + pci->DocPixelsFromWindowY(bi.aiWidths[BORDER_BOTTOM]);
    lParentWidth          = max(0L, lParentWidth-sizeBorderHVSpace2.cx);
    lParentHeight         = max(0L, lParentHeight-sizeBorderHVSpace2.cy);
    if(_pBitsCtx && _pVideoObj)
    {
        ulState = _pBitsCtx->GetState(); // BUGBUG (lmollico): ulState should be an IMGLOAD_*
        fHasSize = OK(_pVideoObj->GetSize(&sizeImg));
    }
    else if(_pImgCtx)
    {
        if(!_fSizeChange)
        {
            _pImgCtx->SelectChanges(IMGCHG_SIZE, 0, FALSE);

            _fSizeChange = TRUE;
        }
        ulState = _pImgCtx->GetState(FALSE, &sizeImg);
        fHasSize = (sizeImg.cx || sizeImg.cy);
    }
    else
    {
        ulState = IMGLOAD_NOTLOADED;
        fHasSize = FALSE;
        _fNeedSize = TRUE;
    }

    /*if(!fHasSize && pci->_fTableCalcInfo)
    {
    CTableCalcInfo* ptci = (CTableCalcInfo*)pci;
    ptci->_fDontSaveHistory = TRUE;
    } wlw note*/

    if(fHasSize ||
        (_pImgCtx && !_fExpandAltText && GetCachedImageSize(_pImgCtx->GetUrl(), &sizeImg)))
    {
        sizePlaceHolder.cx = sizeNew.cx = pci->DocPixelsFromWindowX(sizeImg.cx);
        sizePlaceHolder.cy = sizeNew.cy = pci->DocPixelsFromWindowY(sizeImg.cy);
    }
    else
    {
        SIZE sizeGrab;
        GetPlaceHolderBitmapSize(ulState&(IMGLOAD_ERROR|IMGLOAD_STOPPED), &sizeNew);
        sizeNew.cx = pci->DocPixelsFromWindowX(sizeNew.cx);
        sizeNew.cy = pci->DocPixelsFromWindowY(sizeNew.cy);
        sizeGrab.cx = pci->DocPixelsFromWindowX(GRABSIZE);
        sizeGrab.cy = pci->DocPixelsFromWindowY(GRABSIZE);

        pchAlt = (TCHAR*)GetAAalt();

        if(pchAlt && *pchAlt)
        {
            CDocument* pDoc = Doc();
            const CCharFormat* pCF = GetFirstBranch()->GetCharFormat();

            CIntlFont intlfont(pci->_hdc,
                pDoc->GetCodePage(),
                pCF ? pCF->_lcid : 0,
                pDoc->_sBaselineFont,
                pchAlt);

            rcText.left = rcText.top = rcText.right = rcText.bottom = 0;
            DrawTextInCodePage(WindowsCodePageFromCodePage(pDoc->GetCodePage()),
                pci->_hdc, pchAlt, -1, &rcText, DT_CALCRECT|DT_NOPREFIX);

            sizePlaceHolder.cx = sizeNew.cx + 3*sizeGrab.cx + rcText.right - rcText.left;
            sizePlaceHolder.cy = max(sizeNew.cy, rcText.bottom - rcText.top) + 2*sizeGrab.cy;
        }
        else
        {
            sizePlaceHolder.cx = sizeNew.cx + 2*sizeGrab.cx;
            sizePlaceHolder.cy = sizeNew.cy + 2*sizeGrab.cy;
        }
    }
    if(uvWidth.IsNullOrEnum() || uvHeight.IsNullOrEnum())
    {
        // If the image Width is set, use it
        if(!uvWidth.IsNullOrEnum() && !fMaxWidth)
        {
            psize->cx = uvWidth.XGetPixelValue(
                pci,
                lParentWidth,
                GetFirstBranch()->GetFontHeightInTwips(&uvWidth));
        }
        else
        {
            // if height only is set, then the image should be proportional
            // to the real size sizeNew.cx
            if(!uvHeight.IsNullOrEnum() && !fMaxHeight && (sizeNew.cy>0))
            {
                psize->cx = MulDivQuick(uvHeight.YGetPixelValue(pci,
                    lParentHeight,
                    GetFirstBranch()->GetFontHeightInTwips(&uvHeight)),
                    sizeNew.cx,
                    sizeNew.cy);
            }
            else
            {
                psize->cx = sizePlaceHolder.cx;
            }
        }

        // If the image Height is set, use it
        if(!uvHeight.IsNullOrEnum() && !fMaxHeight)
        {
            psize->cy = uvHeight.YGetPixelValue(pci,
                lParentHeight,
                GetFirstBranch()->GetFontHeightInTwips(&uvHeight));
        }
        else
        {
            // if width only is set, then the image should be proportional
            // to the real size sizeNew.cx
            if(!uvWidth.IsNullOrEnum() && !fMaxWidth && (sizeNew.cx>0))
            {
                psize->cy = MulDivQuick(uvWidth.XGetPixelValue(pci,
                    lParentWidth,
                    GetFirstBranch()->GetFontHeightInTwips(&uvWidth)),
                    sizeNew.cy,
                    sizeNew.cx);
            }
            else
            {
                psize->cy = sizePlaceHolder.cy;
            }
        }

    }
    else
    {
        if(!fMaxWidth)
        {
            psize->cx = uvWidth.XGetPixelValue(pci,
                lParentWidth,
                GetFirstBranch()->GetFontHeightInTwips(&uvWidth));
        }
        else
        {
            psize->cx = sizePlaceHolder.cx;
        }
        if(!fMaxHeight)
        {
            psize->cy = uvHeight.YGetPixelValue(pci,
                lParentHeight,
                GetFirstBranch()->GetFontHeightInTwips(&uvHeight));
        }
        else
        {
            psize->cy = sizePlaceHolder.cy;
        }
    }

    if(_fExpandAltText && pchAlt && *pchAlt && !fHasSize)
    {
        if(psize->cx < sizePlaceHolder.cx)
        {
            psize->cx = sizePlaceHolder.cx;
        }
        if(psize->cy < sizePlaceHolder.cy)
        {
            psize->cy = sizePlaceHolder.cy;
        }
    }

    if(psize->cx > 0)
    {
        psize->cx += sizeBorderHVSpace2.cx;
    }
    else
    {
        psize->cx = sizeBorderHVSpace2.cx;
    }
    if(psize->cy > 0)
    {
        psize->cy += sizeBorderHVSpace2.cy;
    }
    else
    {
        psize->cy = sizeBorderHVSpace2.cy;
    }
}

HRESULT CImgHelper::CacheImage(HDC hdc, CRect* prcDst, SIZE* pSize, DWORD dwFlags, ULONG ulState)
{
    HRESULT hr = S_OK;
    HDC     hdcMem = NULL;
    HBITMAP hbmSav = NULL;

    hdcMem = GetMemoryDC();
    if(hdcMem == NULL)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    _xWidCache = prcDst->Width();
    _yHeiCache = prcDst->Height();

    _colorMode = GetDefaultColorMode();

    if(_hbmCache)
    {
        DeleteObject(_hbmCache);
    }

    _hbmCache = CreateCompatibleBitmap(hdc, _xWidCache, _yHeiCache);
    if(_hbmCache == NULL)
    {
        hr = E_OUTOFMEMORY;
        goto Cleanup;
    }

    hbmSav = (HBITMAP)SelectObject(hdcMem, _hbmCache);

    if(!(ulState & IMGTRANS_OPAQUE))
    {
        dwFlags |= DRAWIMAGE_NOTRANS;
    }

    _pImgCtx->DrawEx(hdcMem, prcDst, dwFlags);

Cleanup:
    if(hbmSav)
    {
        SelectObject(hdcMem, hbmSav);
    }
    if(hdcMem)
    {
        ReleaseMemoryDC(hdcMem);
    }

    RRETURN(hr);
}

void CImgHelper::DrawCachedImage(HDC hdc, CRect* prcDst, DWORD dwFlags, ULONG ulState)
{
    HDC     hdcMem = NULL;
    HBITMAP hbmSav = NULL;
    DWORD   dwRop;

    hdcMem = GetMemoryDC();
    if(hdcMem == NULL)
    {
        goto Cleanup;
    }

    hbmSav = (HBITMAP)SelectObject(hdcMem, _hbmCache);

    if(ulState & IMGTRANS_OPAQUE)
    {
        dwRop = SRCCOPY;
    }
    else
    {
        _pImgCtx->DrawEx(hdc, prcDst, dwFlags|DRAWIMAGE_MASKONLY);

        dwRop = SRCAND;
    }

    BitBlt(hdc, prcDst->left, prcDst->top,
        prcDst->Width(), prcDst->Height(),
        hdcMem, prcDst->left, prcDst->top, dwRop);

Cleanup:
    if(hbmSav)
    {
        SelectObject(hdcMem, hbmSav);
    }
    if(hdcMem)
    {
        ReleaseMemoryDC(hdcMem);
    }
}

//+---------------------------------------------------------------------------
//
//  Member:     Draw
//
//  Synopsis:   Paint the object. Note that this function does not save draw
//              info. Derived classes must override Draw() and save draw info
//               before calling this function.
//
//----------------------------------------------------------------------------
void CImgHelper::Draw(CFormDrawInfo* pDI)
{
    ULONG   ulState;
    HDC     hdc = pDI->GetDC(TRUE);

    if(_pBitsCtx && _pVideoObj)
    {
        ulState = _pBitsCtx->GetState();
    }
    else if(_pImgCtx)
    {
        ulState = _pImgCtx->GetState();
    }
    else
    {
        ulState = IMGLOAD_NOTLOADED;
    }

    if(!(ulState&(IMGBITS_NONE|IMGLOAD_NOTLOADED|IMGLOAD_ERROR)) && _fHideForSecurity)
    {
        ulState = IMGLOAD_ERROR;
    }

    if(ulState&(IMGBITS_NONE|IMGLOAD_NOTLOADED|IMGLOAD_ERROR))
    {
        SIZE sizeGrab = { GRABSIZE, GRABSIZE };
        CDocument* pDoc = Doc();
        BOOL fMissing = !!(ulState & IMGLOAD_ERROR);

        if(fMissing || !GetImgCtx() || !pDoc->_pOptionSettings->fShowImages || pDoc->_pOptionSettings->fShowImagePlaceholder)
        {
            const CCharFormat* pCF = GetFirstBranch()->GetCharFormat();
            COLORREF fgColor = (pCF && pCF->_ccvTextColor.IsDefined()) ?
                pCF->_ccvTextColor.GetColorRef() : RGB(0, 0, 0);

            DrawPlaceHolder(hdc, pDI->_rc,
                (LPTSTR)GetAAalt(),
                pDoc->GetCodePage(), pCF ? pCF->_lcid : 0, pDoc->_sBaselineFont,
                &sizeGrab, fMissing,
                fgColor, _pOwner->GetBackgroundColor(), NULL, FALSE, pDI->DrawImageFlags());
        }
    }
    else if(_lCookie && _fIsInPlace)
    {
        CImgAnim* pImgAnim = GetImgAnim();

        if(pImgAnim)
        {
            _pImgCtx->DrawFrame(hdc, pImgAnim->GetImgAnimState(_lCookie), &pDI->_rc, NULL, NULL, pDI->DrawImageFlags());
        }
    }
    else if(_pImgCtx)
    {
        CRect* prcDst = &pDI->_rc;
        DWORD dwFlags = pDI->DrawImageFlags();

        if(_fCache)
        {
            SIZE size;
            LONG xWidDst = prcDst->Width();
            LONG yHeiDst = prcDst->Height();

            _pImgCtx->GetState(FALSE, &size);

            if((ulState&IMGLOAD_COMPLETE)
                && (size.cx!=xWidDst || size.cy!=yHeiDst)
                && (size.cx!=1 || size.cy!=1))
            {
                HRESULT hr = S_OK;

                if(_xWidCache!=xWidDst
                    || _yHeiCache!=yHeiDst
                    || _colorMode!=GetDefaultColorMode())
                {
                    hr = CacheImage(hdc, prcDst, &size, dwFlags, ulState);
                }

                if(hr == S_OK)
                {
                    DrawCachedImage(hdc, prcDst, dwFlags, ulState);
                    return;
                }
            }
        }

        _pImgCtx->DrawEx(hdc, prcDst, dwFlags);
    }
}