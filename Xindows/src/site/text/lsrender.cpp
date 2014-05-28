
#include "stdafx.h"
#include "lsrender.h"

#include "MarginInfo.h"

#include "msls/lsulinfo.h"
#include "msls/lsqline.h"
#include "msls/lsdsply.h"

// This is the number of characters we allocate by default on the stack
// for string buffers. Only if the number of characters in a string exceeds
// this will we need to allocate memory for string buffers.
#define STACK_ALLOCED_BUF_SIZE  256

/*
*  LSReExtTextOutW(CCcs *pccs, hdc, x, y, fuOptions, lprc, lpString, cbCount,lpDx)
*
*  @mfunc
*      Patch around the Win95 FE bug.
*
*  @rdesc
*      Returns whatever ExtTextOut returns
*/
BOOL LSReExtTextOutW(
                     CCcs* pccs,                 //@parm the font
                     HDC hdc,                    //@parm handle to device context
                     int xp,                     //@parm x-coordinate of reference point
                     int yp,                     //@parm y-coordinate of reference point
                     UINT fuOptions,             //@parm text-output options
                     CONST RECT* lprect,         //@parm optional clipping and/or opaquing rectangle
                     const WCHAR* lpwchString,   //@parm points to string
                     UINT cchCount,              //@parm number of characters in string
                     CONST INT* lpDx)            //@parm Ptr to array of intercharacter spacing values
{
    // NB (cthrash) Ported from RichEdit rel0144 11/22/96
    // This is a protion of Word code adapted for our needs.
    // This is a work around for Win95FE bugs that cause GPFs in GDI if multiple
    // characters above Unicode 0x7F are passed to ExtTextOutW.
    int          cch;
    const WCHAR* lpwchT = lpwchString;
    const WCHAR* lpwchStart = lpwchT;
    const WCHAR* lpwchEnd = lpwchString + cchCount;
    const int*   lpdxpCur;
    BOOL         fRet = 0;
    LONG         lWidth=0;
    UINT         uiTA = GetTextAlign(hdc);

    if(uiTA & TA_UPDATECP)
    {
        POINT pt;

        SetTextAlign(hdc, uiTA&~TA_UPDATECP);

        GetCurrentPositionEx(hdc, &pt);

        xp = pt.x;
        yp = pt.y;
    }

    while(lpwchT < lpwchEnd)
    {
        // characters less than 0x007F do not need special treatment
        // we output then in contiguous runs
        if(*lpwchT > 0x007F)
        {
            if((cch=lpwchT-lpwchStart) > 0)
            {
                lpdxpCur = lpDx ? lpDx+(lpwchStart-lpwchString) : NULL;

                // Output the run of chars less than 0x7F
                fRet = ::ExtTextOutW(hdc, xp, yp, fuOptions, lprect,
                    lpwchStart, cch, (int*)lpdxpCur);
                if(!fRet)
                {
                    return fRet;
                }

                fuOptions &= ~ETO_OPAQUE; // Don't erase mutliple times!!!

                // Advance
                if(lpdxpCur)
                {
                    while(cch--)
                    {
                        xp += *lpdxpCur++;
                    }
                }
                else
                {
                    while(cch)
                    {
                        pccs->Include(lpwchStart[--cch], lWidth);
                        xp += lWidth;
                    }
                }

                lpwchStart = lpwchT;
            }

            // Output chars above 0x7F one at a time to prevent Win95 FE GPF
            lpdxpCur = lpDx ? lpDx+(lpwchStart-lpwchString) : NULL;

            fRet = ::ExtTextOutW(hdc, xp, yp, fuOptions, lprect, 
                lpwchStart, 1, (int*)lpdxpCur);

            if(!fRet)
            {
                return fRet;
            }

            fuOptions &= ~ETO_OPAQUE; // Don't erase mutliple times!!!

            // Advance
            if(lpdxpCur)
            {
                xp += *lpdxpCur;
            }
            else
            {
                pccs->Include(*lpwchStart, lWidth);

                xp += lWidth;
            }

            lpwchStart++;
        }

        lpwchT++;
    }

    // output the final run; also, if we were called with cchCount == 0,
    // make a call here to erase the rectangle
    if((cch=lpwchT-lpwchStart)>0 || !cchCount)
    {
        fRet = ::ExtTextOutW(hdc, xp, yp, fuOptions, lprect, lpwchStart, cch,
            (int*)(lpDx?lpDx+(lpwchStart-lpwchString):NULL));
    }

    if(uiTA & TA_UPDATECP)
    {
        SetTextAlign(hdc, uiTA);

        // Update the xp for the final run.
        AssertSz(!lpDx, "In blast mode -- lpDx should be null.");

        while(cch)
        {
            pccs->Include(lpwchStart[--cch], lWidth);

            xp += lWidth;
        }                

        MoveToEx(hdc, xp, yp, NULL);
    }

    return fRet;
}

//-----------------------------------------------------------------------------
//
//  Function:   WideCharToMultiByteForSymbol
//
//  Purpose:    This function is a hacked-up version of WC2MB that we use to
//              convert when rendering in the symbol font.  Since symbol fonts
//              exist in a codepage-free world, we need to always convert to
//              multibyte to get the desired result.  The only problem comes
//              when you've got content such as the following:
//
//                  <font face="Wingdings">&#171;</font>
//
//              This, incidentally, is the recommended way of creating a little
//              star in your document.  We wouldn't need this hacky converter
//              if the user entered the byte value for 171 instead of the named
//              entity -- this however is problematic for DBCS locales when
//              the byte value may be a leadbyte (and thus would absorb the
//              subsequent byte.)
//
//              Anyway, the hack here is to 'convert' all Unicode codepoints
//              less than 256 by stripping the MSB, and converting normally
//              the rest.
//
//              Note the other unfortunate hack is that we need to convert
//              the offset array
//
//-----------------------------------------------------------------------------
int WideCharToMultiByteForSymbolQuick(UINT, DWORD, LPCWSTR, int, LPSTR, int);
int WideCharToMultiByteForSymbolSlow(UINT, DWORD, LPCWSTR, int, LPSTR, int, const int*, int*);

int WideCharToMultiByteForSymbol(
                             UINT       uiCodePage,
                             DWORD      dwFlags,
                             LPCWSTR    pch,
                             int        cch,
                             LPSTR      pb,
                             int        cb,
                             const int* lpDxSrc,
                             int*       lpDxDst)
{
    AssertSz(cch!=-1, "Don't pass -1 here.");
    int iRet;

    if(lpDxSrc)
    {
        iRet = WideCharToMultiByteForSymbolSlow(uiCodePage, dwFlags, pch, cch, pb, cb, lpDxSrc, lpDxDst);
    }
    else
    {
        iRet = WideCharToMultiByteForSymbolQuick(uiCodePage, dwFlags, pch, cch, pb, cb);
    }

    return iRet;
}

int WideCharToMultiByteForSymbolQuick(
                                  UINT      uiCodePage,
                                  DWORD     dwFlags,
                                  LPCWSTR   pch,
                                  int       cch,
                                  LPSTR     pb,
                                  int       cb)
{
    // This is the quick pass, where we don't need to readjust the width array (lpDx)
    const WCHAR* pchStop = pch + cch;
    const WCHAR* pchStart = NULL;
    const char*  pbStart = pb;
    const char*  pbStop = pb + cb;

    for(; pch<pchStop; pch++)
    {
        TCHAR ch = *pch;

        if(ch > 255)
        {
            const BYTE b = InWindows1252ButNotInLatin1(ch);

            if(b)
            {
                ch = b;
            }
        }

        if(ch > 255)
        {
            // Accumulate the non Latin-1 characters -- remember the start
            pchStart = pchStart ? pchStart : pch;
        }
        else
        {
            if(pchStart)
            {
                // We have accumulated some non-Latin1 chars -- convert these first
                const int cb = WideCharToMultiByte(
                    uiCodePage, dwFlags,
                    pchStart, pch-pchStart,
                    pb, pbStop-pb,
                    NULL, NULL);

                pb += cb;

                pchStart = NULL;
            }

            if(pb < pbStop)
            {
                // Tack on the Latin1 character
                *pb++ = ch;
            }
            else
            {
                break;
            }
        }
    }

    if(pchStart)
    {
        // Take care of non-Latin1 chars at the end of the string
        const int cb = WideCharToMultiByte(
            uiCodePage, dwFlags,
            pchStart, pch-pchStart,
            pb, pbStop-pb,
            NULL, NULL);

        pb += cb;
    }

    return pb - pbStart;
}

int WideCharToMultiByteForSymbolSlow(
                                 UINT       uiCodePage,
                                 DWORD      dwFlags,
                                 LPCWSTR    pch,
                                 int        cch,
                                 LPSTR      pb,
                                 int        cb,
                                 const int* lpDxSrc,
                                 int*       lpDxDst)
{
    // This is the slow pass, where we need to readjust the width array (lpDx)
    // Note lpDxDst is assumed to have (at least) cb bytes in it.
    const WCHAR* pchStop = pch + cch;
    const char*  pbStart = pb;
    const char*  pbStop = pb + cb;

    while(pch<pchStop && pb<pbStop)
    {
        const TCHAR ch = *pch++;

        if(ch < 256)
        {
            *pb++ = char(ch);
            *lpDxDst++ = *lpDxSrc++;
        }
        else
        {
            const BYTE b = InWindows1252ButNotInLatin1(ch);

            if(b)
            {
                *pb++ = b;
                *lpDxDst++ = *lpDxSrc++;
            }
            else
            {
                const int cb = WideCharToMultiByte(
                    uiCodePage, dwFlags,
                    &ch, 1,
                    pb, pbStop-pb,
                    NULL, NULL);

                *lpDxDst = *lpDxSrc++;
                lpDxDst += cb;
                pb += cb;
            }
        }
    }

    return pb-pbStart;
}

#define MAX_CHUNK   4000L

DWORD GetMaxRenderLine()
{
    static DWORD dwMaxRenderLine = 0;
    HKEY hkInetSettings;
    DWORD dwVal;
    DWORD dwSize;
    long lResult;

    if(!dwMaxRenderLine)
    {
        dwMaxRenderLine = MAX_CHUNK;
        /*if(g_dwPlatformID == VER_PLATFORM_WIN32_NT) wlw note*/
        {
            lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _T("Software\\Microsoft\\Internet Explorer\\Main"), 0, KEY_QUERY_VALUE, &hkInetSettings);
            if(ERROR_SUCCESS == lResult)  
            {
                dwSize = sizeof(dwVal);
                lResult = RegQueryValueEx(hkInetSettings, _T("MaxRenderLine"), 0, NULL, (LPBYTE)&dwVal, &dwSize);
                if(ERROR_SUCCESS == lResult) 
                {
                    dwMaxRenderLine = dwVal;
                }
                RegCloseKey(hkInetSettings);
            }
        }
    }

    return dwMaxRenderLine;
}

/*
*  LSReExtTextOut
*
*  @mfunc
*      Dispatchs to ExtTextOut that is appropriate for the O/S.
*
*  @rdesc
*      Returns whatever ExtTextOut returns
*/
BOOL LSReExtTextOut(
               CCcs* pccs,                   //@parm the font
               HDC hdc,                      //@parm handle to device context
               int x,                        //@parm x-coordinate of reference point
               int y,                        //@parm y-coordinate of reference point
               UINT fuOptions,               //@parm text-output options
               CONST RECT* lprc,             //@parm optional clipping and/or opaquing rectangle
               const WCHAR* pchString,       //@parm points to string
               long cchString,               //@parm number of characters in string
               CONST INT* lpDx,              //@parm pointer to array of intercharacter spacing values
               CONVERTMODE cm)               //@parm CM_NONE, CM_SYMBOL, CM_MULTIBYTE, CM_FEONNONFE
{
    BOOL fRetVal = TRUE;
    UINT uiCodePage = pccs->GetBaseCcs()->_sCodePage;

    // Win95 has trouble drawing long strings, even drawing large coordinates.
    // So, we just stop after a few thousand characters.
    cchString = min((long)GetMaxRenderLine(), cchString);

    // For EUDC-chars (which must be in their own text run, and thus it's own
    // LsReExtTextOut call) cannot be rendered in Win95 through ExtTextOutW; it
    // must be rendered using ExtTextOutA.  For Win98, NT4, and W2K, ExtTextOutW
    // will work just fine.
    if(CVM_FEONNONFE==cm && cchString && IsEUDCChar(pchString[0]))
    {
        // NB (cthrash) We've seen emprically that Win98-PRC is not up to
        // the task, so continue to call ETOA for this platform.  Lame.
        cm = CVM_NONE;
    }

    if(CVM_SYMBOL==cm || CVM_MULTIBYTE==cm)
    {
        // If we are doing LOWBYTE or MULTIBYTE conversion,
        // unless we have a FE font on a non-FE machine.
        CStackDataAry<char, STACK_ALLOCED_BUF_SIZE> aryBuf;
        CStackDataAry<int, STACK_ALLOCED_BUF_SIZE> aryDx;

        char* pbString;
        int* lpDxT = NULL;
        int cbString = 0;
        BOOL fDoDxMap = lpDx && cchString>1;

        // Double the buffer size, unless in LOWBYTE mode
        int cbBuffer = cchString * 2;

        // Get a buffer as big as it needs to be.
        if(aryBuf.Grow(cbBuffer)!=S_OK
            || (lpDx && aryDx.Grow(cbBuffer)!=S_OK))
        {
            cbBuffer = STACK_ALLOCED_BUF_SIZE;
        }
        pbString = aryBuf;

        if(lpDx)
        {
            lpDxT = aryDx;
        }

        // BUGBUG (cthrash) We should really clean this code up a bit --
        // The only difference between CM_MULTIBYTE and CM_SYMBOL as it
        // currently exists is the handling of the U+0080-U+009F range.
        if(cm == CVM_MULTIBYTE)
        {
            // Convert string
            cbString = WideCharToMultiByte(uiCodePage, 0, pchString, cchString,
                pbString, cbBuffer, NULL, NULL);

            if(lpDx && !fDoDxMap)
            {
                lpDxT[0] = *lpDx;
                lpDxT[1] = 0;
            }
        }
        else
        {
            // Note here that we do the lpDx conversion as we go, unlike the
            // CVM_MULTIBYTE scenario.
            cbString = WideCharToMultiByteForSymbol(
                uiCodePage, 0,
                pchString, cchString,
                pbString, cbBuffer,
                lpDx, lpDxT);

            // If we were successful, we don't need to map the lpDx
            fDoDxMap &= cbString == 0;
        }

        if(!cbString)
        {
            // The conversion failed for one reason or another.  We should
            // make every effort to use WCTMB before we fall back to
            // taking the low-byte of every wchar (below), otherwise we
            // risk dropping the high-bytes and displaying garbage.

            // Use the cpg from the font, since the uiCodePage passed is
            //  the requested codepage and the font-mapper may very well
            //  have mapped to a different one.
            TEXTMETRIC tm;
            UINT uAltCodePage = GetLatinCodepage();

            if(GetTextMetrics(hdc, &tm) && tm.tmCharSet!=DEFAULT_CHARSET)
            {
                UINT uFontCodePage = DefaultCodePageFromCharSet(tm.tmCharSet,
                   _afxGlobalData._cpDefault, 0);

                if(uFontCodePage != uiCodePage)
                {
                    uAltCodePage = uFontCodePage;
                }
            }

            if(uAltCodePage != uiCodePage)
            {
                cbString = WideCharToMultiByte(
                    uAltCodePage, 0, pchString, cchString,
                    pbString, cbBuffer, NULL, NULL);
                uiCodePage = uAltCodePage;
            }
        }

        if(!cbString) // Convert WCHARs to CHARs
        {
            int cbCopy;

            // FUTURE:  We come here for both SYMBOL_CHARSET fonts and for
            // DBCS bytes stuffed into wchar's (one byte per wchar) when
            // the requested code page is not installed on the machine and
            // the MBTWC fails. Instead, we could have another conversion
            // mode that collects each DBCS char as a single wchar and then
            // remaps to a DBCS string for ExtTextOutA. This would allow us
            // to display text if the system has the right font even tho it
            // doesn't have the right cpg.

            // If we are converting this WCHAR buffer in this manner
            // (by taking only the low-byte's of the WCHAR's), it is
            // because:
            //  1) cm == CVM_SYMBOL
            //  2) WCTMB above failed for some reason or another.  It may
            //      be the case that the string is entirely ASCII in which
            //      case dropping the high-bytes is not a big deal (otherwise
            //      we assert).

            cbCopy = cbString = cchString;

            while(cbCopy--)
            {
                AssertSz(pchString[cbCopy]<=0xFF, "LSReExtTextOut():  Found "
                    "a WCHAR with non-zero high-byte when using "
                    "CM_SYMBOL conversion mode.");
                pbString[cbCopy] = pchString[cbCopy];
            }

            fDoDxMap = FALSE;
            lpDxT = (int*)lpDx;
        }

        // This is really ugly -- our lpDx array is for wide chars, and we
        // may need to convert it for multibyte chars.  Do this only when
        // necessary.
        if(fDoDxMap)
        {
            Assert(cbString <= cbBuffer);

            const int* lpDxSrc = lpDx;
            int* lpDxDst = lpDxT;
            int cb = cbString;
            const WCHAR* pch = pchString;

            memset(lpDxT, 0, cbString*sizeof(int));

            while(cb)
            {
                int cbChar = WideCharToMultiByte(
                    uiCodePage, 0,
                    pch, 1, NULL, NULL,
                    NULL, NULL);

                Assert(cbChar && cb>=cbChar);

                *lpDxDst = *lpDxSrc++;
                lpDxDst += cbChar;
                cb -= cbChar;
                pch++;
            }
        }

        fRetVal = ExtTextOutA(hdc, x, y, fuOptions, lprc, pbString, cbString, lpDxT);
    }
    else if(CVM_FEONNONFE == cm)
    {
        fRetVal = LSReExtTextOutW(pccs, hdc, x, y, fuOptions, lprc,
            pchString, cchString, lpDx);
    }
    else
    {
        fRetVal = ::ExtTextOutW(hdc, x, y, fuOptions, lprc,
            pchString, cchString, (int*)lpDx);
    }

    return fRetVal;
}

/*
*  LSUniscribeTextOut
*
*  @mfunc
*      Shapes, places and prints complex text through Uniscribe (USP10.DLL);
*
*  @rdesc
*      Returns indication that we successfully printed through Uniscribe
*/
HRESULT LSUniscribeTextOut(
                           HDC hdc,
                           int iX,
                           int iY,
                           UINT uOptions,
                           CONST RECT* prc,
                           LPCTSTR pString,
                           UINT cch,
                           int* piDx)
{
    HRESULT hr = S_OK;
    SCRIPT_STRING_ANALYSIS ssa;
    DWORD dwFlags;
    int xWidth = 0;
    const SIZE* pSize = NULL;
    const int* pich = NULL;
    const SCRIPT_LOGATTR* pSLA = NULL;

    if(piDx)
    {
        *piDx = 0;
    }

    Assert(!FALSE/*g_fExtTextOutGlyphCrash wlw note*/);

    dwFlags = SSA_GLYPHS | SSA_FALLBACK | SSA_BREAK;
    if(uOptions & ETO_RTLREADING)
    {
        dwFlags |= SSA_RTL;
    }

    // BUG 29234 - prc is sometimes passed in NULL
    if(prc != NULL)
    {
        xWidth = prc->right - prc->left;
        dwFlags |= SSA_CLIP;
    }

    while(cch)
    {
        // Using a 0 as the fourth parameter will default the max glyph count at cch * 1.5
        hr = ScriptStringAnalyse(hdc, pString, cch, 0, -1, dwFlags, xWidth,
            NULL, NULL, NULL, NULL, NULL, &ssa);

        if(FAILED(hr))
        {
            goto done;
        }

        pSize = ScriptString_pSize(ssa);

        if(piDx)
        {
            if(pSize->cx > *piDx)
            {
                *piDx = pSize->cx;
            }
        }

        pich = ScriptString_pcOutChars(ssa);

        if((UINT)*pich < cch)
        {
            pSLA = ScriptString_pLogAttr(ssa);

            int i = *pich;

            if(pSLA)
            {
                pSLA += *pich;
                while(i>0 && !(pSLA->fSoftBreak) && !(pSLA->fWhiteSpace))
                {
                    pSLA--;
                    i--;
                }

                if(i != *pich)
                {
                    ScriptStringFree(&ssa);

                    // Using a 0 as the fourth parameter will default the max glyph count at cch * 1.5
                    hr = ScriptStringAnalyse(hdc, pString, i, 0, -1, dwFlags, xWidth,
                        NULL, NULL, NULL, NULL, NULL, &ssa);

                    if(FAILED(hr))
                    {
                        goto done;
                    }

                    pich = ScriptString_pcOutChars(ssa);
                    Assert(*pich == i);
                }
            }
        }

        hr = ScriptStringOut(ssa, iX, iY, uOptions, prc, 0, 0, FALSE);

        iY += pSize->cy;

        if(*pich>0 && (UINT)*pich<cch)
        {
            cch -= *pich;
            pString += *pich;
        }
        else
        {
            cch = 0;
        }

        ScriptStringFree(&ssa);
    }

done:
    return hr;
}

/*
*  LSIsEnhancedMetafileDC( hDC )
*
*  @mfunc
*      Check if hDC is a Enhanced Metafile DC.
*      There is work around the Win95 FE ::GetObjectType() bug.
*
*  @rdesc
*      Returns TRUE for EMF DC.
*/
BOOL LSIsEnhancedMetafileDC(HDC hDC) //@parm handle to device context
{
    BOOL fEMFDC = FALSE;

    DWORD dwObjectType;

    dwObjectType = ::GetObjectType(hDC);

    if(OBJ_ENHMETADC==dwObjectType || OBJ_ENHMETAFILE==dwObjectType)
    {
        fEMFDC = TRUE;
    }
    else if(OBJ_DC == dwObjectType)
    {
        // HACK Alert,  Enhanced Metafile DC does not support any Escape function
        // and shoudl return 0.
        int iEscapeFuction = QUERYESCSUPPORT;
        fEMFDC = Escape(hDC, QUERYESCSUPPORT, sizeof(int),
            (LPCSTR)&iEscapeFuction, NULL) == 0;
    }

    return fEMFDC;
}

//-----------------------------------------------------------------------------
//
//  Function:   CLSRenderer::CLSRenderer
//
//  Synopsis:   Constructor for the LSRenderer
//
//-----------------------------------------------------------------------------
CLSRenderer::CLSRenderer(const CDisplay* const pdp, CFormDrawInfo* pDI) :
CLSMeasurer(pdp, pDI, FALSE)
{
    Assert(pdp);
    Init(pDI);
    _pci->_hdc = _hdc = pDI->GetDC();
    _hfontOrig = (HFONT)::GetCurrentObject(_hdc, OBJ_FONT);

    CFlowLayout* pFLayout = pdp->GetFlowLayout();
    CMarkup* pMarkup = pdp->GetMarkup();
    pMarkup->GetSelectionChunksForLayout(pFLayout, &_aryHighlight, & _cpSelMin, &_cpSelMax);
}

//-----------------------------------------------------------------------------
//
//  Function:   CLSRenderer::CLSRenderer
//
//  Synopsis:   Alternative constructor for the LSRenderer
//
//-----------------------------------------------------------------------------
CLSRenderer::CLSRenderer (const CDisplay* const pdp, LONG cp, CFormDrawInfo* pDI) :
CLSMeasurer(pdp, cp, pDI)
{
    Assert(pdp);
    Init(pDI);
    _pci->_hdc = _hdc = pDI->GetDC();
    _hfontOrig = (HFONT)::GetCurrentObject(_hdc, OBJ_FONT);

    CFlowLayout* pFLayout = pdp->GetFlowLayout();
    CMarkup* pMarkup = pdp->GetMarkup();
    pMarkup->GetSelectionChunksForLayout(pFLayout, &_aryHighlight, & _cpSelMin, & _cpSelMax);
}

//+====================================================================================
//
// Method: Destructor for CLSRenderer
//
// Synopsis: Remove any Highlight Segments which may have been allocated.
//
//------------------------------------------------------------------------------------
CLSRenderer::~CLSRenderer()
{
    // restore original font
    SelectFontEx(_hdc, _hfontOrig);

    HighlightSegment** ppSegment;
    int i;

    for(i=_aryHighlight.Size(),ppSegment=_aryHighlight;
        i>0;
        i--,ppSegment++)
    {
        delete(*ppSegment);
    }
}

/*
*  CLSRenderer::Init
*
*  @mfunc
*      initialize everything to zero
*/
void CLSRenderer::Init(CFormDrawInfo* pDI)
{
    Assert(pDI);

    _pDI = pDI;
    _rcView = _rcRender = _rcClip = _afxGlobalData._Zero.rc;
    _dwFlags() = 0;
    _ptCur.x = 0;
    _ptCur.y = 0;
    _fRenderSelection = TRUE;
}

/*
*  CLSRenderer::StartRender (&rcView, &rcRender, yHeightBitmap)
*
*  @mfunc
*      Prepare this renderer for rendering operations
*
*  @rdesc
*      FALSE if nothing to render, TRUE otherwise
*/
BOOL CLSRenderer::StartRender(
                               const RECT& rcView,         //@parm View rectangle
                               const RECT& rcRender)       //@parm Rectangle to render
{
    if(!_hdc)
    {
        _hdc = _pDI->GetDC();
    }

    AssertSz(_hdc, "CLSRenderer::StartRender() - No rendering DC");

    // Set view and rendering rects
    _rcView = rcView;
    _rcRender = rcRender;

    // Set background mode
    SetBkMode(_hdc, TRANSPARENT);

    _lastTextOutBy = DB_NONE;

    // If this is not the main display or it is a metafile
    // we want to ignore the logic to render selections
    if(_pDI->IsMetafile())
    {
        _fRenderSelection = FALSE;
    }

    // For hack around ExtTextOutW Win95FE Font and EMF problems.
    _fEnhancedMetafileDC = FALSE;

    return TRUE;
}

/*
*  CLSRenderer::NewLine (&li)
*
*  @mfunc
*      Init this CLSRenderer for rendering the specified line
*/
void CLSRenderer::NewLine(const CLine& li)
{
    _li = li;

    Assert(GetCp()+_li._cch <= GetLastCp());

    // BUGBUG: RTL FIX (paulnel) Are we safe to use _rcView.right here?
    //  for LTR display
    //  |--------------------------------Container width---------------------|
    //  |------------display view width------------------|                   |
    //  |  X-------------- LTR wraping line ---------->  |                   |
    //  |  X-------------- LTR overflowing line ---------|-------------->    |
    //  |  <-------------- RTL wrapping line ---------X  |                   |
    //  |  <-------------- RTL overflowing line ---------|-----------------X |
    if(!_pdp->IsRTL())
    {
        if(!_li._fRTL)
        {
            _ptCur.x = _rcView.left + _li._xLeftMargin + _li._xLeft;
        }
        else
        {
            _ptCur.x = _rcView.left + _li._xLeftMargin + _li._xLeft + _li._xWidth - 1;
        }
    }
    //  for RTL display the Origin is top, right. Use negative measurements
    //  |--------------------------------Container width---------------------|
    //  |                   |------------display view width------------------|
    //  |                   |  <-------------- RTL wrapping line ---------X  |
    //  |  <-------------- RTL overflowing line --------------------------X  |
    //  |                   |  X-------------- LTR wraping line ---------->  |
    //  |    X--------------|----------------- LTR overflowing line ------>  |
    else
    {
        if(_li._fRTL)
        {
            _ptCur.x = _rcView.right- _li._xRight - _li._xRightMargin - 1;
        }
        else
        {
            _ptCur.x = _rcView.right- _li._xRight - _li._xRightMargin - _li._xWidth;
        }
    }

    if(_li._fRelative) // if the current line is relative
    {
        if(!_pdp->IsRTL())
        {
            _ptCur.x += _xRelOffset;
        }
        else
        {
            _ptCur.x -= _xRelOffset;
        }
    }

    _cchWhiteAtBeginningOfLine = 0;
    _xAccumulatedWidth = 0;
}

//-----------------------------------------------------------------------------
//
//  Function:   RenderLine
//
//  Synopsis:   Renders one line of text
//
//  Params:     [li]: The line to be drawn
//
//  Returns:    void
//
//-----------------------------------------------------------------------------
VOID CLSRenderer::RenderLine(CLine& li, long xRelOffset)
{
    CMarginInfo marginInfo;
    POINT ptLS;
    LONG cpStartContainerLine;
    LONG xWidthContainerLine;
    LONG cchTotal;
    CTreePos* ptpNext = NULL;
    CTreePos* ptp;
    HRESULT hr;

    _xRelOffset = xRelOffset;

    NewLine(li);

    if(_li._fHidden)
    {
        goto Cleanup;
    }

    _rcClip = *_pDI->ClipRect();

    if(_li._fHasBulletOrNum)
    {
        CMarginInfo marginInfo;
        LONG cp = GetCp();
        CTreePos* ptp = GetPtp();
        LONG cchAdvance = 0;

        hr = StartUpLSDLL(_pLS, _pdp->GetMarkup());
        if(hr)
        {
            goto Cleanup;
        }

        AccountForRelativeLines(
            li,
            &cpStartContainerLine,
            &xWidthContainerLine,
            &_cpStartRender,
            &_cpStopRender,
            &cchTotal);

        _pdp->FormattingNodeForLine(cp, ptp, cchTotal-1, &cchAdvance, &ptp, NULL);
        cp += cchAdvance;

        _pLS->_treeInfo._fInited = FALSE;

        // Subsequently we will also reset the renderer inside renderbulletchar
        // with the correct pCFLI. Do it for now so that Setup works.
        _pLS->SetRenderer(this, NULL);
        if(S_OK==_pLS->Setup(0, cp, ptp, &marginInfo, NULL, FALSE))
        {
            COneRun* por = _pLS->_listFree.GetFreeOneRun(NULL);
            if(por)
            {
                por->_lscpBase = GetCp();
                por->_pPF = _pLS->_treeInfo._pPF;
                por->_fInnerPF = _pLS->_treeInfo._fInnerPF;
                por->_pCF = (CCharFormat*)_pLS->_treeInfo._pCF;
                por->_fInnerCF = _pLS->_treeInfo._fInnerCF;
                por->_fCharsForNestedElement  = _pLS->_treeInfo._fHasNestedElement;
                por->_fCharsForNestedLayout   = _pLS->_treeInfo._fHasNestedLayout;
                por->_fCharsForNestedRunOwner = _pLS->_treeInfo._fHasNestedRunOwner;
                por->_ptp = _pLS->_treeInfo._ptpFrontier;
                por->_sid = sidAsciiLatin;

                por = _pLS->AttachOneRunToCurrentList(por);
                if(por)
                {
                    if(_lastTextOutBy != DB_LINESERV)
                    {
                        _lastTextOutBy = DB_LINESERV;
                        SetTextAlign(_hdc, TA_TOP|TA_LEFT);
                    }

                    RenderStartLine(por);
                }
            }
        }

        // reset so that the next person to call setup on CLineServices
        // will cleanup after us.
        _pLS->_treeInfo._fInited = FALSE;
    }

    if(_li.IsFrame())
    {
        goto Cleanup;
    }

    if(_li._cch == 0)
    {
        goto Cleanup;
    }

    // Set the renderer before we do any output.
    _pLS->SetRenderer(this, _pdp->GetWrapLongLines());

    // If the line can be blasted to the screen and it was successfully blasted
    // then we are done.
    if(_li._fCanBlastToScreen)
    {
        // BUGBUG(SujalP): This is added for debugging a stress crash. Should remove
        // before shipping.
        static CLSRenderer* g_pRendererDbg;

        g_pRendererDbg = this;

        ptpNext = BlastLineToScreen(li);
        if(ptpNext)
        {
            goto Done;
        }
    }

    hr = StartUpLSDLL(_pLS, _pdp->GetMarkup());
    if(hr)
    {
        goto Cleanup;
    }

    if(_lastTextOutBy != DB_LINESERV)
    {
        _lastTextOutBy = DB_LINESERV;
        SetTextAlign(_hdc, TA_TOP|TA_LEFT);
    }

    AccountForRelativeLines(
        li,
        &cpStartContainerLine,
        &xWidthContainerLine,
        &_cpStartRender,
        &_cpStopRender,
        &cchTotal);

    _li._cch += _cp - cpStartContainerLine;

    if(cpStartContainerLine != GetCp())
    {
        SetCp(cpStartContainerLine, NULL);
    }

    _pdp->FormattingNodeForLine(GetCp(), GetPtp(), _li._cch, &_cchPreChars, &ptp, &_fMeasureFromTheStart);
    if(!_fMeasureFromTheStart)
    {
        cpStartContainerLine += _cchPreChars;
        _li._cch -= _cchPreChars;
        SetPtp(ptp, cpStartContainerLine);
    }

    InitForMeasure(MEASURE_BREAKATWORD);
    LSDoCreateLine(
        cpStartContainerLine, NULL,
        &marginInfo,
        xWidthContainerLine,
        &_li,
        FALSE, NULL);
    if(!_pLS->_plsline)
    {
        goto Cleanup;
    }

    ptLS = _ptCur;

    LsDisplayLine(
        _pLS->_plsline,   // The line to be drawn
        &ptLS,            // The point at which to draw the line
        1,                // Draw in transparent mode
        &_rcClip);        // The clipping rectangle
    ptpNext = _pLS->FigureNextPtp(_cp+_li._cch);

Cleanup:
    _pLS->DiscardLine();

Done:
    // increment y position to next line
    if(_li._fForceNewLine)
    {
        _ptCur.y += _li._yHeight;
    }

    // Go past the contents of this line
    Advance(_li._cch, ptpNext);

    return;
}

//-----------------------------------------------------------------------------
//
//  Function:   BlastLineToScreen
//
//  Synopsis:   Renders one line in one shot w/o remeasuring if it can be done.
//
//  Params:     li: The line to be rendered
//
//  Returns:    Whether it rendered anything at all
//
//-----------------------------------------------------------------------------
CTreePos* CLSRenderer::BlastLineToScreen(CLine& li)
{
    CTreePos* ptpRet = NULL;

    Assert(_li._fCanBlastToScreen);
    Assert(!_li._fHidden);

    // For now, lets not do relative lines...
    if(_li._fPartOfRelChunk)
    {
        goto Cleanup;
    }

    // We need to render all the characters in the line
    _cpStartRender = GetCp();
    _cpStopRender  = _cpStartRender + _li._cch;

    // marka - examine Cached _cpSelMin, _cpSelMax for when we render.
    if(_cpSelMax != -1)
    {
        // Selection starts outside the line and ends
        // inside or beyond the line, then use old method to render
        if(_cpSelMin<=_cpStartRender && _cpSelMax>_cpStartRender)
        {
            goto Cleanup;
        }

        // Selection starts inside the line. Irrespective
        // of where it ends, use the old method
        if(_cpSelMin>=_cpStartRender && _cpSelMin<_cpStopRender)
        {
            goto Cleanup;
        }
    }

    // Now we are pretty sure that we can blast this line, so lets go do it.
    {
        CTreePos* ptp;
        LONG      cp  = GetCp();
        LONG      cch = _li._cch;
        LONG      cpAtPtp;

        LONG      cchToRender;
        LONG      cchSkip;
        CTreeNode* pNode = _pdp->FormattingNodeForLine(cp, GetPtp(), cch, &cchSkip, &ptp, NULL);
        COneRun   onerun;

        const TCHAR* pch = NULL;
        CTxtPtr   tp(_pdp->GetMarkup(), cp+=cchSkip);
        LONG      cchValid = 0;

        LONG      xOld = -1;
        LONG      yOld = LONG_MIN;
        LONG      xCur = _ptCur.x;
        LONG      yCur = _ptCur.y;
        LONG      yOriginal = _ptCur.y;
        BOOL      fRenderedText = FALSE;
        LONG      cchCharsTrimmed = 0;
        BOOL      fHasInclEOLWhite = pNode->GetParaFormat()->HasInclEOLWhite(SameScope(pNode, _pFlowLayout->ElementContent()));

        DEBUG_ONLY(BOOL fNoMoreTextOut = FALSE);

        if(_lastTextOutBy != DB_BLAST)
        {
            _lastTextOutBy = DB_BLAST;
            SetTextAlign(_hdc, TA_TOP|TA_LEFT|TA_UPDATECP);
        }

        cch -= cchSkip;
        cpAtPtp = ptp->GetCp();
        while(cch > 0)
        {
            if(cchValid == 0)
            {
                Assert(cp == (long)tp.GetCp());

                pch = tp.GetPch(cchValid);
                Assert(pch != NULL);
                if(pch == NULL)
                {
                    goto Cleanup;
                }
                cchValid = min(cchValid, cch);
                tp.AdvanceCp(cchValid);
            }

            cchToRender = 0;
            if(ptp->IsPointer())
            {
                ptp = ptp->NextTreePos();
                continue;
            }
            if(ptp->IsNode())
            {
                pNode = ptp->Branch();

                // Start off with a default number of chars to render
                cchToRender = 1;
                Assert(cchToRender == ptp->GetCch());

                if(ptp->IsBeginElementScope())
                {
                    const CCharFormat* pCF = pNode->GetCharFormat();
                    CElement* pElement = pNode->Element();

                    if(pCF->IsDisplayNone())
                    {
                        cchToRender = GetNestedElementCch(pElement, &ptp);
                    }
                    else if(pNode->NeedsLayout())
                    {
                        CLayout* pLayout = pNode->GetUpdatedLayoutPtr();

                        Assert(pLayout != _pFlowLayout);
                        if(pElement->IsInlinedElement())
                        {
                            LONG xWidth;

                            _pFlowLayout->GetSiteWidth(pLayout, _pci, FALSE, 0, &xWidth);
                            xCur += xWidth;
                            MoveToEx(_pci->_hdc, xCur, yCur, NULL);
                            fRenderedText = TRUE;
                        }
                        cchToRender = _pLS->GetNestedElementCch(pNode->Element(), &ptp);

                        // We either have overlapping layouts (which is when the cchToRender
                        // etc test will succeed), or our last ptp has to be that of the
                        // layout we just rendered.
                        Assert(cchToRender!=(pNode->Element()->GetElementCch()+2)
                            || pNode->Element()==ptp->Branch()->Element());
                    }
                }

                if(ptp->IsEndNode())
                {
                    pNode = pNode->Parent();
                }
                ptp = ptp->NextTreePos();
                cpAtPtp += cchToRender;
            }
            else
            {
                Assert(ptp->IsText());
                Assert(pNode);

                if(ptp->Cch() == 0)
                {
                    ptp = ptp->NextTreePos();
                    continue;
                }

                LONG cchRemainingInTextPos = ptp->Cch() - (cp - cpAtPtp);
                LONG cchCanRenderNow = min(cchRemainingInTextPos, cchValid);
                BOOL fWhiteSpaceSkip = FALSE;
                CTreePos* ptpThis = ptp;

                if(fRenderedText)
                {
                    cchToRender = cchCanRenderNow;
                }
                else
                {
                    LONG i = 0;
                    if(!fHasInclEOLWhite)
                    {
                        for(i=0; i<cchCanRenderNow; i++)
                        {
                            if(!IsWhite(pch[i]))
                            {
                                break;
                            }
                            fWhiteSpaceSkip = TRUE;
                        }
                    }
                    if(fWhiteSpaceSkip)
                    {
                        cchToRender = i;
                    }
                    else
                    {
                        cchToRender = cchCanRenderNow;
                    }
                }

                cchRemainingInTextPos -= cchToRender;
                if(cchRemainingInTextPos == 0)
                {
                    cpAtPtp += ptp->Cch();
                    ptp = ptp->NextTreePos();
                }

                if(!fWhiteSpaceSkip)
                {
                    POINT ptTemp;
                    CCcs* pccs;
                    CBaseCcs* pBaseCcs;
                    BOOL fUnderlined;

                    fRenderedText = TRUE;

                    // Is this needed?
                    memset(&onerun, 0, sizeof(onerun));

                    onerun._fInnerCF = SameScope(pNode, _pFlowLayout->ElementOwner());
                    onerun._pCF = (CCharFormat*)pNode->GetCharFormat();
                    onerun._bConvertMode = CVM_UNINITED;
                    onerun._ptp = ptpThis;
                    onerun.SetSidFromTreePos(ptpThis);
                    onerun._lscpBase = cp;  // For CLineServices::UnUnifyHan
                    Assert(onerun._pCF);

                    fUnderlined = onerun._pCF->_fStrikeOut
                        || onerun._pCF->_fOverline
                        || onerun._pCF->_fUnderline;

                    if(fUnderlined)
                    {
                        cchCharsTrimmed = TrimTrailingSpaces(
                            cchToRender,       // number of chars being rendered now
                            cp+cchToRender,    // cp of next run to be blasted
                            ptp,
                            cch-cchToRender);  // chars remaining to be rendered
                        cchToRender -= cchCharsTrimmed;
                        Assert(cchToRender >= 0);
                        if(cchToRender == 0)
                        {
                            fUnderlined = FALSE;
                            cchToRender = cchCharsTrimmed;
                        }
                    }

                    pccs = _pLS->GetCcs(&onerun, _pDI->GetDC(), _pDI);
                    if(pccs == NULL)
                    {
                        goto Cleanup;
                    }

                    pBaseCcs = pccs->GetBaseCcs();

                    yCur = yOriginal + _li._yHeight - _li._yDescent +
                        pBaseCcs->_yDescent - pBaseCcs->_yHeight - pBaseCcs->_yOffset;
                    if(yCur != yOld)
                    {
                        yOld = yCur;
#ifdef _DEBUG
                        BOOL fSuccess =
#endif
                            MoveToEx(_pci->_hdc, xCur, yCur, NULL);
#ifdef _DEBUG
                        AssertSz(fSuccess, "Failed to do moveto, bad HDC?");
                        if(!fSuccess)
                        {
                            DWORD winerror;
                            winerror = ::GetLastError();
                        }
#endif
                    }
                    xOld = xCur;

                    DEBUG_ONLY(GetCurrentPositionEx(_pci->_hdc, &ptTemp));
                    Assert(ptTemp.x == xCur);
                    Assert(ptTemp.y == yCur);
                    Assert(fNoMoreTextOut == FALSE);

                    onerun._lscpBase = cp;
                    TextOut(
                        &onerun,                 // por
                        FALSE,                   // fStrikeOut
                        FALSE,                   // fUnderLine
                        NULL,                    // pptText
                        pch,                     // pch
                        NULL,                    // lpDx
                        cchToRender,             // cwchRun
                        lstflowES,               // ktFlow
                        0,                       // kDisp
                        (const POINT*)&_afxGlobalData._Zero, // pptRun
                        NULL,                    // heightPres
                        -1,                      // dupRun
                        0,                       // dupLineUnderline
                        NULL);                   // pRectClip

                    // We will have moved to the right so get our new position
                    GetCurrentPositionEx(_pci->_hdc, &ptTemp);
                    xCur = ptTemp.x;

                    // Keep yCur as it is.
                    if(fUnderlined)
                    {
                        LSULINFO lsUlInfo;

                        if(lserrNone != _pLS->GetRunUnderlineInfo(
                            &onerun,    // PLSRUN
                            NULL,       // PCHEIGHTS
                            lstflowES,  // ktFlow
                            &lsUlInfo))
                        {
                            goto Cleanup;
                        }

                        ptTemp.x = xOld;
                        ptTemp.y = yOriginal;
                        DrawUnderline(
                            &onerun,                          // por
                            lsUlInfo.kulbase,                 // kUlBase
                            &ptTemp,                          // pptStart
                            xCur - xOld,                      // dupUl
                            lsUlInfo.dvpFirstUnderlineSize,   // dvpUl
                            lstflowES,                        // kTFlow
                            0,                                // kDisp
                            &_rcClip);                        // prcClip
                        MoveToEx(_pci->_hdc, xCur, yCur, NULL);
                        cchToRender += cchCharsTrimmed;
                        DEBUG_ONLY(fNoMoreTextOut = !!cchCharsTrimmed);
                    }
                }
            }
            cp += cchToRender;
            cch -= cchToRender;

            cchValid -= cchToRender;

            // cchValid can go to < 0 if the current text run finishes *within* a
            // site or hidden stuff. In both these cases we will have cchToRender
            // be > cchValid. Take care of this.
            if(cchValid < 0)
            {
                tp.AdvanceCp(-cchValid);
                cchValid = 0;
            }
            else
            {
                pch += cchToRender;
            }
        }
        memset(&onerun, 0, sizeof(onerun));
        ptpRet = ptp;
    }

Cleanup:
    return ptpRet;
}

//-----------------------------------------------------------------------------
//
//  Function:   TrimTrailingSpaces
//
//  Synopsis:   Returns the number of white chars to be trimmed (if necessary)
//              at the end of a run
//
//  Returns:    Number of characters to be trimmed
//
//-----------------------------------------------------------------------------
LONG CLSRenderer::TrimTrailingSpaces(
                                     LONG cchToRender,
                                     LONG cp,
                                     CTreePos* ptp,
                                     LONG cchRemainingInLine)
{
    LONG cchAdvance = 0;
    BOOL fTrim      = TRUE;
    const CCharFormat* pCF;
    LONG cchTrim;
    CTreeNode* pNode;
    CElement* pElement;
    LONG cpMostOfRunToTrim = cp;

    LONG junk;

    ptp = _pdp->GetMarkup()->TreePosAtCp(cp, &junk);
    while(cchRemainingInLine > 0)
    {
        if(ptp->GetCch() == 0)
        {
            cchAdvance = 0;
        }
        else
        {
            pNode = ptp->GetBranch();
            pCF = pNode->GetCharFormat();

            if(ptp->IsNode())
            {
                if(ptp->IsBeginNode())
                {
                    pElement = pNode->Element();

                    if(pCF->IsDisplayNone())
                    {
                        cchAdvance = GetNestedElementCch(pElement, &ptp);
                    }
                    else if(ptp->IsEdgeScope() && pElement->NeedsLayout())
                    {
                        if(pElement->IsInlinedElement())
                        {
                            fTrim = FALSE;
                            break;
                        }
                        else
                        {
                            cchAdvance = GetNestedElementCch(pElement, &ptp);
                        }
                    }
                }
            }
            else if(ptp->Cch() > 0)
            {
                fTrim = FALSE;
                break;
            }
        }

        cp += cchAdvance;
        cchRemainingInLine -= cchAdvance;
        ptp = ptp->NextTreePos();
        cchAdvance = ptp->GetCch();
    }

    if(fTrim)
    {
        CTxtPtr tp(_pdp->GetMarkup(), cpMostOfRunToTrim);

        cchTrim = 0;
        while(cchToRender && IsWhite(tp.GetPrevChar()))
        {
            cchToRender--;
            cchTrim++;
            tp.AdvanceCp(-1);
        }
    }
    else
    {
        cchTrim = 0;
    }

    return cchTrim;
}

//-----------------------------------------------------------------------------
//
//  Function:   TextOut
//
//  Synopsis:   Renders one run of text
//
//  Params:     Same as that for LineServices DrawTextRun callback
//
//  Returns:    Number of characters rendered
//
//  Dev Note:   Any changes made to this function should be reflected in
//              CLSRenderer::GlyphOut() (as appropriate)
//
//-----------------------------------------------------------------------------
LONG CLSRenderer::TextOut(
                     COneRun* por,           // IN
                     BOOL fStrikeout,        // IN
                     BOOL fUnderline,        // IN
                     const POINT* pptText,   // IN
                     LPCWSTR pch,            // IN was plwchRun
                     const int* lpDx,        // IN was rgDupRun
                     DWORD cwchRun,          // IN
                     LSTFLOW kTFlow,         // IN
                     UINT kDisp,             // IN
                     const POINT* pptRun,    // IN
                     PCHEIGHTS heightPres,   // IN
                     long dupRun,            // IN
                     long dupLimUnderline,   // IN
                     const RECT* pRectClip)  // IN
{
    DWORD    dwDCObjType = GetObjectType(_hdc);
    CONVERTMODE cm;
    TCHAR    ch;
    CString  strTransformedText;
    COLORREF crTextColor = 0;
    COLORREF crNewTextColor = 0;
    DWORD    cch = cwchRun;
    const    CCharFormat* pCF = por->GetCF();
    CCcs*    pccs = SetNewFont(por);
    CBaseCcs* pBaseCcs;
    RECT     rcClipSave = _rcClip;
    LONG     yDescent = _li._yDescent;
    CDataAry<int> aryDx;
    LONG     yCur;

    if(!pccs)
    {
        goto Cleanup;
    }

    if(ShouldSkipThisRun(por, dupRun))
    {
        goto Cleanup;
    }

    pBaseCcs = pccs->GetBaseCcs();

    cm = pBaseCcs->GetConvertMode(_fEnhancedMetafileDC, _pDI->IsMetafile());

    // This is where this run will be drawn.
    _ptCur.x = pptRun->x - ((!(kTFlow & fUDirection))?0:dupRun-1) - _xAccumulatedWidth;

    // If pptText is NULL, that means that we are blasting to the screen, so we
    // should just use _ptCur.y as our yCur
    AssertSz((pptText==NULL || _ptCur.y==pptText->y || por->GetCF()->_fIsRuby),
        "Our baseline differs from LS baseline!");
    yCur = (pptText==NULL) ? _ptCur.y : pptText->y;

    // Trim all nondisplayable linebreaking chars off end
    for(; cch ; cch--)
    {
        ch = pch[cch-1];

        if(!(ch==LF || ch==CR))
        {
            break;
        }
    }

    if(!cch)
    {
        goto Cleanup;
    }

    if(!GetMarkup()->Doc()->PaintBackground())
    {
        // If we are part of a print doc and the background is not printed,
        // we conceptually replace the background with white (because most
        // paper is white).  This means that we have to enhance the contrast
        // of bright-colored text on dark backgrounds.

        COLORREF crNewBkColor = RGB(255, 255, 255); // white background
        int nMinimalLuminanceDifference = 150; // for now

        crNewTextColor = GetTextColor(_hdc);

        // Sujal insists that if both colors are the same (in this case both
        // white), then the second color will be modified.
        ContrastColors(crNewTextColor, crNewBkColor, nMinimalLuminanceDifference);

        // Finally, turn off the higher order bits
        crNewTextColor &= CColorValue::MASK_COLOR;

        crTextColor = SetTextColor(_hdc, crNewTextColor);
    }

    AssertSz(pCF, "CLSRenderer::TextOut: We better have char format here");

    // In some fonts in some locales, NBSPs aren't rendered like spaces.
    // Under these circumstances, we need to convert NBSPs to spaces
    // Before calling LSReExtTextOut.
    if(pCF && _li._fHasNBSPs && pBaseCcs->ConvertNBSPs(_pDI->GetDC(), _pDI->_pDoc))
    {
        const TCHAR* pchStop;
        TCHAR* pch2;

        HRESULT hr = strTransformedText.Set(pch, cch);
        if(hr)
        {
            goto Cleanup;
        }
        pch = strTransformedText;

        pch2 = (TCHAR*)pch;
        pchStop = pch + cch;

        while(pch2 < pchStop)
        {
            if(*pch2 == WCH_NBSP)
            {
                *pch2 = L' ';
            }

            pch2++;
        }
    }

    // Reverse the run if our text is flowing from right to left.
    // BUGBUG (mikejoch) We actually want to be more conditional about doing
    // this for metafiles, as we don't want to reverse runs that would
    // otherwise be glyphed. Maybe by looking at
    // por->GetComplexRun()->GetAnalysis()?
    if(kTFlow & fUDirection)
    {
        TCHAR* pch1;
        TCHAR* pch2;
        const int* pDx1;
        int* pDx2;

        if(pch != strTransformedText)
        {
            HRESULT hr = strTransformedText.Set(pch, cch);
            if(hr)
            {
                goto Cleanup;
            }
            pch = strTransformedText;
        }

        if(FAILED(aryDx.Grow(cch)))
        {
            goto Cleanup;
        }

        pch1 = (TCHAR*) pch;
        pch2 = pch1 + cch - 1;
        while(pch1 < pch2)
        {
            TCHAR ch = *pch1;
            *pch1 = *pch2;
            *pch2 = ch;
            pch1++;
            pch2--;
        }

        pDx1 = lpDx + cch - 1;
        pDx2 = aryDx;
        while(pDx1 >= lpDx)
        {
            *pDx2++ = *pDx1--;
        }
        lpDx = aryDx;
    }

    if(_pLS->IsAdornment())
    {
        LONG yTextHeight = _li._yHeight - _li.GetYTop();

        // If the first line in the LI is a frame line, then its _yDescent
        // value is garbage. This also means that we will not center WRT
        // the text following the frame line. However, this is much better
        // than before when we did not render the bullet at all when the
        // first line in a LI was a frame line.(SujalP)
        if(_li.IsFrame())
        {
            yDescent = 0;
        }

        // Center the bullet in the ascent of the text line only if
        // the line of text is smaller than the bullet. This adjustment
        // keeps the bullets visible, but tends to keep them low on the
        // line for large text fonts.
        if(pBaseCcs->_yHeight-pBaseCcs->_yDescent > yTextHeight-yDescent)
        {
            yCur += ((pBaseCcs->_yHeight - pBaseCcs->_yDescent) -
                (yTextHeight - yDescent))/2;
            _ptCur.y = yCur;
        }
    }

    if(!_fBrowseMode || !pCF->IsVisibilityHidden())
    {
        int* lpDxNew = NULL;
        const int* lpDxPtr = lpDx;
        long lGridOffset = 0;

        if(pCF->HasCharGrid(TRUE))
        {
            Assert(cch > 0);
            lpDxNew = new int[cch];
            if(!lpDxNew)
            {
                goto Cleanup;
            }
            lpDxPtr = lpDxNew;

            switch(por->GetCF()->GetLayoutGridType(TRUE))
            {
            case styleLayoutGridTypeStrict:
            case styleLayoutGridTypeFixed:
                {
                    if(por->IsOneCharPerGridCell())
                    {
                        long durThisChar = 0;
                        long durNextChar = 0;
                        long lThisDoubleGridOffset = 0;
                        long lError = 0;

                        pccs->Include(pch[0], durThisChar);
                        lThisDoubleGridOffset = lpDx[0] - durThisChar;
                        lGridOffset = lThisDoubleGridOffset/2;

                        for(unsigned int i=0; i<cch-1; i++)
                        {
                            long lNextDoubleGridOffset = 0;

                            pccs->Include(pch[i+1], durNextChar);
                            lNextDoubleGridOffset = lpDx[i+1] - durNextChar;

                            // The width of the this char is essentially it's width
                            // plus the remaining space ahead of it in its grid cell,
                            // plus the space between the beginning of the next grid
                            // cell and the next character.
                            lpDxNew[i] = lpDx[i] - (lThisDoubleGridOffset - lNextDoubleGridOffset + lError)/2;
                            lError = (lThisDoubleGridOffset - lNextDoubleGridOffset + lError)%2;

                            lThisDoubleGridOffset = lNextDoubleGridOffset;
                            durThisChar = durNextChar;
                        }
                        lpDxNew[cch-1] = lpDx[cch-1] - (lThisDoubleGridOffset + lError)/2;
                    }
                    else
                    {
                        // Get next run in a chain
                        COneRun* pNextRun = por->_pNext;
                        while(pNextRun && !pNextRun->_ptp->IsText())
                        {
                            pNextRun = pNextRun->_pNext;
                        }
                        if(!por->IsSameGridClass(pNextRun))
                        {
                            pNextRun = 0;
                        }
                        if(!pNextRun)
                        {
                            // we are at the end of a runs chain
                            LONG duCumulative = 0;
                            for(unsigned long i=0; i<cch; i++)
                            {
                                duCumulative += lpDx[i];
                            }
                            lGridOffset = por->_xDrawOffset;
                            lpDxNew[cch-1] = lpDx[cch-1] + por->_xWidth - duCumulative;
                            memcpy(lpDxNew, lpDx, (cch-1)*sizeof(int));
                        }
                        else
                        {   // we are at the begining or inside a runs chain
                            lGridOffset = por->_xDrawOffset;
                            lpDxPtr = lpDx;
                        }
                    }
                }
                break;

            case styleLayoutGridTypeLoose:
            default:
                // In loose grid mode (default mode) the width of a char is 
                // width of its grid cell.
                lpDxPtr = lpDx;
            }
        }

        Assert(GetBkMode(_hdc) == TRANSPARENT);

        yCur += _li._yHeight - yDescent +
            pBaseCcs->_yDescent - pBaseCcs->_yHeight - pBaseCcs->_yOffset;

        if(_fDisabled)
        {
            if(_crForeDisabled != _crShadowDisabled)
            {
                //draw the shadow
                SetTextColor(_hdc, _crShadowDisabled);
                LSReExtTextOut(
                    pccs,
                    _hdc,
                    _ptCur.x+1+lGridOffset,
                    yCur+1,
                    ETO_CLIPPED,
                    &_rcClip,
                    pch,
                    cch,
                    lpDxPtr,
                    cm);
            }
            SetTextColor(_hdc, _crForeDisabled);
        }

        LSReExtTextOut(
            pccs,
            _hdc,
            _ptCur.x+lGridOffset,
            yCur,
            ETO_CLIPPED,
            &_rcClip,
            pch,
            cch,
            lpDxPtr,
            cm);

        if(por->HasBgColor())
        {
            RECT rcSelect;

            AdjustColors(por);

            // Determine the bounds of the selection.
            rcSelect.left = _ptCur.x;
            rcSelect.right = rcSelect.left + dupRun;
            rcSelect.top = yCur;
            rcSelect.bottom = rcSelect.top + pBaseCcs->_yHeight;

            // If current run is selected, fix up _rcClip.
            if(por->IsSelected())
            {
                FixupClipRectForSelection();
            }

            // Clip the rect to the view and current clipping rect.
            IntersectRect(&rcSelect, &rcSelect, &_rcClip);

            if(_fDisabled)
            {
                if(_crForeDisabled != _crShadowDisabled)
                {
                    //draw the shadow
                    SetTextColor(_hdc, _crShadowDisabled);
                    LSReExtTextOut(
                        pccs,
                        _hdc,
                        _ptCur.x+1+lGridOffset,
                        yCur+1,
                        ETO_CLIPPED|ETO_OPAQUE,
                        &rcSelect,
                        pch,
                        cch,
                        lpDxPtr,
                        cm);
                }
                SetTextColor(_hdc, _crForeDisabled);
            }

            LSReExtTextOut(
                pccs,
                _hdc,
                _ptCur.x+lGridOffset,
                yCur,                       
                ETO_CLIPPED|ETO_OPAQUE,
                &rcSelect,
                pch,
                cch,
                lpDxPtr,
                cm);
        }

        if(lpDxNew)
        {
            delete lpDxNew;
        }
    }

    _ptCur.x = ((!(kTFlow&fUDirection)) ? _rcClip.right : _rcClip.left);

Cleanup:
    _rcClip = rcClipSave;
    return cch;
}

//-----------------------------------------------------------------------------
//
//  Function:   GlyphOut
//
//  Synopsis:   Renders one run of glyphs
//
//  Params:     Same as that for LineServices DrawGlyphs callback
//
//  Returns:    Number of glyphs rendered
//
//  Dev Note:   Any changes made to this function should be reflected in
//              CLSRenderer::TextOut() (as appropriate)
//
//-----------------------------------------------------------------------------
LONG CLSRenderer::GlyphOut(
                      COneRun* por,           // IN was plsrun
                      BOOL fStrikeout,        // IN
                      BOOL fUnderline,        // IN
                      PCGINDEX pglyph,        // IN
                      const int* pDx,         // IN was rgDu
                      const int* pDxAdvance,  // IN was rgDuBeforeJust
                      PGOFFSET pOffset,       // IN was rgGoffset
                      PGPROP pVisAttr,        // IN was rgGProp
                      PCEXPTYPE pExpType,     // IN was rgExpType
                      DWORD cglyph,           // IN
                      LSTFLOW kTFlow,         // IN
                      UINT kDisp,             // IN
                      const POINT* pptRun,    // IN
                      PCHEIGHTS heightsPres,  // IN
                      long dupRun,            // IN
                      long dupLimUnderline,   // IN
                      const RECT* pRectClip)  // IN
{
    COLORREF            crTextColor = 0;
    COLORREF            crNewTextColor = 0;
    const CCharFormat*  pCF = por->GetCF();
    CComplexRun*        pcr = por->GetComplexRun();
    CCcs*               pccs = SetNewFont(por);
    CBaseCcs*           pBaseCcs;
    RECT                rcClipSave = _rcClip;
    LONG                yDescent = _li._yDescent;
    SCRIPT_CACHE*       psc;
    LONG yCur;

    // LS should not be calling this function for metafiles.
    Assert(!FALSE/*g_fExtTextOutGlyphCrash wlw note*/);
    Assert(!_fEnhancedMetafileDC && !_pDI->IsMetafile());

    if(!pccs)
    {
        goto Cleanup;
    }

    if(ShouldSkipThisRun(por, dupRun))
    {
        goto Cleanup;
    }

    pBaseCcs = pccs->GetBaseCcs();
    psc = pBaseCcs->GetUniscribeCache();

    // This is where this run will be drawn.
    _ptCur.x = pptRun->x - ((kTFlow&fUDirection)?dupRun-1:0) - _xAccumulatedWidth;

    AssertSz((_ptCur.y==pptRun->y || por->GetCF()->_fIsRuby),
        "Our baseline differs from LS baseline!");
    yCur = pptRun->y;

    if(!GetMarkup()->Doc()->PaintBackground())
    {
        // If we are part of a print doc and the background is not printed,
        // we conceptually replace the background with white (because most
        // paper is white).  This means that we have to enhance the contrast
        // of bright-colored text on dark backgrounds.
        COLORREF crNewBkColor = RGB(255, 255, 255); // white background
        int nMinimalLuminanceDifference = 150; // for now

        crNewTextColor = GetTextColor(_hdc);

        // Sujal insists that if both colors are the same (in this case both
        // white), then the second color will be modified.
        ContrastColors(crNewTextColor, crNewBkColor, nMinimalLuminanceDifference);

        // Finally, turn off the higher order bits
        crNewTextColor &= CColorValue::MASK_COLOR;

        crTextColor = SetTextColor(_hdc, crNewTextColor);
    }

    AssertSz(pCF, "CLSRenderer::GlyphOut: We better have char format here");
    AssertSz(pcr, "CLSRenderer::GlyphOut: We better have CComplexRun here");

    if(_pLS->IsAdornment())
    {
        LONG yTextHeight = _li._yHeight - _li.GetYTop();

        // If the first line in the LI is a frame line, then its _yDescent
        // value is garbage. This also means that we will not center WRT
        // the text following the frame line. However, this is much better
        // than before when we did not render the bullet at all when the
        // first line in a LI was a frame line.(SujalP)
        if(_li.IsFrame())
        {
            yDescent = 0;
        }

        // Center the bullet in the ascent of the text line only if
        // the line of text is smaller than the bullet. This adjustment
        // keeps the bullets visible, but tends to keep them low on the
        // line for large text fonts.
        if(pBaseCcs->_yHeight-pBaseCcs->_yDescent > yTextHeight-yDescent)
        {
            yCur += ((pBaseCcs->_yHeight - pBaseCcs->_yDescent) -
                (yTextHeight-yDescent))/2;
            _ptCur.y = yCur;
        }
    }

    if(!_fBrowseMode || !pCF->IsVisibilityHidden())
    {
        Assert(GetBkMode(_hdc) == TRANSPARENT);

        yCur += _li._yHeight - yDescent +
            pBaseCcs->_yDescent - pBaseCcs->_yHeight - pBaseCcs->_yOffset;

        if(_fDisabled)
        {
            if(_crForeDisabled != _crShadowDisabled)
            {
                //draw the shadow
                SetTextColor(_hdc, _crShadowDisabled);
                ScriptTextOut(
                    _hdc,
                    psc,
                    _ptCur.x+1,
                    yCur+1,
                    ETO_CLIPPED,
                    &_rcClip,
                    pcr->GetAnalysis(),
                    NULL,
                    0,
                    pglyph,
                    cglyph,
                    pDxAdvance,
                    pDx,
                    pOffset);
            }
            SetTextColor(_hdc, _crForeDisabled);
        }

        ScriptTextOut(
            _hdc,
            psc,
            _ptCur.x,
            yCur,
            ETO_CLIPPED,
            &_rcClip,
            pcr->GetAnalysis(),
            NULL,
            0,
            pglyph,
            cglyph,
            pDxAdvance,
            pDx,
            pOffset);

        // If the run is selected, then we need to set up selection colors.
        if(por->HasBgColor())
        {
            RECT rcSelect;

            AdjustColors(por);

            // Determine the bounds of the selection.
            rcSelect.left = _ptCur.x;
            rcSelect.right = rcSelect.left + dupRun;
            rcSelect.top = yCur;
            rcSelect.bottom = rcSelect.top + pBaseCcs->_yHeight;

            // If current run is selected, fix up _rcClip.
            if(por->IsSelected())
            {
                FixupClipRectForSelection();
            }

            // Clip the rect to the view and current clipping rect.
            IntersectRect(&rcSelect, &rcSelect, &_rcClip);

            if(_fDisabled)
            {
                if(_crForeDisabled != _crShadowDisabled)
                {
                    //draw the shadow
                    SetTextColor(_hdc, _crShadowDisabled);
                    ScriptTextOut(
                        _hdc,
                        psc,
                        _ptCur.x+1,
                        yCur+1,
                        ETO_CLIPPED|ETO_OPAQUE,
                        &rcSelect,
                        pcr->GetAnalysis(),
                        NULL,
                        0,
                        pglyph,
                        cglyph,
                        pDxAdvance,
                        pDx,
                        pOffset);
                }
                SetTextColor(_hdc, _crForeDisabled);
            }

            ScriptTextOut(
                _hdc,
                psc,
                _ptCur.x,
                yCur,
                ETO_CLIPPED|ETO_OPAQUE,
                &rcSelect,
                pcr->GetAnalysis(),
                NULL,
                0,
                pglyph,
                cglyph,
                pDxAdvance,
                pDx,
                pOffset);
        }
    }

    _ptCur.x = ((kTFlow&fUDirection) ? _rcClip.left : _rcClip.right);

Cleanup:
    _rcClip = rcClipSave;
    return cglyph;
}

//+----------------------------------------------------------------------------
//
// Member:      CLSRenderer::FixupClipRectForSelection()
//
// Synopsis:    This function shrinks the clip rect when we have a selection
//              in order to avoid rendering selection out of bounds.
//              ONLY TO BE CALLED WHEN THERE IS A SELECTION !!!!
//
//-----------------------------------------------------------------------------
void CLSRenderer::FixupClipRectForSelection()
{
    long xOrigin;

    if(!_pdp->IsRTL())
    {
        xOrigin = _rcView.left  + _xRelOffset;
    }
    else
    {
        xOrigin = _rcView.right - _xRelOffset;
    }

    _pdp->GetClipRectForLine(&_rcClip, _ptCur.y, xOrigin, &_li);
    IntersectRect(&_rcClip, &_rcClip, _pDI->ClipRect());
}

//+----------------------------------------------------------------------------
//
// Member:      CLSRenderer::AdjustColors()
//
// Synopsis:    This function adjusts colors for selection
//
//-----------------------------------------------------------------------------
void CLSRenderer::AdjustColors(COneRun* por)
{
    COLORREF crNewTextColor, crNewBkColor;
    const CCharFormat* pCF = por->GetCF();
    CColorValue ccvBackColor = por->GetCurrentBgColor();

    if(GetMarkup()->IsEditable(TRUE))
    {
        // If we are in edit mode, then we just invert the text color
        // and the back color and draw in opaque mode. We cannot rely
        // on the hdc to provide us accurate bk color info, (because
        // bk could have been painted earlier and we are painting the
        // text in transparent mode, so the bk color may be invalid)
        // Hence, lets get the TRUE bk color.
        crNewTextColor = ~GetTextColor(_hdc);
        if(ccvBackColor.IsDefined())
        {
            crNewBkColor = ~ccvBackColor.GetColorRef();
        }
        else
        {
            crNewBkColor = ~crNewTextColor;
        }

        ContrastColors(crNewTextColor, crNewBkColor, 100);

        // Finally, turn off the higher order bits
        crNewTextColor &= CColorValue::MASK_COLOR;
        crNewBkColor &= CColorValue::MASK_COLOR;
    }
    else
    {
        Assert(pCF);

        crNewTextColor = GetSysColor (COLOR_HIGHLIGHTTEXT);
        crNewBkColor = GetSysColor (COLOR_HIGHLIGHT);

        if(((CCharFormat*)pCF)->SwapSelectionColors())
        {
            COLORREF crTemp;

            crTemp = crNewBkColor;
            crNewBkColor = crNewTextColor;
            crNewTextColor = crTemp;
        }
    }

    crNewTextColor = GetNearestColor(_hdc, crNewTextColor);
    crNewBkColor   = GetNearestColor(_hdc, crNewBkColor);

    SetTextColor(_hdc, crNewTextColor);
    SetBkColor(_hdc, crNewBkColor);
}

/*
*  CLSRenderer::SetNewFont (BOOL fJustRestore)
*
*  @mfunc
*      Select appropriate font and color in the _hdc based on the
*      current character format. Also sets the background color
*      and mode.
*
*  @rdesc
*      TRUE if it succeeds
*/
CCcs* CLSRenderer::SetNewFont(COneRun* por)
{
    const CCharFormat* pCF = por->GetCF();
    COLORREF cr;
    CCcs* pccs;
    CBaseCcs* pBaseCcs;

    Assert(pCF);

    // get information about disabled
    if(pCF->_fDisabled)
    {
        _fDisabled = TRUE;

        _crForeDisabled = GetSysColorQuick(COLOR_3DSHADOW);
        _crShadowDisabled = GetSysColorQuick(COLOR_3DHILIGHT);

        if(_crForeDisabled == CLR_INVALID || _crShadowDisabled == CLR_INVALID)
        {
            _crForeDisabled = _crShadowDisabled = GetSysColorQuick(COLOR_GRAYTEXT);
        }
    }
    else
    {
        _fDisabled = FALSE;
    }

    // Retrieves new font to use
    Assert(_pDI->_hic);
    Assert(_pDI->GetDC());

    pccs = _pLS->GetCcs(por, _pDI->GetDC(), _pDI);
    if(pccs == NULL)
    {
        goto Cleanup;
    }
    pBaseCcs = pccs->GetBaseCcs();

    // 1. Select font in _hdc
    AssertSz(pBaseCcs->_hfont, "CLSRenderer::SetNewFont pBaseCcs->_hfont is NULL");
    // It's more expensive on WinNT 4.0 to do GetCurrentObject than
    // it is to always select the font.  Go figure.
    {
        SelectFontEx(_hdc, pBaseCcs->_hfont);
    }

    // 2. Select the pen color
    if(por->HasTextColor())
    {
        cr = por->_ccvTextColor.GetColorRef();
    }
    else
    {
        cr = pCF->_ccvTextColor.GetColorRef();
    }

    if(cr == RGB(255,255,255))
    {
        const INT nTechnology = GetDeviceCaps(_hdc, TECHNOLOGY);

        if(nTechnology==DT_RASPRINTER || nTechnology==DT_PLOTTER)
        {
            cr = RGB(0,0,0);
        }
    }
    SetTextColor(_hdc, cr);


    // 3. Set up the drawing mode.
    SetBkMode(_hdc, TRANSPARENT);

Cleanup:
    return pccs;
}

/*
*  CLSRenderer::RenderStartLine()
*
*  @mfunc
*      Render bullet if at start of line
*
*  @rdesc
*      TRUE if this method succeeded
*/
BOOL CLSRenderer::RenderStartLine(COneRun* por)
{
    if(_li._fHasBulletOrNum)
    {
        AssertSz(FALSE, "must improve");
    }

    return TRUE;
}

COneRun* CLSRenderer::FetchLIRun(
                        LSCP lscp,          // IN
                        LPCWSTR* ppwchRun,  // OUT
                        DWORD* pcchRun)     // OUT
{
    COneRun* por = _pLS->_listCurrent._pHead;
    Assert(por);
    Assert(lscp>=por->_lscpBase && lscp<por->_lscpBase+por->_lscch);
    LONG cpOffset = lscp - por->_lscpBase;
    *ppwchRun = por->_pchBase;
    *pcchRun = por->_lscch;
    if(cpOffset > 0)
    {
        *ppwchRun += cpOffset;
        *pcchRun -= cpOffset;
    }
    Assert(*pcchRun > 0);
    _pLS->CheckForUnderLine(por);

    return por;
}

//-----------------------------------------------------------------------------
//
//  Function:   Drawunderline
//
//  Synopsis:   Draws an underline
//
//  Params:     Same as that passed by lineservices to drawunderline
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
LSERR CLSRenderer::DrawUnderline(
                           COneRun* por,           // IN
                           UINT kUlBase,           // IN
                           const POINT* pptStart,  // IN
                           DWORD dupUl,            // IN
                           DWORD dvpUl,            // IN
                           LSTFLOW kTFlow,         // IN
                           UINT kDisp,             // IN
                           const RECT* prcClip)    // IN
{
    LSERR lserr = lserrNone;

    if(por->Cp()>=_cpStartRender && por->Cp()<_cpStopRender)
    {
        const CCharFormat* pCF = por->GetCF();
        POINT ptStart = *pptStart;
        const BOOL fUnderline = por->_dwImeHighlight
            ? por->_fUnderlineForIME : pCF->_fUnderline;
        CCcs* pccs;
        CBaseCcs* pBaseCcs;
        LONG  y;

        pccs = SetNewFont(por);
        if(!pccs)
        {
            lserr = lserrOutOfMemory;
            goto Cleanup;
        }

        pBaseCcs = pccs->GetBaseCcs();
        y = pptStart->y + _li._yHeight - _li._yDescent +
            pBaseCcs->_yDescent - pBaseCcs->_yHeight - pBaseCcs->_yOffset;

        ptStart.x -= _xAccumulatedWidth;

        if(por->HasBgColor())
        {
            AdjustColors(por);
        }

        if(pCF->_fStrikeOut)
        {
            POINT pt = ptStart;
            LONG dwWidth = max(1, ((pBaseCcs->_yHeight/10)/2)*2+1); // must be odd

            pt.y = pptStart->y + _li._yHeight - _li._yDescent -
                pBaseCcs->_yOffset + pBaseCcs->_yDescent -
                (pBaseCcs->_yHeight+1)/2;

            DrawEnabledDisabledLine(kUlBase, &pt, dupUl, dwWidth, kTFlow, kDisp, prcClip);
        }

        if(pCF->_fOverline)
        {
            POINT pt = ptStart;
            pt.y = y;
            DrawEnabledDisabledLine(kUlBase, &pt, dupUl, dvpUl, kTFlow, kDisp, prcClip);
        }

        if(fUnderline)
        {
            POINT pt = ptStart;
            if(pBaseCcs->_yOffset)
            {
                pt.y = y + pBaseCcs->_yHeight - (pBaseCcs->_yDescent+1)/2;
            }
            else
            {
                // BUGBUG (t-ramar): Since we are using _li._yTxtDescent here, this causes a bit of
                // strangeness with underlining of Ruby pronunciation text. Basically, since the pronunciation
                // text is 50% the size of the base text (_li._yTxtDescent is usually based on the base text)
                // underlining of the pronunciation text will appear a little too low
                pt.y = y + (_li._yTxtDescent/2)  - pBaseCcs->_yDescent + pBaseCcs->_yHeight;
            }

            // NOTE(SujalP): Some fonts may have no descent in which we can show the
            // underline, and we may end up with pt.y being below the line. So adjust
            // the underline such that it is always within the line.
            if(pt.y >= pptStart->y+_li.GetYBottom())
            {
                pt.y = pptStart->y + _li.GetYBottom() - 1;
            }

            DrawEnabledDisabledLine(kUlBase, &pt, dupUl, dvpUl, kTFlow, kDisp, prcClip);
        }
    }

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   DrawEnabledDisabledLine
//
//  Synopsis:   Draws an enabled or a disabled line at the specified coordinates
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
LSERR CLSRenderer::DrawEnabledDisabledLine(
                                     UINT    kUlBase,           // IN
                                     const   POINT* pptStart,   // IN
                                     DWORD   dupUl,             // IN
                                     DWORD   dvpUl,             // IN
                                     LSTFLOW kTFlow,            // IN
                                     UINT    kDisp,             // IN
                                     const   RECT* prcClip)     // IN
{
    COLORREF curTxtColor = GetTextColor(_hdc);
    LSERR    lserr       = lserrNone;
    CRect    ulRect;

    // If we're drawing from RtL then we need to shift one pixel to the right
    // in order to account for the left biased drawing of rects (pixel at
    // rc.left is drawn, pixel at rc.right is not).
    ulRect.left   = pptStart->x - ((!(kTFlow&fUDirection))?0:dupUl-1);
    ulRect.top    = pptStart->y;
    ulRect.right  = pptStart->x + ((!(kTFlow&fUDirection))?dupUl:1);
    ulRect.bottom = pptStart->y + dvpUl;

    if(IntersectRect(&ulRect, &ulRect, prcClip))
    {
        int bkModeOld = 0;

        if(_fDisabled)
        {
            if(_crForeDisabled != _crShadowDisabled)
            {
                CRect ulRectDisabled(ulRect);

                //draw the shadow
                curTxtColor = _crShadowDisabled;
                ulRectDisabled.OffsetRect(1, 1);
                lserr = DrawLine(kUlBase, curTxtColor, &ulRectDisabled);
                if(lserr != lserrNone)
                {
                    goto Cleanup;
                }

                // now set the drawing mode to transparent
                bkModeOld = SetBkMode(_hdc, TRANSPARENT);
            }
            curTxtColor = _crForeDisabled;
        }

        // draw the actual line
        lserr = DrawLine(kUlBase, curTxtColor, &ulRect);
        if(lserr != lserrNone)
        {
            goto Cleanup;
        }

        // restore the background mode.
        if(_fDisabled && _crForeDisabled!=_crShadowDisabled)
        {
            SetBkMode(_hdc, bkModeOld);
        }
    }

Cleanup:
    return lserr;
}

//-----------------------------------------------------------------------------
//
//  Function:   DrawLine
//
//  Synopsis:   Draws a line at the specified coordinates
//
//  Returns:    LSERR
//
//-----------------------------------------------------------------------------
LSERR CLSRenderer::DrawLine(
                      UINT     kUlBase,        // IN
                      COLORREF colorUnderLine, // IN
                      CRect*   pRectLine)      // IN
{
    LSERR   lserr = lserrNone;
    HPEN    hPen,   hPenOld;
    UINT    lopnStyle = ((kUlBase&CFU_UNDERLINE_BITS)==CFU_UNDERLINEDOTTED) ? PS_DOT : PS_SOLID;

    hPen = CreatePen(lopnStyle, 0, colorUnderLine);
    if(hPen)
    {
        hPenOld = (HPEN) SelectObject(_hdc, hPen);
        if(pRectLine->bottom - pRectLine->top <= 1)
        {
            MoveToEx(_hdc, pRectLine->left, pRectLine->top, NULL);
            LineTo(_hdc, pRectLine->right, pRectLine->top);
        }
        else
        {
            HBRUSH hBrush, hBrushOld;

            hBrush = CreateSolidBrush(colorUnderLine);
            if(hBrush)
            {
                hBrushOld = (HBRUSH)SelectObject(_hdc, hBrush);
                Rectangle(_hdc, pRectLine->left, pRectLine->top, pRectLine->right, pRectLine->bottom);
                SelectObject(_hdc, hBrushOld);
                DeleteObject(hBrush);
            }
        }
        SelectObject(_hdc, hPenOld);
        DeleteObject(hPen);
    }

    return lserr;
}


UINT g_uLatinCodepage;
BOOL g_fGotLatinCodepage = FALSE;
//+---------------------------------------------------------------------------
//      GetLatinCodepage
//----------------------------------------------------------------------------
UINT GetLatinCodepage()
{
    if(!g_fGotLatinCodepage)
    {
        // When converting Latin-1 characters, we will use g_uLatinCodepage.
        // The first choice is to use Windows-1252 (Windows Latin-1).  That
        // should be present in all systems, unless deliberately removed.  If
        // that fails, we go for our second choice, which is MS-DOS Latin-1.
        // If that fails, we'll just go with CP_ACP. (cthrash)
        if(!IsValidCodePage(g_uLatinCodepage=1252) && !IsValidCodePage(g_uLatinCodepage=850))
        {
            g_uLatinCodepage = CP_ACP;
        }

        g_fGotLatinCodepage = TRUE;
    }

    return g_uLatinCodepage;
}