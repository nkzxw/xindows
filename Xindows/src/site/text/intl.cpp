
#include "stdafx.h"
#include "intl.h"

#include <inetreg.h>

#include "txtdefs.h"

IMultiLanguage*  g_pMultiLanguage  = NULL;
IMultiLanguage2* g_pMultiLanguage2 = NULL; // JIT langpack support

// Globals
PMIMECPINFO g_pcpInfo = NULL;
ULONG       g_ccpInfo = 0;
BOOL        g_cpInfoInitialized = FALSE;

extern BYTE g_bUSPJitState;

static struct
{
    CODEPAGE cp;
    BYTE     bGDICharset;
}
s_aryCpMap[] =
{
    { CP_1252,  ANSI_CHARSET },
    { CP_1252,  SYMBOL_CHARSET },
    { CP_1250,  EASTEUROPE_CHARSET },
    { CP_1251,  RUSSIAN_CHARSET },
    { CP_1253,  GREEK_CHARSET },
    { CP_1254,  TURKISH_CHARSET },
    { CP_1255,  HEBREW_CHARSET },
    { CP_1256,  ARABIC_CHARSET },
    { CP_1257,  BALTIC_CHARSET },
    { CP_1258,  VIETNAMESE_CHARSET },
    { CP_THAI,  THAI_CHARSET },
    { CP_UTF_8, ANSI_CHARSET }
};

// Note: This list is a small subset of the mlang list, used to avoid
// loading mlang in certain cases.
//
// *** There should be at least one entry in this list for each entry
//     in s_aryCpMap so that we can look up the name quickly. ***
//
// NB (cthrash) List obtained from the following URL, courtesy of ChristW
//
//   http://iptdweb/intloc/internal/redmond/projects/ie4/charsets.htm

//BUGBUG - Remove Thai, Arabic DOS and Hebrew DOS
//         from this array after MLANG adds it to
//         the MIME database
static const MIMECSETINFO s_aryInternalCSetInfo[] =
{
    { CP_1252,  CP_ISO_8859_1,  _T("iso-8859-1") },
    { CP_1252,  CP_1252,  _T("windows-1252") },
    { NATIVE_UNICODE_CODEPAGE, CP_UTF_8, _T("utf-8") },
    { NATIVE_UNICODE_CODEPAGE, CP_UTF_8, _T("unicode-1-1-utf-8") },
    { NATIVE_UNICODE_CODEPAGE, CP_UTF_8, _T("unicode-2-0-utf-8") },
    { CP_1250,  CP_1250,  _T("windows-1250") },     // Eastern European windows
    { CP_1251,  CP_1251,  _T("windows-1251") },     // Cyrillic windows
    { CP_1253,  CP_1253,  _T("windows-1253") },     // Greek windows
    { CP_1254,  CP_1254,  _T("windows-1254") },     // Turkish windows
    { CP_1257,  CP_1257,  _T("windows-1257") },     // Baltic windows
    { CP_1258,  CP_1258,  _T("windows-1258") },     // Vietnamese windows
    { CP_1255,  CP_1255,  _T("windows-1255") },     // Hebrew windows
    { CP_1256,  CP_1256,  _T("windows-1256") },     // Arabic windows
    { CP_THAI,  CP_THAI,  _T("windows-874") },      // Thai Windows
};

//+-----------------------------------------------------------------------
//
//  Function:   EnsureCodePageInfo
//
//  Synopsis:   Hook up to mlang.  Note: it is not an error condition to
//              get back no codepage info structures from mlang.
//
//  IE5 JIT langpack support:
//              Now user can install fonts and nls files on the fly 
//              without restarting the session.
//              This means we always have to get real information
//              as to which codepage is valid and not.
//
//------------------------------------------------------------------------
HRESULT EnsureCodePageInfo(BOOL fForceRefresh)
{
    if(g_cpInfoInitialized && !fForceRefresh)
    {
        return S_OK;
    }

    HRESULT         hr;
    UINT            cNum;
    IEnumCodePage*  pEnumCodePage = NULL;

    PMIMECPINFO     pcpInfo = NULL;
    ULONG           ccpInfo = 0;

    hr = EnsureMultiLanguage();
    if(hr)
    {
        goto Cleanup;
    }

    hr = MlangEnumCodePages(MIMECONTF_BROWSER, &pEnumCodePage);
    if(hr)
    {
        goto Cleanup;
    }

    if(g_pMultiLanguage2)
    {
        hr = g_pMultiLanguage2->GetNumberOfCodePageInfo(&cNum);
    }
    else
    {
        hr = g_pMultiLanguage->GetNumberOfCodePageInfo(&cNum);
    }
    if(hr)
    {
        goto Cleanup;
    }

    if(cNum > 0)
    {
        pcpInfo = (PMIMECPINFO)MemAlloc(sizeof(MIMECPINFO)*(cNum));
        if(!pcpInfo)
        {
            hr = E_OUTOFMEMORY;
            goto Cleanup;
        }
        hr = pEnumCodePage->Next(cNum, pcpInfo, &ccpInfo);
        if(hr)
        {
            goto Cleanup;
        }

        if(ccpInfo > 0)
        {
            MemRealloc((void**)&pcpInfo, sizeof(MIMECPINFO)*(ccpInfo));
        }
        else
        {
            MemFree(pcpInfo);
            pcpInfo = NULL;
        }
    }

Cleanup:
    {
        LOCK_GLOBALS;

        if(!hr)
        {
            if(g_pcpInfo)
            {
                MemFree(g_pcpInfo);
            }

            g_pcpInfo = pcpInfo;
            g_ccpInfo = ccpInfo;
        }

        // If any part of the initialization fails, we do not want to keep
        // returning this function, unless, of course, fForceRefresh is true.
        g_cpInfoInitialized = TRUE;
    }

    ReleaseInterface(pEnumCodePage);
    RRETURN(hr);
}

//+-----------------------------------------------------------------------
//
//  Function:   DeinitMultiLanguage
//
//  Synopsis:   Detach from mlang.  Called from DllAllThreadsDetach.
//              Globals are locked during the call.
//
//------------------------------------------------------------------------
void DeinitMultiLanguage()
{
    MemFree(g_pcpInfo);
    g_cpInfoInitialized = FALSE;
    g_pcpInfo = NULL;
    g_ccpInfo = 0;
    ClearInterface(&g_pMultiLanguage2);
    ClearInterface(&g_pMultiLanguage);

}

//+-----------------------------------------------------------------------
//
//  Function:   EnsureMultiLanguage
//
//  Synopsis:   Make sure mlang is loaded.
//
//------------------------------------------------------------------------
HRESULT EnsureMultiLanguage()
{
    if(g_pMultiLanguage)
    {
        return S_OK;
    }

    HRESULT hr = S_OK;

    LOCK_GLOBALS;

    // Need to check again after locking globals.
    if(g_pMultiLanguage)
    {
        goto Cleanup;
    }

    hr = CoCreateInstance(
        CLSID_CMultiLanguage,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_IMultiLanguage,
        (void**)&g_pMultiLanguage);
    if(hr)
    {
        goto Cleanup;
    }

    // JIT langpack
    // get the aggregated ML2, this may fail if IE5 is not present
    g_pMultiLanguage->QueryInterface(IID_IMultiLanguage2, (void**)&g_pMultiLanguage2);

Cleanup:
    RRETURN(hr);
}

//+-----------------------------------------------------------------------
//
//  Function:   MlangEnumCodePages
//
//              Just a utility function that wraps IML(2)::EnumCodePages
//              supposedly use the curernt UI language for mlang call
//
// bugbug: do we really want to live without IE5?
//
//------------------------------------------------------------------------
HRESULT MlangEnumCodePages(DWORD grfFlags, IEnumCodePage** ppEnumCodePage)
{
    HRESULT hr = EnsureMultiLanguage();

    if(hr==S_OK && ppEnumCodePage)
    {
        if(g_pMultiLanguage2)
        {
            hr = g_pMultiLanguage2->EnumCodePages(grfFlags, GetUserDefaultUILanguage(), ppEnumCodePage);
        }
        else
        {
            hr = g_pMultiLanguage->EnumCodePages(grfFlags,  ppEnumCodePage);
        }
    }
    RRETURN(hr);
}

//+-----------------------------------------------------------------------
//
//  Function:   QuickMimeGetCodePageInfo
//
//  Synopsis:   This function gets cp info from mlang, but caches some
//              values to avoid loading mlang unless necessary.
//
//------------------------------------------------------------------------
HRESULT SlowMimeGetCodePageInfo(CODEPAGE cp, PMIMECPINFO pMimeCpInfo)
{
    HRESULT hr = EnsureMultiLanguage();
    if(hr == S_OK)
    {
        if(g_pMultiLanguage2)
        {
            hr = g_pMultiLanguage2->GetCodePageInfo(cp, GetUserDefaultUILanguage(), pMimeCpInfo);
        }
        else
        {
            hr = g_pMultiLanguage->GetCodePageInfo(cp, pMimeCpInfo);
        }
    }
    return hr;
}

HRESULT QuickMimeGetCodePageInfo(CODEPAGE cp, PMIMECPINFO pMimeCpInfo)
{
    HRESULT hr = S_OK;
    static MIMECPINFO s_mimeCpInfoBlank =
    {
        0, 0, 0, _T(""), _T(""), _T(""), _T(""),
        _T("Courier New"), _T("MS Sans Serif"), DEFAULT_CHARSET
    };

    // If mlang is not loaded and it is acceptable to use our defaults,
    //  try to avoid loading it by searching through our internal
    //  cache.
    for(int n=0; n<ARRAYSIZE(s_aryCpMap); ++n)
    {
        if(cp == s_aryCpMap[n].cp)
        {
            memcpy(pMimeCpInfo, &s_mimeCpInfoBlank, sizeof(MIMECPINFO));
            pMimeCpInfo->uiCodePage = pMimeCpInfo->uiFamilyCodePage = s_aryCpMap[n].cp;
            pMimeCpInfo->bGDICharset = s_aryCpMap[n].bGDICharset;

            // Search for the name in the cset array
            int j;
            for(j=0; j<ARRAYSIZE(s_aryInternalCSetInfo); ++j)
            {
                if(cp == s_aryInternalCSetInfo[j].uiInternetEncoding)
                {
                    _tcscpy(pMimeCpInfo->wszWebCharset, s_aryInternalCSetInfo[j].wszCharset);
                    break;
                }
            }

            // Assert there was an entry in s_aryInternalCSetInfo for this codepage
            Assert(j < ARRAYSIZE(s_aryInternalCSetInfo));
            return S_OK;
        }
    }

    hr = SlowMimeGetCodePageInfo(cp, pMimeCpInfo);
    if(hr)
    {
        goto Error;
    }

Cleanup:
    return hr;

Error:
    // Could not load mlang, fill in with a default but return the error
    memcpy(pMimeCpInfo, &s_mimeCpInfoBlank, sizeof(MIMECPINFO));
    goto Cleanup;
}

//+-----------------------------------------------------------------------
//
//  Function:   QuickMimeGetCharsetInfo
//
//  Synopsis:   This function gets charset info from mlang, but caches some
//              values to avoid loading mlang unless necessary.
//
//------------------------------------------------------------------------
HRESULT QuickMimeGetCharsetInfo(LPCTSTR lpszCharset, PMIMECSETINFO pMimeCSetInfo)
{
    HRESULT hr = S_OK;

    // If mlang is not loaded and it is acceptable to use our defaults,
    //  try to avoid loading it by searching through our internal
    //  cache.
    for(int n=0; n<ARRAYSIZE(s_aryInternalCSetInfo); ++n)
    {
        if(_tcsicmp((TCHAR*)lpszCharset, s_aryInternalCSetInfo[n].wszCharset) == 0)
        {
            *pMimeCSetInfo = s_aryInternalCSetInfo[n];
            return S_OK;
        }
    }

    hr = EnsureMultiLanguage();
    if(hr)
    {
        goto Cleanup;
    }

    if(g_pMultiLanguage2)
    {
        hr = g_pMultiLanguage2->GetCharsetInfo((LPTSTR)lpszCharset, pMimeCSetInfo);
    }
    else
    {
        hr = g_pMultiLanguage->GetCharsetInfo((LPTSTR)lpszCharset, pMimeCSetInfo);
    }
    if(hr)
    {
        goto Cleanup;
    }

Cleanup:
    return hr;
}

//+-----------------------------------------------------------------------
//
//  Function:   GetCodePageFromMlangString
//
//  Synopsis:   Return a codepage id from an mlang codepage string.
//
//------------------------------------------------------------------------
HRESULT GetCodePageFromMlangString(LPCTSTR pszMlangString, CODEPAGE* pCodePage)
{
    HRESULT hr;
    MIMECSETINFO mimeCharsetInfo;

    hr = QuickMimeGetCharsetInfo(pszMlangString, &mimeCharsetInfo);
    if(hr)
    {
        hr = E_INVALIDARG;
        *pCodePage = CP_UNDEFINED;
    }
    else
    {
        *pCodePage = mimeCharsetInfo.uiInternetEncoding;
    }

    return hr;
}

//+-----------------------------------------------------------------------
//
//  Function:   GetMlangStringFromCodePage
//
//  Synopsis:   Return an mlang codepage string from a codepage id.
//
//------------------------------------------------------------------------
HRESULT GetMlangStringFromCodePage(CODEPAGE codepage, LPTSTR pMlangStringRet, size_t cchMlangString)
{
    HRESULT    hr;
    MIMECPINFO mimeCpInfo;
    TCHAR*     pCpString = _T("<undefined-cp>");
    size_t     cchCopy;

    Assert(codepage != CP_ACP);

    hr = QuickMimeGetCodePageInfo(codepage, &mimeCpInfo);
    if(hr == S_OK)
    {
        pCpString = mimeCpInfo.wszWebCharset;
    }

    cchCopy = min(cchMlangString-1, _tcslen(pCpString));
    _tcsncpy(pMlangStringRet, pCpString, cchCopy);
    pMlangStringRet[cchCopy] = 0;

    return hr;
}

//+-----------------------------------------------------------------------
//
//  Function:   WindowsCodePageFromCodePage
//
//  Synopsis:   Return a Windows codepage from an mlang CODEPAGE
//
//------------------------------------------------------------------------
UINT WindowsCodePageFromCodePage(CODEPAGE cp)
{
    HRESULT hr;

    Assert(cp != CP_UNDEFINED);

    if(cp == CP_AUTO)
    {
        // cp 50001 (CP_AUTO)is designated to cross-language detection,
        // it really should not come in here but we'd return the default 
        // code page.
        return _afxGlobalData._cpDefault;
    }

    if(cp==CP_1252 || cp==CP_ISO_8859_1)
    {
        return 1252;        // short-circuit most common case
    }
    else if(IsStraightToUnicodeCodePage(cp))
    {
        return CP_UCS_2;    // BUGBUG (cthrash) Should be NATIVE_UNICODE_CODEPAGE?
    }
    else if(g_cpInfoInitialized)
    {
        // Get the codepage from mlang
        hr = EnsureCodePageInfo(FALSE);
        if(hr)
        {
            goto Cleanup;
        }

        {
            LOCK_GLOBALS;

            for(UINT n=0; n<g_ccpInfo; n++)
            {
                if(cp == g_pcpInfo[n].uiCodePage)
                {
                    return g_pcpInfo[n].uiFamilyCodePage;
                }
            }
        }
    }

    // BUGBUG (cthrash) There's a chance that this codepage is a 'hidden'
    // codepage and that MLANG may actually know the family codepage.
    // So we ask them again, only differently.
    hr = EnsureMultiLanguage();
    if(hr)
    {
        goto Cleanup;
    }

    {
        UINT uiFamilyCodePage;

        if(g_pMultiLanguage2)
        {
            hr = g_pMultiLanguage2->GetFamilyCodePage(cp, &uiFamilyCodePage);
        }
        else
        {
            hr = g_pMultiLanguage->GetFamilyCodePage(cp, &uiFamilyCodePage);
        }
        if(hr)
        {
            goto Cleanup;
        }
        else
        {
        return uiFamilyCodePage;
        }
    }

Cleanup:
    return GetACP();
}

//+-----------------------------------------------------------------------
//
//  Function:   WindowsCharsetFromCodePage
//
//  Synopsis:   Return a Windows charset from an mlang CODEPAGE id
//
//------------------------------------------------------------------------
BYTE WindowsCharsetFromCodePage(CODEPAGE cp)
{
    HRESULT    hr;
    MIMECPINFO mimeCpInfo;

    if(cp == CP_ACP)
    {
        return DEFAULT_CHARSET;
    }

    hr = QuickMimeGetCodePageInfo(cp, &mimeCpInfo);
    if(hr)
    {
        return DEFAULT_CHARSET;
    }
    else
    {
        return mimeCpInfo.bGDICharset;
    }
}

//+-----------------------------------------------------------------------
//
//  Function:   DefaultCodePageFromCharSet
//
//  Synopsis:   Return a Windows codepage from a Windows font charset
//
//------------------------------------------------------------------------
UINT DefaultCodePageFromCharSet(BYTE bCharSet, CODEPAGE cp, LCID lcid)
{
    HRESULT hr;
    UINT    n;
    static  BYTE bCharSetPrev = DEFAULT_CHARSET;
    static  CODEPAGE cpPrev = CP_UNDEFINED;
    static  CODEPAGE cpDefaultPrev = CP_UNDEFINED;
    CODEPAGE cpDefault;

    if(DEFAULT_CHARSET==bCharSet
        || (bCharSet==ANSI_CHARSET && (cp==CP_UCS_2 || cp==CP_UTF_8)))
    {
        return _afxGlobalData._cpDefault; // Don't populate the statics.
    }
    else if(bCharSet==bCharSetPrev && cpPrev==cp)
    {
        // Here's our gamble -- We have a high likelyhood of calling with
        // the same arguments over and over.  Short-circuit this case.
        return cpDefaultPrev;
    }

    // First pick the *TRUE* codepage
    if(lcid)
    {
        char pszCodePage[5];

        GetLocaleInfoA( lcid, LOCALE_IDEFAULTANSICODEPAGE, pszCodePage, ARRAYSIZE(pszCodePage) );

        cp = atoi(pszCodePage);
    }

    // First check our internal lookup table in case we can avoid mlang
    for(n=0; n<ARRAYSIZE(s_aryCpMap); ++n)
    {
        if(cp==s_aryCpMap[n].cp && bCharSet==s_aryCpMap[n].bGDICharset)
        {
            cpDefault = cp;
            goto Cleanup;
        }
    }

    hr = EnsureCodePageInfo(FALSE);
    if(hr)
    {
        cpDefault = WindowsCodePageFromCodePage(cp);
        goto Cleanup;
    }

    {
        LOCK_GLOBALS;

        // First see if we find an exact match for both cp and bCharset.
        for(n=0; n<g_ccpInfo; n++)
        {
            if(cp==g_pcpInfo[n].uiCodePage && bCharSet==g_pcpInfo[n].bGDICharset)
            {
                cpDefault = g_pcpInfo[n].uiFamilyCodePage;
                goto Cleanup;
            }
        }

        // Settle for the first match of bCharset.
        for(n=0; n<g_ccpInfo; n++)
        {
            if(bCharSet == g_pcpInfo[n].bGDICharset)
            {
                cpDefault = g_pcpInfo[n].uiFamilyCodePage;
                goto Cleanup;
            }
        }

        cpDefault = _afxGlobalData._cpDefault;
    }

Cleanup:
    bCharSetPrev = bCharSet;
    cpPrev = cp;
    cpDefaultPrev = cpDefault;

    return cpDefault;
}

//+-----------------------------------------------------------------------
//
//  Function:   DefaultFontInfoFromCodePage
//
//  Synopsis:   Fills a LOGFONT structure with appropriate information for
//              a 'default' font for a given codepage.
//
//------------------------------------------------------------------------
HRESULT DefaultFontInfoFromCodePage(CODEPAGE cp, LOGFONT* lplf)
{
    MIMECPINFO mimeCpInfo;
    HFONT      hfont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    // The strategy is thus: If we ask for a stock font in the default
    // (CP_ACP) codepage, return the logfont information of that font.
    // If we don't, replace key pieces of information.  Note we could
    // do better -- we could get the right lfPitchAndFamily.
    if(!hfont)
    {
        CPINFO cpinfo;

        GetCPInfo(WindowsCodePageFromCodePage(cp), &cpinfo);

        hfont = (HFONT)((cpinfo.MaxCharSize==1) ? GetStockObject(ANSI_VAR_FONT) : GetStockObject(SYSTEM_FONT));

        AssertSz(hfont, "We'd better have a font now.");
    }

    GetObject(hfont, sizeof(LOGFONT), (LPVOID)lplf);

    if(cp!=CP_ACP
        && cp!=_afxGlobalData._cpDefault
        && (cp!=CP_ISO_8859_1 || _afxGlobalData._cpDefault!=CP_1252))
    {
        QuickMimeGetCodePageInfo(cp, &mimeCpInfo);

        _tcscpy(lplf->lfFaceName, mimeCpInfo.wszProportionalFont);
        lplf->lfCharSet = mimeCpInfo.bGDICharset;
    }
    else
    {
        // NB (cthrash) On both simplified and traditional Chinese systems,
        // we get a bogus lfCharSet value when we ask for the DEFAULT_GUI_FONT.
        // This later confuses CCcs::MakeFont -- so override.
        if(cp == 950)
        {
            lplf->lfCharSet = CHINESEBIG5_CHARSET;
        }
        else if(cp == 936)
        {
            lplf->lfCharSet = GB2312_CHARSET;
        }
    }

    return S_OK;
}

//+-----------------------------------------------------------------------
//
//  Function:   CodePageFromString
//
//  Synopsis:   Map a charset to a forms3 CODEPAGE enum.  Searches in the
//              argument a string of the form charset=xxx.  This is used
//              by the META tag handler in the HTML preparser.
//
//              If fLookForWordCharset is TRUE, pch is presumed to be in
//              the form of charset=XXX.  Otherwise the string is
//              expected to contain just the charset string.
//
//------------------------------------------------------------------------
CODEPAGE CodePageFromString(TCHAR* pch, BOOL fLookForWordCharset)
{
    CODEPAGE cp = CP_UNDEFINED;

    while(pch && *pch)
    {
        for(;IsWhite(*pch);pch++) ;

        if(!fLookForWordCharset || (_tcslen(pch)>=7 && _tcsnicmp(pch, 7, _T("charset"), 7)==0))
        {
            if(fLookForWordCharset)
            {
                pch = _tcschr(pch, L'=');
                pch = pch ? ++pch : NULL;
            }

            if(pch)
            {
                for(;IsWhite(*pch);pch++) ;

                if(*pch)
                {
                    TCHAR *pchEnd, chEnd;

                    for(pchEnd=pch;
                        *pchEnd&&!(*pchEnd==L';'||IsWhite(*pchEnd));
                        pchEnd++) ;

                    chEnd = *pchEnd;
                    *pchEnd = L'\0';

                    cp = CodePageFromAlias(pch);

                    *pchEnd = chEnd;

                    break;
                }
            }
        }

        if(pch)
        {
            pch = _tcschr(pch, L';');
            if(pch) pch++;
        }

    }

    return cp;
}

LCID LCIDFromString(TCHAR* pchArg)
{
    LCID lcid = 0;
    HRESULT hr;

    if(!pchArg)
    {
        goto Cleanup;
    }

    hr = EnsureMultiLanguage();
    if(hr || !g_pMultiLanguage)
    {
        goto Cleanup;
    }

    hr = g_pMultiLanguage->GetLcidFromRfc1766(&lcid, pchArg);
    if(hr)
    {
        lcid = 0;
    }

Cleanup:
    return lcid;
}

// **************************************************************************
// NB (cthrash) From RichEdit (_uwrap/unicwrap) start {

/*
*  CharSetIndexFromChar (ch)
*
*  @mfunc
*      returns index into CharSet/CodePage table rgCpgChar corresponding
*      to the Unicode character ch provided such an assignment is
*      reasonably unambiguous, that is, the currently assigned Unicode
*      characters in various ranges have Windows code-page equivalents.
*      Ambiguous or impossible assignments return UNKNOWN_INDEX, which
*      means that the character can only be represented by Unicode in this
*      simple model.  Note that both UNKNOWN_INDEX and HAN_INDEX are negative
*      values, i.e., they imply further processing to figure out what (if
*      any) charset index to use.  Other indices may also require run
*      processing, such as the blank in BiDi text.  We need to mark our
*      right-to-left runs with an Arabic or Hebrew char set, while we mark
*      left-to-right runs with a left-to-right char set.
*
*  @rdesc
*      returns CharSet index
*/
CHARSETINDEX CharSetIndexFromChar(TCHAR ch) // Unicode character to examine
{
    if(ch < 256)
    {
        return ANSI_INDEX;
    }

    if(ch < 0x700)
    {
        if(ch >= 0x600)
        {
            return ARABIC_INDEX;
        }

        if(ch > 0x590)
        {
            return HEBREW_INDEX;
        }

        if(ch < 0x500)
        {
            if(ch >= 0x400)
            {
                return RUSSIAN_INDEX;
            }

            if(ch >= 0x370)
            {
                return GREEK_INDEX;
            }
        }
    }
    else if(ch < 0xAC00)
    {
        if(ch >= 0x3400)                // CJK Ideographs
        {
            return HAN_INDEX;
        }

        if(ch>0x3040 && ch<0x3100)      // Katakana and Hiragana
        {
            return SHIFTJIS_INDEX;
        }

        if(ch<0xe80 && ch>=0xe00)       // Thai
        {
            return THAI_INDEX;
        }
    }
    else if(ch < 0xD800)
    {
        return HANGUL_INDEX;
    }
    else if(ch > 0xff00)
    {
        if(ch < 0xff65)                 // Fullwidth ASCII and halfwidth
        {
            return HAN_INDEX;           // CJK punctuation
        }
        if(ch < 0xffA0)                 // Halfwidth Katakana
        {
            return SHIFTJIS_INDEX;
        }
        if(ch < 0xffe0)                 // Halfwidth Jamo
        {
            return HANGUL_INDEX;
        }
        if(ch < 0xffef)                 // Fullwidth punctuation and currency
        {
            return HAN_INDEX;           // signs; halfwidth forms, arrows
        }                               // and shapes
    }

    return UNKNOWN_INDEX;
}

/*
*  CheckDBCInUnicodeStr (ptext)
*
*  @mfunc
*      returns TRUE if there is a DBC in the Unicode buffer
*
*  @rdesc
*      returns TRUE | FALSE
*/
BOOL CheckDBCInUnicodeStr(TCHAR* ptext)
{
    CHARSETINDEX iCharSetIndex;

    if(ptext)
    {
        while(*ptext)
        {
            iCharSetIndex = CharSetIndexFromChar(*ptext++);

            if(iCharSetIndex==HAN_INDEX || iCharSetIndex==SHIFTJIS_INDEX || iCharSetIndex==HANGUL_INDEX)
            {
                return TRUE;
            }
        }
    }

    return FALSE;
}

/*
*  GetLocaleLCID ()
*
*  @mfunc      Maps an LCID for thread to a Code Page
*
*  @rdesc      returns Code Page
*/
LCID GetLocaleLCID()
{
    return GetThreadLocale();
}

/*
*  GetLocaleCodePage ()
*
*  @mfunc      Maps an LCID for thread to a Code Page
*
*  @rdesc      returns Code Page
*/
UINT GetLocaleCodePage()
{
    return ConvertLanguageIDtoCodePage(GetThreadLocale());
}

/*
*  GetKeyboardLCID ()
*
*  @mfunc      Gets LCID for keyboard active on current thread
*
*  @rdesc      returns Code Page
*/
LCID GetKeyboardLCID()
{
    return (WORD)(DWORD_PTR)GetKeyboardLayout(0 /*idThread*/);
}

/*
*  IsFELCID(lcid)
*
*  @mfunc
*      Returns TRUE iff lcid is for a FE country.
*
*  @rdesc
*      TRUE iff lcid is for a FE country.
*
*/
BOOL IsFELCID(LCID lcid)
{
    switch(PRIMARYLANGID(LANGIDFROMLCID(lcid)))
    {
    case LANG_CHINESE:
    case LANG_JAPANESE:
    case LANG_KOREAN:
        return TRUE;
    }

    return FALSE;
}

/*
*  IsFECharset(bCharSet)
*
*  @mfunc
*      Returns TRUE iff charset may be for a FE country.
*
*  @rdesc
*      TRUE iff charset may be for a FE country.
*
*/
BOOL IsFECharset(BYTE bCharSet)
{
    switch(bCharSet)
    {
    case CHINESEBIG5_CHARSET:
    case SHIFTJIS_CHARSET:
    case HANGEUL_CHARSET:
    case JOHAB_CHARSET:
    case GB2312_CHARSET:
        return TRUE;
    }

    return FALSE;
}

/*
*  IsRtlLCID(lcid)
*
*  @mfunc
*      Returns TRUE iff lcid is for a RTL script.
*
*  @rdesc
*      TRUE iff lcid is for a RTL script.
*
*/
BOOL IsRtlLCID(LCID lcid)
{
    return (IsRTLLang(PRIMARYLANGID(LANGIDFROMLCID(lcid))));
}

/*
*  IsRTLLang(lcid)
*
*  @mfunc
*      Returns TRUE iff lang is for a RTL script.
*
*  @rdesc
*      TRUE iff lang is for a RTL script.
*
*/
BOOL IsRTLLang(LANGID lang)
{
    switch(lang)
    {
    case LANG_ARABIC:
    case LANG_HEBREW:
    case LANG_URDU:
    case LANG_FARSI:
    case LANG_YIDDISH:
    case LANG_SINDHI:
    case LANG_KASHMIRI:
        return TRUE;
    }

    return FALSE;
}

//+----------------------------------------------------------------------------
//  Member:     GetScriptProperties(eScript)
//
//  Synopsis:   Return a pointer to the script properties describing the script
//              eScript.
//-----------------------------------------------------------------------------
static const SCRIPT_PROPERTIES** s_ppScriptProps = NULL;
static int s_cScript = 0;
static const SCRIPT_PROPERTIES s_ScriptPropsDefault =
{
    LANG_NEUTRAL,   // langid
    FALSE,          // fNumeric
    FALSE,          // fComplex
    FALSE,          // fNeedsWordBreaking
    FALSE,          // fNeedsCaretInfo
    ANSI_CHARSET,   // bCharSet
    FALSE,          // fControl
    FALSE,          // fPrivateUseArea
    FALSE,          // fReserved
};

const SCRIPT_PROPERTIES* GetScriptProperties(WORD eScript)
{
    if(s_ppScriptProps == NULL)
    {
        HRESULT hr;

        Assert(s_cScript == 0);
        if(g_bUSPJitState == JIT_OK)
        {
            hr = ::ScriptGetProperties(&s_ppScriptProps, &s_cScript);
        }
        else
        {
            hr = E_PENDING;
        }

        if(FAILED(hr))
        {
            // This should only fail if USP cannot be loaded. We shouldn't
            // really have made it here in the first place if this is true,
            // but you never know...
            return &s_ScriptPropsDefault;
        }
    }
    Assert(s_ppScriptProps!=NULL && eScript<s_cScript && s_ppScriptProps[eScript]!=NULL);
    return s_ppScriptProps[eScript];
}

//+----------------------------------------------------------------------------
//  Member:     GetNumericScript(lang)
//
//  Synopsis:   Returns the script that should be used to shape digits in the
//              given language.
//-----------------------------------------------------------------------------
const WORD GetNumericScript(DWORD lang)
{
    WORD eScript = 0;

    // We should never get here without having called GetScriptProperties().
    Assert(s_ppScriptProps!=NULL && eScript<s_cScript && s_ppScriptProps[eScript]!=NULL);
    for(eScript=0; eScript<s_cScript; eScript++)
    {
        if(s_ppScriptProps[eScript]->langid==lang && s_ppScriptProps[eScript]->fNumeric)
        {
            return eScript;
        }
    }

    return SCRIPT_UNDEFINED;
}

//+----------------------------------------------------------------------------
//  Member:     ScriptItemize(...)
//
//  Synopsis:   Dynamically grows the needed size of paryItems as needed to
//              successfully itemize the input string.
//-----------------------------------------------------------------------------
HRESULT WINAPI ScriptItemize(
         PCWSTR                 pwcInChars,     // In   Unicode string to be itemized
         int                    cInChars,       // In   Character count to itemize
         int                    cItemGrow,      // In   Items to grow by if paryItems is too small
         const SCRIPT_CONTROL*  psControl,      // In   Analysis control (optional)
         const SCRIPT_STATE*    psState,        // In   Initial bidi algorithm state (optional)
         CDataAry<SCRIPT_ITEM>* paryItems,      // Out  Array to receive itemization
         PINT                   pcItems)        // Out  Count of items processed
{
    HRESULT hr;

    // ScriptItemize requires that the max item buffer size be AT LEAST 2
    Assert(cItemGrow > 2);

    if(paryItems->Size() < 2)
    {
        hr = paryItems->Grow(cItemGrow);
        if(FAILED(hr))
        {
            goto Cleanup;
        }
    }

    do
    {
        hr = ScriptItemize(pwcInChars, cInChars, paryItems->Size(),
            psControl, psState, (SCRIPT_ITEM*)*paryItems, pcItems);

        if(hr == E_OUTOFMEMORY)
        {
            if(FAILED(paryItems->Grow(paryItems->Size()+cItemGrow)))
            {
                goto Cleanup;
            }
        }
    } while(hr == E_OUTOFMEMORY);

Cleanup:
    if(SUCCEEDED(hr))
    {
        // NB (mikejoch) *pcItems doesn't include the sentinel item.
        Assert(*pcItems < paryItems->Size());
        paryItems->SetSize(*pcItems+1);
    }
    else
    {
        *pcItems = 0;
        paryItems->DeleteAll();
    }

    return hr;
}

// (_uwrap/unicwrap) end }
// **************************************************************************

// BUGBUG (cthrash) This class (CIntlFont) should be axed when we implement
// a light-weight fontlinking implementation of Line Services.  This LS
// implementation will work as a DrawText replacement, and can be used by
// intrinsics as well.
CIntlFont::CIntlFont(HDC hdc, CODEPAGE codepage, LCID lcid, SHORT sBaselineFont, const TCHAR* psz)
{
    BYTE bCharSet = WindowsCharsetFromCodePage(codepage);

    _hdc = hdc;
    _hFont = NULL;

    Assert(sBaselineFont>=0 && sBaselineFont<=4);
    Assert(psz);

    if(IsStraightToUnicodeCodePage(codepage))
    {
        BOOL fSawHan = FALSE;
        SCRIPT_ID sid;

        // If the document is in a Unicode codepage, we need determine the
        // best-guess charset for this string.  To do so, we pick the first
        // interesting script id.
        while(*psz)
        {
            sid = ScriptIDFromCh(*psz++);

            if(sid == sidHan)
            {
                fSawHan = TRUE;
                continue;
            }

            if(sid > sidAsciiLatin)
            {
                break;
            }
        }

        if(*psz)
        {
            // We found something interesting, go pick that font up
            codepage = DefaultCodePageFromScript(&sid, CP_UCS_2, lcid);
        }
        else if(!fSawHan)
        {
            // the string contained nothing interesting, go with the stock GUI font
            _hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
            _fIsStock = sBaselineFont == 2;
        }
        else
        {
            // We saw a Han character, but nothing else which would
            // disambiguate the script.  Furthermore, we don't have a good
            // fallback as we did in the above case.
            sid = sidHan;
            codepage = DefaultCodePageFromScript(&sid, CP_UCS_2, lcid);
        }
    }
    else if(ANSI_CHARSET == bCharSet)
    {
        // If we're looking for an ANSI font, or if we're under NT and
        // looking for a non-FarEast font, the stockfont will suffice.
        _hFont = (HFONT)GetStockObject(ANSI_VAR_FONT);
        _fIsStock = (sBaselineFont == 2);
    }
    else
    {
        codepage = WindowsCodePageFromCodePage(codepage);
    }

    if(!_hFont && codepage==_afxGlobalData._cpDefault)
    {
        // If we're going to get the correct native charset the, the
        // GUI font will work nicely.
        _hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        _fIsStock = (sBaselineFont == 2);
    }

    if(_hFont)
    {
        if(!_fIsStock)
        {
            LOGFONT lf;
            GetObject(_hFont, sizeof(lf), &lf);
            lf.lfHeight = MulDivQuick(lf.lfHeight, 4+sBaselineFont, 6);
            lf.lfWidth  = 0;
            lf.lfOutPrecision |= OUT_TT_ONLY_PRECIS;

            _hFont = CreateFontIndirect(&lf);
        }
    }
    else
    {
        // We'd better cook up a font if all else fails.
        LOGFONT lf;
        DefaultFontInfoFromCodePage(codepage, &lf);

        lf.lfHeight = MulDivQuick(lf.lfHeight, 4+sBaselineFont, 6);
        lf.lfWidth  = 0;
        lf.lfOutPrecision |= OUT_TT_ONLY_PRECIS;

        _hFont = CreateFontIndirect(&lf);
        _fIsStock = FALSE;
    }

    _hOldFont = (HFONT)SelectObject(hdc, _hFont);
}

CIntlFont::~CIntlFont()
{
    SelectObject(_hdc, _hOldFont);

    if(!_fIsStock)
    {
        DeleteObject(_hFont);
    }
}

// "Far East" characters
BOOL IsFEChar(TCHAR ch)
{
    const SCRIPT_ID sid = ScriptIDFromCh(ch);

    return (sid-sidFEFirst)<=(sidFELast-sidFEFirst);
}

//+-----------------------------------------------------------------------------
//
//  Function:   MlangGetDefaultFont
//
//  Synopsis:   Queries MLANG for a font that supports a particular script.
//
//  Returns:    An atom value for the script-appropriate font.  0 on error.
//              0 implies the system font.
//
//------------------------------------------------------------------------------
HRESULT MlangGetDefaultFont(SCRIPT_ID sid, SCRIPTINFO* psi)
{
    IEnumScript* pEnumScript = NULL;
    HRESULT hr;

    Assert(psi);

    hr = EnsureMultiLanguage();
    if(hr)
    {
        goto Cleanup;
    }

    if(g_pMultiLanguage2)
    {
        UINT cNum;
        SCRIPTINFO si;

        hr = g_pMultiLanguage2->GetNumberOfScripts(&cNum);
        if(hr)
        {
            goto Cleanup;
        }

        hr = g_pMultiLanguage2->EnumScripts(SCRIPTCONTF_SCRIPT_USER, 0, &pEnumScript);
        if(hr)
        {
            goto Cleanup;
        }

        while(cNum--)
        {
            ULONG c;

            hr = pEnumScript->Next(1, &si, &c);
            if(hr || !c)
            {
                hr = !c ? E_FAIL : hr;
                goto Cleanup;
            }

            if(si.ScriptId == sid)
            {
                *psi = si;
                break;
            }
        }
    }

Cleanup:
    ReleaseInterface(pEnumScript);
    RRETURN(hr);
}

//+-----------------------------------------------------------------------------
//
//  Function:   IsNarrowCharset(BYTE bCharSet)
//
//  Returns:    FALSE if charset is SHIFTJIS (128), HANGEUL (129), GB2312 (134)
//              or CHINESEBIG5 (136).
//              TRUE for anything else.
//
//------------------------------------------------------------------------------
BOOL IsNarrowCharSet(BYTE bCharSet)
{
    // Hack (cthrash) GDI does not define (currently, anyway) charsets 130, 131
    // 132, 133 or 135.  Make the test simpler by rejecting everything outside
    // of [SHIFTJIS,CHINESEBIG5].
    return (unsigned int)(bCharSet-SHIFTJIS_CHARSET)>(unsigned int)(CHINESEBIG5_CHARSET-SHIFTJIS_CHARSET);
}

CODEPAGE DefaultCodePageFromCharSet(BYTE bCharSet, UINT uiFamilyCodePage)
{
    CODEPAGE cp;

    switch(bCharSet)
    {
    default:
    case DEFAULT_CHARSET:
    case OEM_CHARSET:
    case SYMBOL_CHARSET:        cp = uiFamilyCodePage;  break;
    case SHIFTJIS_CHARSET:      cp = CP_JPN_SJ;         break;
    case HANGUL_CHARSET:        cp = CP_KOR_5601;       break;
    case GB2312_CHARSET:        cp = CP_CHN_GB;         break;
    case CHINESEBIG5_CHARSET:   cp = CP_TWN;            break;
    case JOHAB_CHARSET:         cp = CP_KOR_5601;       break;
    case EASTEUROPE_CHARSET:    cp = CP_1250;           break;
    case RUSSIAN_CHARSET:       cp = CP_1251;           break;
    case ANSI_CHARSET:          cp = CP_1252;           break;
    case GREEK_CHARSET:         cp = CP_1253;           break;
    case TURKISH_CHARSET:       cp = CP_1254;           break;
    case HEBREW_CHARSET:        cp = CP_1255;           break;
    case ARABIC_CHARSET:        cp = CP_1256;           break;
    case BALTIC_CHARSET:        cp = CP_1257;           break;                                    
    case VIETNAMESE_CHARSET:    cp = CP_1258;           break;
    case THAI_CHARSET:          cp = CP_THAI;           break;
    }

    return cp;
}

const WCHAR g_achLatin1MappingInUnicodeControlArea[32] =
{
    0x20ac, // 0x80
    0x0081, // 0x81
    0x201a, // 0x82
    0x0192, // 0x83
    0x201e, // 0x84
    0x2026, // 0x85
    0x2020, // 0x86
    0x2021, // 0x87
    0x02c6, // 0x88
    0x2030, // 0x89
    0x0160, // 0x8a
    0x2039, // 0x8b
    0x0152, // 0x8c <min>
    0x008d, // 0x8d
    0x017d, // 0x8e
    0x008f, // 0x8f
    0x0090, // 0x90
    0x2018, // 0x91
    0x2019, // 0x92
    0x201c, // 0x93
    0x201d, // 0x94
    0x2022, // 0x95
    0x2013, // 0x96
    0x2014, // 0x97
    0x02dc, // 0x98
    0x2122, // 0x99 <max>
    0x0161, // 0x9a
    0x203a, // 0x9b
    0x0153, // 0x9c
    0x009d, // 0x9d
    0x017e, // 0x9e
    0x0178  // 0x9f
};