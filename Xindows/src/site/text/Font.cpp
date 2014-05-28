/*
*  @doc    INTERNAL
*
*  @module FONT.C -- font cache |
*
*      includes font cache, char width cache;
*      create logical font if not in cache, look up
*      character widths on an as needed basis (this
*      has been abstracted away into a separate class
*      so that different char width caching algos can
*      be tried.) <nl>
*
*  Owner: <nl>
*      David R. Fulmer <nl>
*      Christian Fortini <nl>
*      Jon Matousek <nl>
*
*  History: <nl>
*      7/26/95     jonmat  cleanup and reorganization, factored out
*                  char width caching code into a separate class.
*
*  Copyright (c) 1995-1996 Microsoft Corporation. All rights reserved.
*/

#include "stdafx.h"
#include "Font.h"

#include "txtdefs.h"
#include "intl.h"
#include "fontlnk.h"

#ifdef _DEBUG
const int cMapSize = 1024;
class CDCToHFONT
{
    HDC   _dc[cMapSize];
    HFONT _font[cMapSize];
public:
    CDCToHFONT()
    {
        memset(_dc, 0, cMapSize*sizeof(HDC));
        memset(_font, 0, cMapSize*sizeof(HFONT));
    };
    void Assign(HDC hdc, HFONT hfont)
    {
        int freePos = -1;
        for(int i=0; i<cMapSize; i++)
        {
            if(_dc[i]==hdc)
            {
                _font[i] = hfont;
                return;
            }
            else if(freePos==-1 && _dc[i]==0)
            {
                freePos = i;
            }
        }
        Assert(freePos != -1);
        _dc[freePos] = hdc;
        _font[freePos] = hfont;
    };
    int FindNextFont(int posStart, HFONT hfont)
    {
        for(int i=posStart; i<cMapSize; i++)
        {
            if(_font[i] == hfont)
            {
                return i;
            }
        }
        return -1;
    };
    void Erase(int pos)
    {
        _dc[pos] = 0;
        _font[pos] = 0;
    };
    HFONT Font(int pos)
    {
        return _font[pos];
    };
    HDC DC(int pos)
    {
        return _dc[pos];
    };
};

CDCToHFONT mDc2Font;

HFONT SelectFontEx(HDC hdc, HFONT hfont)
{
#ifndef _WIN64
    //$ Win64: GetObjectType is returning zero on Axp64
    Assert(OBJ_FONT == GetObjectType((HGDIOBJ)hfont));
#endif
    mDc2Font.Assign(hdc, hfont);
    return SelectFont(hdc, hfont);
}

BOOL DeleteFontEx(HFONT hfont)
{
    int pos = -1;
    while(-1 != (pos=mDc2Font.FindNextFont(pos+1, hfont))) 
    {
        Trace0("##### Attempt to delete selected font\n");
        SelectObject(mDc2Font.DC(pos), GetStockObject(SYSTEM_FONT));
        mDc2Font.Erase(pos);
    }
    return DeleteObject((HGDIOBJ)hfont);
}
#endif //_DEBUG

#define CLIP_DFA_OVERRIDE   0x40    // Used to disable Korea & Taiwan font association
#define yHeightCharMost     32760   // corresponds to yHeightCharPtsMost in textedit.h

static TCHAR lfScriptFaceName[LF_FACESIZE] = _T("Script");
const TCHAR* Arial() { return _T("Arial"); }
const TCHAR* TimesNewRoman() { return _T("Times New Roman"); }

CFontCache g_FontCache;

static SCRIPT_IDS sidsSystem = sidsNotSet;

// U+0000 -> U+4DFF     All Latin and other phonetics.
// U+4E00 -> U+ABFF     CJK Ideographics
// U+AC00 -> U+FFFF     Korean+, as Korean ends at U+D7A3

// For 2 caches at CJK Ideo split, max cache sizes {256, 512} that give us a
// respective collision rate of <4% and <22%, and overall rate of <8%.
// Stats based on a 300K char Japanese text file.
INT maxCacheSize[TOTALCACHES] = { 255, 511, 511 };

#define IsZeroWidth(ch) ((ch)>=0x0300 && IsZeroWidthChar((ch)))

#ifdef _DEBUG
LONG CBaseCcs::s_cMaxCccs = 0;
LONG CBaseCcs::s_cTotalCccs = 0;
#endif //_DEBUG

//+-----------------------------------------------------------------------
//
//  Function:   DeinitUniscribe
//
//  Synopsis:   Clears script caches in USP.DLL private heap. This will
//              permit a clean shut down of USP.DLL (Uniscribe).
//
//------------------------------------------------------------------------
void DeinitUniscribe()
{
    g_FontCache.FreeScriptCaches();
}

/*
*  CFontCache & fc()
*
*  @func
*      return the global font cache.
*  @comm
*      current #defined to store 16 logical fonts and
*      respective character widths.
*
*/
CFontCache& fc()
{
    return g_FontCache;
}

// ===================================  CFontCache  ====================================
void InitFontCache()
{
    g_FontCache.Init();
}

void DeInitFontCache()
{
    g_FontCache.DeInit();
}

// Return our small face name cache to a zero state.
void ClearFaceCache()
{
    EnterCriticalSection(&g_FontCache._csFaceCache);
    g_FontCache._iCacheLen = 0;
    g_FontCache._iCacheNext = 0;
    LeaveCriticalSection(&g_FontCache._csFaceCache);
}

/*
*  CFontCache::DeInit()
*
*  @mfunc
*      This is allocated as an extern. Don't let the
*      CRT determine when we get dealloc'ed.
*/
void CFontCache::DeInit()
{
    DWORD i;
    for(i=0; i<cccsMost; i++)
    {
        if(_rpBaseCcs[i])
        {
            _rpBaseCcs[i]->NullOutScriptCache();
            _rpBaseCcs[i]->PrivateRelease();
        }
    }

    DeleteCriticalSection(&_cs);
    DeleteCriticalSection(&_csOther);
    DeleteCriticalSection(&_csFaceCache);
    DeleteCriticalSection(&_csFaceNames);

    _atFontInfo.Free(); // Maybe we should do this in a critical section, but naw...

#ifdef _DEBUG
    Trace0("Font Metrics:");
    Trace1("\tSize of FontCache:%ld", sizeof(CFontCache));
    Trace1("\tSize of Cccs:%ld + a min of 1024 bytes", sizeof(CCcs));
    Trace1("\tMax no. of fonts allocated:%ld", CBaseCcs::s_cMaxCccs);
    Trace1("\tNo. of fonts replaced: %ld\n", _cCccsReplaced);
#endif
}

//+----------------------------------------------------------------------------
//
//  Function:   CFontCache::EnsureScriptIDsForFont
//
//  Synopsis:   When we add a new facename to our _atFontInfo cache, we
//              defer the calculation of the script IDs (sids) for this
//              face.  An undetermined sids has the value of sidsNotSet.
//              Inside of CBaseCcs::MakeFont, we need to make sure that the
//              script IDs are computed, as we will need this information
//              to fontlink properly.
//
//  Returns:    SCRIPT_IDS.  If an error occurs, we return sidsAll, which
//              effectively disables fontlinking for this font.
//
//-----------------------------------------------------------------------------
SCRIPT_IDS CFontCache::EnsureScriptIDsForFont(HDC hdc, CBaseCcs* pBaseCcs, BOOL fDownloadedFont)
{
    const LONG latmFontInfo = pBaseCcs->_latmLFFaceName;

    if(latmFontInfo)
    {
        CFontInfo* pfi;
        HRESULT hr = _atFontInfo.GetInfoFromAtom(latmFontInfo-1, &pfi);

        if(!hr)
        {
            if(pfi->_sids == sidsNotSet)
            {
                if(!fDownloadedFont)
                {
                    pfi->_sids = ScriptIDsFromFont(hdc, pBaseCcs->_hfont, pBaseCcs->_sPitchAndFamily&TMPF_TRUETYPE);

                    if(pBaseCcs->_fLatin1CoverageSuspicious)
                    {
                        pfi->_sids &= ~(ScriptBit(sidLatin));
                    }
                }
                else
                {
                    pfi->_sids = sidsAll; // don't fontlink for embedded fonts
                }
            }

            return pfi->_sids;
        }
    }
    else
    {
        if(sidsSystem == sidsNotSet)
        {
            sidsSystem = ScriptIDsFromFont(hdc, pBaseCcs->_hfont, pBaseCcs->_sPitchAndFamily&TMPF_TRUETYPE);
        }

        return sidsSystem;
    }

    return sidsAll;
}

//+----------------------------------------------------------------------------
//
//  Function:   CFontCache::ScriptIDsForAtom
//
//  Returns:    Return the SCRIPT_IDS associated with the facename-base atom.
//
//-----------------------------------------------------------------------------
SCRIPT_IDS CFontCache::ScriptIDsForAtom(LONG latmFontInfo)
{
    Assert(latmFontInfo);

    CFontInfo* pfi;
    HRESULT hr = _atFontInfo.GetInfoFromAtom(latmFontInfo-1, &pfi);

    if(!hr)
    {
        Assert(pfi->_sids == sidsNotSet);

        return pfi->_sids;
    }

    return sidsAll;
}

LONG CFontCache::GetAtomWingdings()
{
    if(!_latmWingdings)
    {
        _latmWingdings = GetAtomFromFaceName(_T("Wingdings"));
        _atFontInfo.SetScriptIDsOnAtom(_latmWingdings-1, sidsAll);
    }
    return _latmWingdings;
}

HRESULT CFontCache::AddScriptIdToAtom(LONG latmFontInfo, SCRIPT_ID sid)
{
    RRETURN(_atFontInfo.AddScriptIDOnAtom(latmFontInfo-1, sid));
}

/*
*  CFontCache::FreeScriptCaches ()
*
*  @mfunc
*      frees SCRIPT_CACHE memory allocated in USP.DLL
*/
void CFontCache::FreeScriptCaches()
{
    DWORD i;
    for(i=0; i<cccsMost; i++)
    {
        if(_rpBaseCcs[i])
        {
            _rpBaseCcs[i]->ReleaseScriptCache();
        }
    }
}

/*
*  CCcs* CFontCache::GetCcs(hdc, pcf, lZoomNumerator, lZoomDenominator, yPixelsPerInch)
*
*  @mfunc
*      Search the font cache for a matching logical font and return it.
*      If a match is not found in the cache,  create one.
*
*  @rdesc
*      A logical font matching the given CHARFORMAT info.
*/
CCcs* CFontCache::GetCcs(HDC hdc, CDocInfo* pdci, const CCharFormat* const pcf)
{
    CCcs* pccs = new CCcs();

    if(pccs)
    {
        pccs->_hdc = hdc;
        pccs->_pBaseCcs = GetBaseCcs(hdc, pdci, pcf, NULL);
        if(!pccs->_pBaseCcs)
        {
            delete pccs;
            pccs = NULL;
        }
    }

    return pccs;
}

CCcs* CFontCache::GetFontLinkCcs(HDC hdc, CDocInfo* pdci, CCcs* pccsBase, const CCharFormat* const pcf)
{
    CCcs* pccs;

#ifdef _DEBUG
    // Allow disabling the height-adjusting feature of fontlinking through tags
    return GetCcs(hdc, pdci, pcf);
#endif

    pccs = new CCcs();
    if(pccs)
    {
        CBaseCcs* pBaseBaseCcs = pccsBase->_pBaseCcs;

        pBaseBaseCcs->AddRef();

        pccs->_hdc = hdc;
        pccs->_pBaseCcs = GetBaseCcs(hdc, pdci, pcf, pBaseBaseCcs);
        if(!pccs->_pBaseCcs)
        {
            delete pccs;
            pccs = NULL;
            goto Cleanup;
        }

        Trace2("GetFontLinkCcs %S based on %S\n",
            fc().GetFaceNameFromAtom(pcf->_latmFaceName),
            fc().GetFaceNameFromAtom(pBaseBaseCcs->_latmLFFaceName));

        pBaseBaseCcs->PrivateRelease();
    }
    else
    {
        pccs = pccsBase;
    }

Cleanup:

    return pccs;
}

CBaseCcs* CFontCache::GetBaseCcs(
         HDC hdc,
         CDocInfo* pdci,
         const CCharFormat* const pcf,  //@parm description of desired logical font
         CBaseCcs* pBaseBaseCcs)        //@parm facename from which we're fontlinking
{
    CBaseCcs*   pBaseCcs = NULL;
    LONG        lfHeight;
    int         i;
    BYTE        bCrc;
    SHORT       hashKey;
    CBaseCcs::CompareArgs cargs;
    BOOL (CBaseCcs::*CompareFunc)(CBaseCcs::CompareArgs*);

    // Duplicate the format structure because we might need to change some of the
    // values by the zoom factor
    // and in the case of sub superscript
    CCharFormat cf = *pcf;

    // FUTURE igorzv
    // Take subscript size, subscript offset, superscript offset, superscript size
    // from the OUTLINETEXMETRIC
    lfHeight = -cf.GetHeightInPixels(pdci);

    // change cf._yHeight in the case of sub superscript
    if(cf._fSuperscript || cf._fSubscript)
    {
        LONG yHeight = cf.GetHeightInTwips(pdci->_pDoc);

        if(cf._fSubscript)
        {
            cf._yOffset -= yHeight / 2;
        }
        else
        {
            cf._yOffset += yHeight / 2;
        }

        cf._bCrcFont = cf.ComputeFontCrc();

        // NOTE: lfHeight has already been scaled by 2/3 if *SCRIPT set
    }

    bCrc = cf._bCrcFont;

    Assert(bCrc == cf.ComputeFontCrc());

    if(!lfHeight)
    {
        lfHeight--; // round error, make this a minimum legal height of -1.
    }

    cargs.pcf = &cf;
    cargs.lfHeight = lfHeight;

    if(pBaseBaseCcs)
    {
        cargs.latmBaseFaceName = pBaseBaseCcs->_latmLFFaceName;
        CompareFunc = &CBaseCcs::CompareForFontLink;
    }
    else
    {
        cargs.latmBaseFaceName = pcf->_latmFaceName;
        CompareFunc = &CBaseCcs::Compare;
    }

    EnterCriticalSection(&_cs);

    // check our hash before going sequential.
    hashKey = bCrc & QUICKCRCSEARCHSIZE;
    if(bCrc == quickCrcSearch[hashKey].bCrc)
    {
        pBaseCcs = quickCrcSearch[hashKey].pBaseCcs;
        if(pBaseCcs && pBaseCcs->_bCrc==bCrc)
        {
            if((pBaseCcs->*CompareFunc)(&cargs))
            {
                goto matched;
            }
        }
        pBaseCcs = NULL;
    }
    quickCrcSearch[hashKey].bCrc = bCrc;

    // squentially search ccs for same character format
    for(i=0; i<cccsMost; i++)
    {
        CBaseCcs* pBaseCcsTemp = _rpBaseCcs[i];

        if(pBaseCcsTemp && pBaseCcsTemp->_bCrc==bCrc)
        {
            if((pBaseCcsTemp->*CompareFunc)(&cargs))
            {
                pBaseCcs = pBaseCcsTemp;
                break;
            }
        }
    }

matched:
    if(!pBaseCcs)
    {
        // we did not find a match, init a new font cache.
        pBaseCcs = GrabInitNewBaseCcs(hdc, &cf, pdci, cargs.latmBaseFaceName);

        if(pBaseCcs && pBaseBaseCcs && !pBaseCcs->_fHeightAdjustedForFontlinking)
        {
            pBaseCcs->FixupForFontLink(hdc, pBaseBaseCcs);
        }
    }
    else
    {
        if(pBaseCcs->_dwAge != _dwAgeNext-1)
        {
            pBaseCcs->_dwAge = _dwAgeNext++;
        }

        // the same font can be used at different offsets.
        pBaseCcs->_yOffset = cf._yOffset ? pdci->DyzFromDyt(pdci->DytFromTwp(cf._yOffset)) : 0;

#ifdef _DEBUG
        long dyzOld = MulDivQuick(cf._yOffset, pdci->_sizeInch.cy, LY_PER_INCH);
        Assert(pBaseCcs->_yOffset==dyzOld || pdci->IsZoomed());
#endif //_DEBUG
    }

    if(pBaseCcs)
    {
        quickCrcSearch[hashKey].pBaseCcs = pBaseCcs;

        // AddRef the entry being returned
        pBaseCcs->AddRef();

        pBaseCcs->EnsureFastCacheExists(hdc);
    }

    LeaveCriticalSection(&_cs);

    return pBaseCcs;
}

// Checks to see if this face name is in the atom table.
// if not, puts it in.  We report externally 1 higher than the
// actualy atom table values, so that we can reserve latom==0
// to the error case, or a blank string.
LONG CFontCache::GetAtomFromFaceName(const TCHAR* szFaceName)
{
    HRESULT hr;
    LONG lAtom = 0;

    // If they pass in the NULL string, when they ask for it out again,
    // they're gonna get a blank string, which is different.
    Assert(szFaceName);

    if(szFaceName && *szFaceName)
    {
        EnterCriticalSection(&_csFaceNames);
        hr = _atFontInfo.GetAtomFromName(szFaceName, &lAtom);
        if(hr)
        {
            // String not in there.  Put it in.
            // Note we defer the calculation of the SCRIPT_IDS.
            hr= _atFontInfo.AddInfoToAtomTable(szFaceName, &lAtom);
            AssertSz(hr==S_OK, "Failed to add Font Face Name to Atom Table");
        }
        if(hr == S_OK)
        {
            lAtom++;
        }
        LeaveCriticalSection(&_csFaceNames);
    }
    // else input was NULL or empty string.  Return latom=0.

    return lAtom;
}

// Checks to see if this face name is in the atom table.
LONG CFontCache::FindAtomFromFaceName(const TCHAR* szFaceName)
{
    HRESULT hr;
    LONG lAtom = 0;

    // If they pass in the NULL string, when they ask for it out again,
    // they're gonna get a blank string, which is different.
    Assert(szFaceName);

    if(szFaceName && *szFaceName)
    {
        EnterCriticalSection(&_csFaceNames);
        hr = _atFontInfo.GetAtomFromName(szFaceName, &lAtom);
        if(hr == S_OK)
        {
            lAtom++;
        }
        LeaveCriticalSection(&_csFaceNames);
    }

    // lAtom of zero means that it wasn't found or that the input was bad.
    return lAtom;
}

const TCHAR* CFontCache::GetFaceNameFromAtom(LONG latmFaceName)
{
    const TCHAR* szReturn = _afxGlobalData._Zero.ach;

    if(latmFaceName)
    {
        HRESULT hr;
        CFontInfo* pfi;

        EnterCriticalSection(&_csFaceNames);
        hr = _atFontInfo.GetInfoFromAtom(latmFaceName-1, &pfi);
        LeaveCriticalSection(&_csFaceNames);
        Assert(!hr);
        Assert(pfi->_cstrFaceName.Length() < LF_FACESIZE);
        szReturn = pfi->_cstrFaceName;
    }
    return szReturn;
}

void CBaseCcs::NullOutScriptCache()
{
    Assert(_sc == NULL);

    // A safety valve so we don't crash.
    if(_sc)
    {
        _sc = NULL;
    }
}

void CBaseCcs::ReleaseScriptCache()
{
    // Free the script cache
    if (_sc)
    {
        ::ScriptFreeCache(&_sc);
        // NB (mikejoch) If ScriptFreeCache() fails then there is no way to
        // free the cache, so we'll end up leaking it. This shouldn't ever
        // happen since the only way for _sc to be non- NULL is via some other
        // USP function succeeding.
    }
    Assert(_sc == NULL);
}

void CBaseCcs::PrivateRelease()
{
    if(InterlockedDecrement((LONG*)&_dwRefCount) == 0)
    {
        delete this;
    }
}

/*
*  CBaseCcs* CFontCache::GrabInitNewBaseCcs(hdc, pcf)
*
*  @mfunc
*      create a logical font and store it in our cache.
*
*/
CBaseCcs* CFontCache::GrabInitNewBaseCcs(
    HDC hdc,                        //@parm HDC into which font will be selected
    const CCharFormat* const pcf,   //@parm description of desired logical font
    CDocInfo* pdci,                 //@parm Y Pixels per Inch
    LONG latmBaseFaceName)
{
    int     i;
    int     iReplace = -1;
    int     iOldest = -1;
    DWORD   dwAgeReplace = 0xffffffff;
    DWORD   dwAgeOldest = 0xffffffff;
    CBaseCcs* pBaseCcsNew = new CBaseCcs();

    // Initialize new CBaseCcs
    if(!pBaseCcsNew || !pBaseCcsNew->Init(hdc, pcf, pdci, latmBaseFaceName))
    {
        if(pBaseCcsNew)
        {
            pBaseCcsNew->PrivateRelease();
        }

        AssertSz(FALSE, "CFontCache::GrabInitNewBaseCcs init of entry FAILED");
        return NULL;
    }

    MemSetName((pBaseCcsNew, "CBaseCcs F:%ls, H:%d, W:%d", pBaseCcsNew->_lf.lfFaceName, -pBaseCcsNew->_lf.lfHeight, pBaseCcsNew->_lf.lfWeight));

    // look for unused entry and oldest in use entry
    for(i=0; i<cccsMost&&_rpBaseCcs[i]; i++)
    {
        CBaseCcs* pBaseCcs = _rpBaseCcs[i];
        if(pBaseCcs->_dwAge < dwAgeOldest)
        {
            iOldest = i;
            dwAgeOldest = pBaseCcs->_dwAge;
        }
        if(pBaseCcs->_dwRefCount == 1)
        {
            if(pBaseCcs->_dwAge < dwAgeReplace)
            {
                iReplace  = i;
                dwAgeReplace = pBaseCcs->_dwAge;
            }
        }
    }

    if(i == cccsMost) // Didn't find an unused entry, use oldest entry
    {
        int hashKey;
        // if we didn't find a replacement, replace the oldest
        if(iReplace == -1)
        {
            Assert(iOldest != -1);
            iReplace = iOldest;
        }

#ifdef _DEBUG
        _cCccsReplaced++;
#endif //_DEBUG

        hashKey = _rpBaseCcs[iReplace]->_bCrc & QUICKCRCSEARCHSIZE;
        if(quickCrcSearch[hashKey].pBaseCcs == _rpBaseCcs[iReplace])
        {
            quickCrcSearch[hashKey].pBaseCcs = NULL;
        }

        Trace((_T("Releasing font(F:%ls, H:%d, W:%d) from slot %d\n"),
            _rpBaseCcs[iReplace]->_lf.lfFaceName,
            -_rpBaseCcs[iReplace]->_lf.lfHeight,
            _rpBaseCcs[iReplace]->_lf.lfWeight,
            iReplace));

        _rpBaseCcs[iReplace]->PrivateRelease();
        i = iReplace;
    }

    Trace((_T("Storing font(F:%ls, H:%d, W:%d) in slot %d\n"),
        pBaseCcsNew->_lf.lfFaceName,
        -pBaseCcsNew->_lf.lfHeight,
        pBaseCcsNew->_lf.lfWeight,
        i));

    _rpBaseCcs[i] = pBaseCcsNew;

    return pBaseCcsNew;
}

//+----------------------------------------------------------------------------
//
//  Function:   CFontCache::SetHanCodePageInfo()
//
//  Synopsis:   Determine the FE font converage of the system.  This
//              information will used to choose between FE fonts when we have
//              a text run of ambiguous Han characters.
//
//  Returns:    A dword of FS_* bits
//
//-----------------------------------------------------------------------------
const BYTE s_abCharSet[] =
{ 
    JOHAB_CHARSET,       // FS_JOHAB                0x00200000L
    CHINESEBIG5_CHARSET, // FS_CHINESETRAD          0x00100000L
    HANGEUL_CHARSET,     // FS_WANSUNG              0x00080000L
    GB2312_CHARSET,      // FS_CHINESESIMP          0x00040000L
    SHIFTJIS_CHARSET,    // FS_JISJAPAN             0x00020000L
    VIETNAMESE_CHARSET,  // FS_VIETNAMESE           0x00000100L
    BALTIC_CHARSET,      // FS_BALTIC               0x00000080L
    ARABIC_CHARSET,      // FS_ARABIC               0x00000040L
    HEBREW_CHARSET,      // FS_HEBREW               0x00000020L
    TURKISH_CHARSET,     // FS_TURKISH              0x00000010L
    GREEK_CHARSET,       // FS_GREEK                0x00000008L
    RUSSIAN_CHARSET,     // FS_CYRILLIC             0x00000004L
    EASTEUROPE_CHARSET,  // FS_LATIN2               0x00000002L
};

const DWORD s_adwFontScripts[] =
{
    FS_JOHAB,            // JOHAB_CHARSET
    FS_CHINESETRAD,      // CHINESEBIG5_CHARSET
    FS_WANSUNG,          // HANGEUL_CHARSET
    FS_CHINESESIMP,      // GB2312_CHARSET
    FS_JISJAPAN,         // SHIFTJIS_CHARSET
    FS_VIETNAMESE,       // VIETNAMESE_CHARSET
    FS_BALTIC,           // BALTIC_CHARSET
    FS_ARABIC,           // ARABIC_CHARSET
    FS_HEBREW,           // HEBREW_CHARSET
    FS_TURKISH,          // TURKISH_CHARSET
    FS_GREEK,            // GREEK_CHARSET
    FS_CYRILLIC,         // RUSSIAN_CHARSET
    FS_LATIN2,           // EASTEUROPE_CHARSET
};

int FAR PASCAL CALLBACK SetSupportedCodePageInfoEnumFontCallback(
    const LOGFONT FAR* lplf,
    const TEXTMETRIC FAR* lptm,
    DWORD FontType,
    LPARAM lParam)
{
    *((BOOL*)lParam) = TRUE; // fFound;

    return 0;
}

DWORD CFontCache::SetSupportedCodePageInfo(HDC hdc)
{
    DWORD dwSupportedCodePageInfo = 0;

    LOGFONT lf;
    BOOL fFound;
    int i;

    lf.lfFaceName[0] = 0;
    lf.lfPitchAndFamily = 0;

    for(i=ARRAYSIZE(s_abCharSet); i;)
    {
        lf.lfCharSet = s_abCharSet[--i];

        fFound = FALSE;

        EnumFontFamiliesEx(hdc, &lf, SetSupportedCodePageInfoEnumFontCallback, LPARAM(&fFound), 0);

        if(fFound)
        {
            dwSupportedCodePageInfo |= s_adwFontScripts[i];
        }
    }

    _dwSupportedCodePageInfo = dwSupportedCodePageInfo;

    return dwSupportedCodePageInfo;
}

// =============================  CBaseCcs  class  ===================================================
/*
*  BOOL CBaseCcs::Init()
*
*  @mfunc
*      Init one font cache object. The global font cache stores
*      individual CBaseCcs objects.
*/
BOOL CBaseCcs::Init(
        HDC hdc,                        //@parm HDC into which font will be selected
        const CCharFormat* const pcf,   //@parm description of desired logical font
        CDocInfo* pdci,                 //@parm Y pixels per inch
        LONG latmBaseFaceName)
{
    _sc = NULL; // Initialize script cache to NULL - COMPLEXSCRIPT

    _bConvertMode = CM_NONE;

    if(MakeFont(hdc, pcf, pdci))
    {
        _bCrc = pcf->_bCrcFont;
        _yCfHeight = pcf->_yHeight;
        _latmBaseFaceName = latmBaseFaceName;
        _fHeightAdjustedForFontlinking = FALSE;

        // BUGBUG (cthrash) This needs to be removed.  We used to calculate
        // this all the time, now we only calculate it on an as-needed basis,
        // which means at intrinsics fontlinking time.
        _dwLangBits = 0;

        // the same font can be used at different offsets.
        _yOffset =  pcf->_yOffset ? pdci->DyzFromDyt(pdci->DytFromTwp(pcf->_yOffset)) : 0;

#ifdef _DEBUG
        long dyzOld = MulDivQuick(pcf->_yOffset, pdci->_sizeInch.cy, LY_PER_INCH);
        Assert(_yOffset==dyzOld || pdci->IsZoomed());
#endif // DBG==1

        _dwAge = fc()._dwAgeNext++;

        return TRUE; // successfully created a new font cache.
    }

    return FALSE;
}

/*
*  BOOL CBaseCcs::MakeFont(hdc, pcf)
*
*  @mfunc
*      Wrapper, setup for CreateFontIndirect() to create the font to be
*      selected into the HDC.
*
*  @rdesc
*      TRUE if OK, FALSE if allocation failure
*/
CODEPAGE DefaultCodePageFromCharSet(BYTE, UINT);

BOOL CBaseCcs::MakeFont(
        HDC hdc,                        //@parm HDC into which  font will be selected
        const CCharFormat* const pcf,   //@parm description of desired logical font
        CDocInfo* pdci)
{
    BOOL fTweakedCharSet = FALSE;
    HFONT hfontOriginalCharset = NULL;
    TCHAR pszNewFaceName[LF_FACESIZE];
    const CODEPAGE cpDoc = pdci->_pDoc->GetFamilyCodePage();
    const LCID lcid = pcf->_lcid;
    BOOL fRes;
    LONG lfHeight;

    // We need the _sCodePage in case we need to call ExtTextOutA rather than ExtTextOutW.  
    _sCodePage = (USHORT)DefaultCodePageFromCharSet(pcf->_bCharSet, cpDoc);

    // Computes font height
    AssertSz(pcf->_yHeight<=INT_MAX, "It's too big");

    //  Roundoff can result in a height of zero, which is bad.
    //  If this happens, use the minimum legal height.
    lfHeight = -(const_cast<CCharFormat* const>(pcf)->GetHeightInPixels(pdci));
    if(lfHeight > 0)
    {
        lfHeight = -lfHeight;       // TODO: do something more intelligent...
    }
    else if(!lfHeight)
    {
        lfHeight--;                 // round error, make this a minimum legal height of -1.
    }
    _lf.lfHeight = _yOriginalHeight = lfHeight;
    _lf.lfWidth  = 0;

    if(pcf->_wWeight != 0)
    {
        _lf.lfWeight = pcf->_wWeight;
    }
    else
    {
        _lf.lfWeight    = pcf->_fBold ? FW_BOLD : FW_NORMAL;
    }
    _lf.lfItalic        = pcf->_fItalic;
    _lf.lfUnderline     = 0;
    _lf.lfStrikeOut     = 0;
    _lf.lfCharSet       = pcf->_bCharSet;
    _lf.lfEscapement    = 0;
    _lf.lfOrientation   = 0;
    _lf.lfOutPrecision  = OUT_DEFAULT_PRECIS;
    _lf.lfQuality       = DEFAULT_QUALITY;
    _lf.lfPitchAndFamily = _bPitchAndFamily = pcf->_bPitchAndFamily;
    _lf.lfClipPrecision = CLIP_DEFAULT_PRECIS|CLIP_DFA_OVERRIDE;

    // HACK (cthrash) Don't pick a non-TT font for generic font types 'serif'
    // or 'sans-serif.'  More precisely, prevent MS Sans Serif and MS Serif
    // from getting selected.
    if(pcf->_latmFaceName==0 && (pcf->_bPitchAndFamily&(FF_ROMAN|FF_SWISS)))
    {
        _lf.lfOutPrecision |= OUT_TT_ONLY_PRECIS;
    }

    // Only use CLIP_EMBEDDED when necessary, or Win95 will make you pay.
    if(pcf->_fDownloadedFont)
    {
        _lf.lfClipPrecision |= CLIP_EMBEDDED;
    }

    // Note (paulpark): We must be careful with _lf.  It is important that
    // _latmLFFaceName be kept in sync with _lf.lfFaceName.  The way to do that
    // is to use the SetLFFaceName function.
    SetLFFaceNameAtm(pcf->_latmFaceName);

    fRes = GetFontWithMetrics(hdc, pszNewFaceName, cpDoc, lcid);

    if((_sPitchAndFamily&0x00F0)!=pcf->_bPitchAndFamily
        && (pcf->_bPitchAndFamily==FF_SCRIPT || pcf->_bPitchAndFamily==FF_DECORATIVE))
    {
        // If we ask for FF_SCRIPT or FF_DECORATIVE, we sometimes get
        // a font which has neither of those characteristics (since
        // the font matcher does not map anything in OEM_CHARSET to
        // DEFAULT_CHARSET).  In this case use our default face names.
        _lf.lfCharSet = DEFAULT_CHARSET; // accept all charsets

        if(pcf->_bPitchAndFamily == FF_SCRIPT)
        {
            SetLFFaceName(lfScriptFaceName);
        }
        else
        {
            SetLFFaceName(Arial()); // BUGBUG (what should this be)
        }
    }

    if(!_tcsiequal(pszNewFaceName, _lf.lfFaceName))
    {
        BOOL fCorrectFont = FALSE;

        if(_bCharSet == SYMBOL_CHARSET)
        {
            // #1. if the face changed, and the specified charset was SYMBOL,
            //     but the face name exists and suports ANSI, we give preference
            //     to the face name
            _lf.lfCharSet = ANSI_CHARSET;
            fTweakedCharSet = TRUE;

            hfontOriginalCharset = _hfont;
            GetFontWithMetrics(hdc, pszNewFaceName, cpDoc, lcid);

            if(_tcsiequal(pszNewFaceName, _lf.lfFaceName))
            {
                // that's right, ANSI is the asnwer
                fCorrectFont = TRUE;
            }
            else
            {
                // no, fall back by default
                // the charset we got was right
                _lf.lfCharSet = pcf->_bCharSet;
            }
        }
        else if(_lf.lfCharSet==DEFAULT_CHARSET && _bCharSet==DEFAULT_CHARSET)
        {
            // #2. If we got the "default" font back, we don't know what it
            // means (could be anything) so we verify that this guy's not SYMBOL
            // (symbol is never default, but the OS could be lying to us!!!)
            // we would like to verify more like whether it actually gave us
            // Japanese instead of ANSI and labeled it "default"...
            // but SYMBOL is the least we can do
            _lf.lfCharSet = SYMBOL_CHARSET;
            SetLFFaceName(pszNewFaceName);
            fTweakedCharSet = TRUE;

            hfontOriginalCharset = _hfont;
            GetFontWithMetrics(hdc, pszNewFaceName, cpDoc, lcid);

            if(_tcsiequal(pszNewFaceName, _lf.lfFaceName))
            {
                // that's right, it IS symbol!
                // 'correct' the font to the 'true' one,
                // and we'll get fMappedToSymbol
                fCorrectFont = TRUE;
            }

            // always restore the charset name, we didn't want to
            // question he original choice of charset here
            _lf.lfCharSet = pcf->_bCharSet;
        }
        else if(_bConvertMode!=CVM_SYMBOL && IsFECharset(_lf.lfCharSet))
        {
            // NOTE (cthrash) _lf.lfCharSet is what we asked for, _bCharset is what we got.
            if(_lf.lfCharSet==GB2312_CHARSET && _lf.lfPitchAndFamily|FIXED_PITCH)
            {
                // HACK (cthrash) On vanilla PRC systems, you will not be able to ask
                // for a fixed-pitch font which covers the non-GB2312 GBK codepoints.
                // We come here if we asked for a fixed-pitch PRC font but failed to 
                // get a facename match.  So we try again, only without FIXED_PITCH
                // set.  The side-effect is that CBaseCcs::Compare needs to compare
                // against the original _bPitchAndFamily else it will fail every time.
                _lf.lfPitchAndFamily = _lf.lfPitchAndFamily ^ FIXED_PITCH;
            }
            else
            {
                // this is a FE Font (from Lang pack) on a nonFEsystem
                SetLFFaceName(pszNewFaceName);
            }

            hfontOriginalCharset = _hfont;

            GetFontWithMetrics(hdc, pszNewFaceName, cpDoc, lcid);

            if(_tcsiequal(pszNewFaceName, _lf.lfFaceName))
            {
                // that's right, it IS the FE font we want!
                // 'correct' the font to the 'true' one.
                fCorrectFont = TRUE;
            }

            fTweakedCharSet = TRUE;
        }

        if(hfontOriginalCharset)
        {
            // either keep the old font or the new one
            if(fCorrectFont)
            {
                DeleteFontEx(hfontOriginalCharset);
                hfontOriginalCharset = NULL;
            }
            else
            {
                // fall back to the original font
                DeleteFontEx(_hfont);

                _hfont = hfontOriginalCharset;
                hfontOriginalCharset = NULL;

                GetTextMetrics(hdc, cpDoc, lcid);
            }
        }
    }

RetryCreateFont:
    if(!pcf->_fDownloadedFont)
    {
        // could be that we just plain symply get mapped to symbol.
        // avoid it
        BOOL fMappedToSymbol = (_bCharSet==SYMBOL_CHARSET && _lf.lfCharSet!=SYMBOL_CHARSET);
        BOOL fChangedCharset = (_bCharSet!=_lf.lfCharSet && _lf.lfCharSet!=DEFAULT_CHARSET);

        if(fChangedCharset || fMappedToSymbol)
        {
            const TCHAR* pchFallbackFaceName =
                (pcf->_bPitchAndFamily&FF_ROMAN) ? TimesNewRoman() : Arial();

            // Here, the system did not preserve the font language or mapped
            // our non-symbol font onto a symbol font,
            // which will look awful when displayed.
            // Giving us a symbol font when we asked for a non-symbol one
            // (default can never be symbol) is very bizzare and means
            // that either the font name is not known or the system
            // has gone complete nuts here.
            // The charset language takes priority over the font name.
            // Hence, I would argue that nothing can be done to save the
            // situation at this point, and we have to
            // delete the font name and retry

            // let's tweak it a bit
            fTweakedCharSet = TRUE;

            if(_tcsiequal(_lf.lfFaceName, pchFallbackFaceName))
            {
                // we've been here already
                // no font with an appropriate charset is on the system
                // try getting the ANSI one for the original font name
                // next time around, we'll null out the name as well!!
                if(_lf.lfCharSet == ANSI_CHARSET)
                {
                    Trace0("Asking for ANSI ARIAL and not getting it?!\n");

                    // those Win95 guys have definitely outbugged me
                    goto Cleanup;
                }

                DeleteFontEx(_hfont);
                _hfont = NULL;
                SetLFFaceNameAtm(pcf->_latmFaceName);
                _lf.lfCharSet = ANSI_CHARSET;
            }
            else
            {
                DeleteFontEx(_hfont);
                _hfont = NULL;
                SetLFFaceName(pchFallbackFaceName);
            }

            GetFontWithMetrics(hdc, pszNewFaceName, cpDoc, lcid);
            goto RetryCreateFont;
        }
    }

Cleanup:
    if(fTweakedCharSet || _bConvertMode==CVM_SYMBOL)
    {
        // we must preserve the original charset value, since it is used in Compare()
        _lf.lfCharSet = pcf->_bCharSet;
        SetLFFaceNameAtm(pcf->_latmFaceName);
    }

    if(hfontOriginalCharset)
    {
        DeleteFontEx(hfontOriginalCharset);
        hfontOriginalCharset = NULL;
    }

    // if we're really really stuck, just get the system font and hope for the best.
    if(_hfont == 0)
    {
        _hfont = (HFONT)GetStockObject(SYSTEM_FONT);
    }

    _fFEFontOnNonFEWin95 = FALSE;

    // Make sure we know what have script IDs computed for this font.  Cache this value
    // to avoid a lookup later.
    _sids = fc().EnsureScriptIDsForFont(hdc, this, pcf->_fDownloadedFont);
    _sids &= ScriptIDsFromCharSet(_bCharSet);

    Trace((_T("CCcs::MakeFont(facename=%S,charset=%d) returned %S(charset=%d)\n"),
        fc().GetFaceNameFromAtom(pcf->_latmFaceName),
        pcf->_bCharSet,
        fc().GetFaceNameFromAtom(_latmLFFaceName),
        _bCharSet));

    return _hfont!=0;
}

/*
*  BOOL CBaseCcs::GetFontWithMetrics (szNewFaceName)
*
*  @mfunc
*      Get metrics used by the measurer and renderer and the new face name.
*/
BOOL CBaseCcs::GetFontWithMetrics(HDC hdc, TCHAR* szNewFaceName, CODEPAGE cpDoc, LCID lcid)
{
    // we want to keep _lf untouched as it is used in Compare().
    _hfont = CreateFontIndirect(&_lf);
    if(_hfont)
    {
        // FUTURE (alexgo) if a font was not created then we may want to select
        //      a default one.
        //      If we do this, then BE SURE not to overwrite the values of _lf as
        //      it is used to match with a pcf in our Compare().
        //
        // get text metrics, in logical units, that are constant for this font,
        // regardless of the hdc in use.
        if(GetTextMetrics(hdc, cpDoc, lcid))
        {
            HFONT hfontOld = SelectFontEx(hdc, _hfont);
            GetTextFace(hdc, LF_FACESIZE, szNewFaceName);
            SelectFontEx(hdc, hfontOld);
        }
        else
        {
            szNewFaceName[0] = 0;
        }
    }

    return (_hfont!=NULL);
}

/*
*  BOOL CBaseCcs::GetTextMetrics()
*
*  @mfunc
*      Get metrics used by the measureer and renderer.
*
*  @comm
*      These are in logical coordinates which are dependent
*      on the mapping mode and font selected into the hdc.
*/
BOOL CBaseCcs::GetTextMetrics(HDC hdc, CODEPAGE cpDoc, LCID lcid)
{
    BOOL        fRes = TRUE;
    HFONT       hfontOld;
    TEXTMETRIC  tm;

    AssertSz(_hfont, "CBaseCcs::Fill - CBaseCcs has no font");
    if(GetCurrentObject(hdc, OBJ_FONT) != _hfont)
    {
        hfontOld = SelectFontEx(hdc, _hfont);

        if(!hfontOld)
        {
            fRes = FALSE;
            DestroyFont();
            goto Cleanup;
        }
    }
    else
    {
        hfontOld = 0;
    }

    if(!::GetTextMetrics(hdc, &tm))
    {
        DEBUG_ONLY(GetLastError());
        fRes = FALSE;
        DeleteFontEx(_hfont);
        _hfont = NULL;
        DestroyFont();
        goto Cleanup;
    }

    // if we didn't know the true codepage, determine this now.
    if(_lf.lfCharSet == DEFAULT_CHARSET)
    {
        // BUGBUG (cthrash) Remove this.  The _sCodePage computed by MakeFont should
        // be accurate enough.
        _sCodePage = (USHORT)DefaultCodePageFromCharSet(tm.tmCharSet, cpDoc, lcid);
    }

    // the metrics, in logical units, dependent on the map mode and font.
    _yHeight         = (SHORT)tm.tmHeight;
    _yDescent        = _yTextMetricDescent = (SHORT)tm.tmDescent;
    _xAveCharWidth   = (SHORT)tm.tmAveCharWidth;
    _xMaxCharWidth   = (SHORT)tm.tmMaxCharWidth;
    _xOverhangAdjust = (SHORT)tm.tmOverhang;
    _sPitchAndFamily = (SHORT)tm.tmPitchAndFamily;
    _bCharSet        = tm.tmCharSet;

    if(_bCharSet==SHIFTJIS_CHARSET || _bCharSet==CHINESEBIG5_CHARSET
        || _bCharSet==HANGEUL_CHARSET || _bCharSet==GB2312_CHARSET)
    {
        if(tm.tmExternalLeading == 0)
        {
            // Increase descent by 1/8 font height for FE fonts
            LONG delta = _yHeight / 8;
            _yDescent += (SHORT)delta;
            _yHeight  += (SHORT)delta;
        }
        else
        {
            // use the external leading
            _yDescent += (SHORT)tm.tmExternalLeading;
            _yHeight += (SHORT)tm.tmExternalLeading;
        }
    }

    // We don't support TRUE_TYPE font yet.
    // bug in windows95, synthesized italic?
    if( _lf.lfItalic && 0==tm.tmOverhang &&
        !(TMPF_TRUETYPE&tm.tmPitchAndFamily) &&
        !((TMPF_DEVICE&tm.tmPitchAndFamily) &&
        (TMPF_VECTOR&tm.tmPitchAndFamily)))
    {
        _xOverhangAdjust = (SHORT)(tm.tmMaxCharWidth >> 1);
    }

    _xOverhang = 0;
    _xUnderhang = 0;
    if(_lf.lfItalic)
    {
        _xOverhang =  SHORT((tm.tmAscent+1) >> 2);
        _xUnderhang =  SHORT((tm.tmDescent+1) >> 2);
    }

    // HACK (cthrash) Many Win3.1 vintage fonts (such as MS Sans Serif, Courier)
    // will claim to support all of Latin-1 when in fact it does not.  The hack
    // is to check the last character, and if the font claims that it U+2122
    // TRADEMARK, then we suspect the coverage is busted.
    _fLatin1CoverageSuspicious = !(_sPitchAndFamily&TMPF_TRUETYPE)
        && (PRIMARYLANGID(LANGIDFROMLCID(_afxGlobalData._lcidUserDefault))==LANG_ENGLISH);

    // if fix pitch, the tm bit is clear
    _fFixPitchFont = !(TMPF_FIXED_PITCH & tm.tmPitchAndFamily);
    _xDefDBCWidth = 0;

    {
        _fLatin1CoverageSuspicious &= (tm.tmLastChar == 0x2122);
    }

    if(_bCharSet == SYMBOL_CHARSET)
    {
        // Must use doc codepage, unless of course we have a Unicode document
        // In this event, we pick cp1252, just to maximize coverage.
        _sCodePage = IsStraightToUnicodeCodePage(cpDoc) ? CP_1252 : cpDoc;

        _bConvertMode = CVM_SYMBOL;
    }

Cleanup:
    if(hfontOld)
    {
        SelectFontEx(hdc, hfontOld);
    }

    return fRes;
}

/*
*  CBaseCcs::NeedConvertNBSPs
*
*  @mfunc
*      Determine NBSPs need conversion or not during render
*      Some fonts wont render NBSPs correctly.
*      Flag fonts which have this problem.
*
*/
BOOL CBaseCcs::NeedConvertNBSPs(HDC hdc, CDocument* pDoc)
{
    HFONT hfontOld;

    Assert(!_fConvertNBSPsSet);

    if(GetCurrentObject(hdc, OBJ_FONT) != _hfont)
    {
        hfontOld = SelectFontEx(hdc, _hfont);

        if(!hfontOld)
        {
            goto Cleanup;
        }
    }
    else
    {
        hfontOld = 0;
    }

    // BUGBUG (cthrash) Once ExtTextOutW is supported in CRenderer, we need
    //                  to set _fConvertNBSPsIfA, so that we can better tune when we need to
    //                  convert NBSPs depending on which ExtTextOut variant we call.
    {
        ABC abcSpace, abcNbsp;

        _fConvertNBSPs = !GetCharABCWidthsW(hdc, L' ', L' ', &abcSpace) ||
            !GetCharABCWidthsW(hdc, WCH_NBSP, WCH_NBSP, &abcNbsp) ||
            abcSpace.abcA!=abcNbsp.abcA ||
            abcSpace.abcB!=abcNbsp.abcB ||
            abcSpace.abcC!=abcNbsp.abcC;
    }

Cleanup:
    if(hfontOld)
    {
        SelectFontEx(hdc, hfontOld);
    }

    _fConvertNBSPsSet = TRUE;
    return TRUE;
}

/*
*  CBaseCcs::DestroyFont
*
*  @mfunc
*      Destroy font handle for this CBaseCcs
*
*/
void CBaseCcs::DestroyFont()
{
    // clear out any old font
    if(_hfont)
    {
        DeleteFontEx(_hfont);
        _hfont = 0;
    }

    // make sure the script cache is freed
    if(_sc)
    {
        ::ScriptFreeCache(&_sc);
        // NB (mikejoch) If ScriptFreeCache() fails then there is no way to
        // free the cache, so we'll end up leaking it. This shouldn't ever
        // happen since the only way for _sc to be non- NULL is via some other
        // USP function succeeding.
    }

    Assert(_sc == NULL);
}

/*
*  CBaseCcs::Compare (pcf, lfHeight, CBaseCcs * pBaseBaseCcs )
*
*  @mfunc
*      Compares this font cache with the font properties of a
*      given CHARFORMAT
*
*  @rdesc
*      FALSE iff did not match exactly.
*/
BOOL CBaseCcs::Compare(CompareArgs* pCompareArgs)
{
    VerifyLFAtom();
    // NB: We no longer create our logical font with an underline and strike through.
    // We draw strike through & underline separately.

    // If are mode is CM_MULTIBYTE, we need the sid to match exactly, otherwise we
    // will not render correctly.  For example, <FONT FACE=Arial>A&#936; will have two
    // text runs, first sidAsciiLatin, second sidCyrillic.  If are conversion mode is
    // multibyte, we need to make two fonts, one with ANSI_CHARSET, the other with
    // RUSSIAN_CHARSET.
    const CCharFormat* pcf = pCompareArgs->pcf;

    BOOL fMatched = (_yCfHeight==pcf->_yHeight)     // because different mapping modes
        && (_lf.lfWeight==(pcf->_fBold?FW_BOLD:FW_NORMAL))
        && (_latmLFFaceName==pcf->_latmFaceName)
        && (_lf.lfItalic==pcf->_fItalic)
        && (_lf.lfHeight==pCompareArgs->lfHeight)   // have diff logical coords
        && (pcf->_bCharSet==DEFAULT_CHARSET
        || _bCharSet==DEFAULT_CHARSET
        || pcf->_bCharSet==_bCharSet)
        && (_bPitchAndFamily==pcf->_bPitchAndFamily);

    DEBUG_ONLY(if(!fMatched))
    {
        Trace((_T("%s%s%s%s%s%s%s\n"),
            (_yCfHeight==pcf->_yHeight) ? _T("") : _T("height "),
            (_lf.lfWeight==(pcf->_fBold ? FW_BOLD : FW_NORMAL)) ? _T("") : _T("weight "),
            (_latmLFFaceName==pcf->_latmFaceName) ? _T("") : _T("facename "),
            (_lf.lfItalic==pcf->_fItalic) ? _T("") : _T("italicness "),
            (_lf.lfHeight==pCompareArgs->lfHeight) ? _T("") : _T("logical-height "),
            (pcf->_bCharSet==DEFAULT_CHARSET || _bCharSet==DEFAULT_CHARSET || pcf->_bCharSet==_bCharSet) ? _T("") : _T("charset "),
            (_lf.lfPitchAndFamily==pcf->_bPitchAndFamily) ? _T("") : _T("pitch&family")));
    }

    return fMatched;
}

BOOL CBaseCcs::CompareForFontLink(CompareArgs* pCompareArgs)
{
    // The difference between CBaseCcs::Compare and CBaseCcs::CompareForFontLink is in it's
    // treatment of adjusted/pre-adjusted heights.  If we ask to fontlink with MS Mincho
    // in the middle of 10px Arial text, we may choose to instantiate an 11px MS Mincho to
    // compenstate for the ascent/descent discrepancies.  10px is the _yOriginalHeight, and
    // 11px is the _lf.lfHeight in this scenario.  If we again ask for 10px MS Mincho while
    // fontlinking Arial, we want to match based on the original height, not the adjust height.
    // CBaseCcs::Compare, on the other hand, is only concerned with the adjusted height.
    VerifyLFAtom();
    // NB: We no longer create our logical font with an underline and strike through.
    // We draw strike through & underline separately.

    // If are mode is CM_MULTIBYTE, we need the sid to match exactly, otherwise we
    // will not render correctly.  For example, <FONT FACE=Arial>A&#936; will have two
    // text runs, first sidAsciiLatin, second sidCyrillic.  If are conversion mode is
    // multibyte, we need to make two fonts, one with ANSI_CHARSET, the other with
    // RUSSIAN_CHARSET.
    const CCharFormat* pcf = pCompareArgs->pcf;

    BOOL fMatched = (_yCfHeight==pcf->_yHeight)         // because different mapping modes
        && (_lf.lfWeight==(pcf->_fBold?FW_BOLD:FW_NORMAL))
        && (_latmLFFaceName==pcf->_latmFaceName)
        && (_latmBaseFaceName==pCompareArgs->latmBaseFaceName)
        && (_lf.lfItalic==pcf->_fItalic)
        && (_yOriginalHeight==pCompareArgs->lfHeight)   // have diff logical coords
        && (pcf->_bCharSet==DEFAULT_CHARSET
        || _bCharSet==DEFAULT_CHARSET
        || pcf->_bCharSet==_bCharSet)
        && (_bPitchAndFamily==pcf->_bPitchAndFamily);

    DEBUG_ONLY(if(!fMatched))
    {
        Trace((_T("%s%s%s%s%s%s%s\n"),
            (_yCfHeight==pcf->_yHeight) ? _T("") : _T("height "),
            (_lf.lfWeight==(pcf->_fBold?FW_BOLD:FW_NORMAL)) ? _T("") : _T("weight "),
            (_latmLFFaceName==pcf->_latmFaceName) ? _T("") : _T("facename "),
            (_lf.lfItalic==pcf->_fItalic) ? _T("") : _T("italicness "),
            (_yOriginalHeight==pCompareArgs->lfHeight) ? _T("") : _T("logical-height "),
            (pcf->_bCharSet==DEFAULT_CHARSET || _bCharSet==DEFAULT_CHARSET || pcf->_bCharSet==_bCharSet) ? _T("") : _T("charset "),
            (_lf.lfPitchAndFamily==pcf->_bPitchAndFamily) ? _T("") : _T("pitch&family")));
    }

    return fMatched;
}

void CBaseCcs::SetLFFaceNameAtm(LONG latmFaceName)
{
    VerifyLFAtom();
    _latmLFFaceName= latmFaceName;
    // Sets the string inside _lf to what _latmLFFaceName represents.
    _tcsncpy(_lf.lfFaceName, fc().GetFaceNameFromAtom(_latmLFFaceName), LF_FACESIZE);
}

void CBaseCcs::SetLFFaceName(const TCHAR* szFaceName)
{
    VerifyLFAtom();
    _latmLFFaceName= fc().GetAtomFromFaceName(szFaceName);
    _tcsncpy(_lf.lfFaceName, szFaceName, LF_FACESIZE);
}


CONVERTMODE CBaseCcs::GetConvertMode(BOOL fEnhancedMetafile, BOOL fMetafile)
{
    CONVERTMODE cm = (CONVERTMODE)_bConvertMode;

    // For hack around ExtTextOutW Win95 FE problems.
    // NB (cthrash) The following is an enumeration of conditions under which
    // ExtTextOutW is broken.  This code is originally from RichEdit.
    if(cm == CVM_MULTIBYTE)
    {
        // If we want CVM_MULTIBYTE, that's what we get.
    }
    else
    {
        if(CVM_SYMBOL != cm)
        {
            if(fMetafile && _afxGlobalData._fFarEastWinNT)
            {
                // FE NT metafile ExtTextOutW hack.
                cm = CVM_MULTIBYTE;
            }
        }
    }

    return cm;
}

/*
*  CBaseCcs::FillWidths (ch, rlWidth)
*
*  @mfunc
*      Fill in this CBaseCcs with metrics info for given device
*
*  @rdesc
*      TRUE if OK, FALSE if failed
*/
BOOL CBaseCcs::FillWidths(
          HDC hdc,
          TCHAR ch,         //@parm The TCHAR character we need a width for.
          LONG& rlWidth)    //@parm the width of the character
{
    BOOL  fRes = FALSE;
    HFONT hfontOld;

    hfontOld = PushFont(hdc);

    // fill up the width info.
    fRes = _widths.FillWidth(hdc, this, ch, rlWidth);

    PopFont(hdc, hfontOld);

    return fRes;
}

// Selects the instance var _hfont so we can measure, but remembers the old
// font so we don't actually change anything when we're done measuring.
// Returns the previously selected font.
HFONT CBaseCcs::PushFont(HDC hdc)
{
    HFONT hfontOld;

    AssertSz(_hfont, "CBaseCcs has no font");

    //  The mapping mode for the HDC is set before we get here.
    hfontOld = (HFONT)GetCurrentObject(hdc, OBJ_FONT);
    Assert(hfontOld != NULL); // Otherwise using NULL is invalid.
    if(hfontOld != _hfont)
    {
        DEBUG_ONLY(HFONT hfontReturn = )SelectFontEx(hdc, _hfont);
        AssertSz(hfontReturn == hfontOld, "GDI failure changing fonts");
    }
    return hfontOld;
}

// Restores the selected font from before PushFont.
// This is really just a lot like SelectFont, but is optimized, and
// will only work if PushFont was called before it.
void CBaseCcs::PopFont(HDC hdc, HFONT hfontOld)
{
    // This assert will fail if Pushfont was not called before popfont,
    // Or if somebody else changes fonts in between.  (They shouldn't.)
    Assert(_hfont == (HFONT)GetCurrentObject(hdc, OBJ_FONT));
#ifndef _WIN64
    //$ Win64: GetObjectType is returning zero on Axp64
    Assert(OBJ_FONT == GetObjectType(hfontOld));
#endif
    if(hfontOld != _hfont)
    {
        DEBUG_ONLY(HFONT hfontReturn = )SelectFontEx(hdc, hfontOld);
        AssertSz(hfontReturn==_hfont, "GDI failure changing fonts");
    }
}

// Goes into a critical section, and allocates memory for a width cache.
void CWidthCache::ThreadSafeCacheAlloc(void** ppCache, size_t iSize)
{
    EnterCriticalSection(&(fc()._csOther));

    if(!*ppCache)
    {
        *ppCache = MemAllocClear(iSize);
    }

    LeaveCriticalSection(&(fc()._csOther));
}

// Fills in the widths of the low 128 characters.
// Allocates memory for the block if needed.
BOOL CWidthCache::PopulateFastWidthCache(HDC hdc, CBaseCcs* pBaseCcs)
{
    BOOL    fRes;
    HFONT   hfontOld;

    // First switch to the appropriate font.
    hfontOld = pBaseCcs->PushFont(hdc);

    // characters in 0 - 127 range (cache 0), so initialize the character widths for
    // for all of them.
    int widths[ FAST_WIDTH_CACHE_SIZE ];
    int i;

    {
        fRes = GetCharWidth32(hdc, 0, FAST_WIDTH_CACHE_SIZE-1, widths);
    }

    // Copy the results back into the real cache, if it worked.
    if(fRes)
    {
        Assert(!_pFastWidthCache); // Since we should only populate this once.

        ThreadSafeCacheAlloc((void**)&_pFastWidthCache, sizeof(CharWidth)*FAST_WIDTH_CACHE_SIZE);

        if(!_pFastWidthCache)
        {
            // We're kinda screwed if we can't get memory for the cache.
            AssertSz(0,"Failed to allocate fast width cache.");
            fRes = FALSE;
            goto Cleanup;
        }

        for(i=0; i<FAST_WIDTH_CACHE_SIZE; i++)
        {
            Assert( widths[i] <= MAXSHORT ); // since there isn't a MAXUSHORT
            _pFastWidthCache[i]= (widths[i]) ? widths[i] : pBaseCcs->_xAveCharWidth;
        }

        // NB (cthrash) Measure NBSPs with space widths.
        SetCacheEntry(WCH_NBSP, _pFastWidthCache[_T(' ')]);
    }

Cleanup:
    pBaseCcs->PopFont(hdc, hfontOld);

    return fRes;
}

// =========================  WidthCache by jonmat  =========================
/*
*  CWidthCache::FillWidth(hdc, ch, xOverhang, rlWidth)
*
*  @mfunc
*      Call GetCharWidth() to obtain the width of the given char.
*
*  @comm
*      The HDC must be setup with the mapping mode and proper font
*      selected *before* calling this routine.
*
*  @rdesc
*      Returns TRUE if we were able to obtain the widths.
*
*/
BOOL CWidthCache::FillWidth(
        HDC         hdc,
        CBaseCcs*   pBaseCcs,   //@parm CCcs object
        const TCHAR ch,         //@parm Char to obtain width for
        LONG&       rlWidth)    //@parm Width of character
{
    BOOL fRes;
    INT numOfDBCS = 0;
    CacheEntry widthData;

    // BUGBUG GetCharWidthW is really broken for bullet on Win95J. Sometimes it will return
    // a width or 0 or 1198 or ...So, hack around it. Yuk!
    // Also, WideCharToMultiByte() on Win95J will NOT convert bullet either.
    Assert(!IsCharFast(ch)); // This code shouldn't be called for that.

    if(ch == WCH_NBSP)
    {
        // Use the width of a space for an NBSP
        fRes = pBaseCcs->Include(hdc, L' ', rlWidth);

        if(!fRes)
        {
            // Error condition, just use the default width.
            rlWidth = pBaseCcs->_xAveCharWidth;
        }

        widthData.ch = ch;
        widthData.width = rlWidth;
        *GetEntry(ch) = widthData;
    }
    else
    {
        INT xWidth = 0;

        // Diacritics, tone marks, and the like have 0 width so return 0 if
        // GetCharWidthW() succeeds.
        BOOL fZeroWidth = IsZeroWidth(ch);

        if(pBaseCcs->_bConvertMode == CVM_SYMBOL)
        {
            TCHAR chT = ch;

            if(chT > 255)
            {
                const BYTE b = InWindows1252ButNotInLatin1(ch);
                if(b)
                {
                    chT = b;
                }
            }

            fRes = chT>255 ? FALSE : GetCharWidthA(hdc, chT, chT, &xWidth);
        }
        else if(pBaseCcs->_bConvertMode!=CVM_MULTIBYTE)
        {
            // GetCharWidthW will crash on a 0xffff.
            Assert(ch != 0xffff);
            fRes = GetCharWidthW(hdc, ch, ch, &xWidth);
        }
        else
        {
            fRes = FALSE;
        }

        // either fAnsi case or GetCharWidthW fail, let's try GetCharWidthA
        if(!fRes || (0==xWidth && !fZeroWidth))
        {
            WORD wDBCS;
            char ansiChar[2] = {0};
            UINT uCP = pBaseCcs->_sCodePage;

            // Convert string
            numOfDBCS = WideCharToMultiByte(uCP, 0, &ch, 1, ansiChar, 2, NULL, NULL);

            if(2 == numOfDBCS)
            {
                wDBCS = (BYTE)ansiChar[0]<<8 | (BYTE)ansiChar[1];
            }
            else
            {
                wDBCS = (BYTE)ansiChar[0];
            }

            fRes = GetCharWidthA(hdc, wDBCS, wDBCS, &xWidth);
        }

        widthData.width = (USHORT)xWidth;

        if(fRes)
        {
            if(0==widthData.width && !fZeroWidth)
            {
                // Sometimes GetCharWidth will return a zero length for small
                // characters. When this happens we will use the default width
                // for the font if that is non-zero otherwise we just us 1
                // because this is the smallest valid value.

                // BUGBUG - under Win95 Trad. Chinese, there is a bug in the
                // font. It is returning a width of 0 for a few characters
                // (Eg 0x09F8D, 0x81E8) In such case, we need to use 2 *
                // pBaseCcs->_xAveCharWidth since these are DBCS
                if(0 == pBaseCcs->_xAveCharWidth)
                {
                    widthData.width = 1;
                }
                else
                {
                    widthData.width = (numOfDBCS==2)
                        ? (pBaseCcs->_xDefDBCWidth?pBaseCcs->_xDefDBCWidth:2*pBaseCcs->_xAveCharWidth)
                        : pBaseCcs->_xAveCharWidth;
                } 
            }
            widthData.ch = ch;
            if(widthData.width <= pBaseCcs->_xOverhangAdjust)
            {
                widthData.width = 1;
            }
            else
            {
                widthData.width   -= pBaseCcs->_xOverhangAdjust;
            }
            rlWidth = widthData.width;
            *GetEntry(ch) = widthData;
        }
    }

    AssertSz(fRes, "no width?");

    Assert(widthData.width == rlWidth); // Did we forget to set it?

    return fRes;
}

/*
*  CWidthCache::~CWidthCache()
*
*  @mfunc
*      Free any allocated caches.
*
*/
CWidthCache::~CWidthCache()
{
    INT i;

    for(i=0; i<TOTALCACHES; i++)
    {
        if(_pWidthCache[i])
        {
            MemFree(_pWidthCache[i]);
        }
    }
    MemFree(_pFastWidthCache);
}

// BUGBUG (cthrash) This needs to be removed as soon as the FontLinkTextOut
// is cleaned up for complex scripts.
void CBaseCcs::EnsureLangBits(HDC hdc)
{
    if(!_dwLangBits)
    {
        // Get the charsets supported by this font.
        if(_bCharSet != SYMBOL_CHARSET)
        {
            _dwLangBits = GetFontScriptBits(hdc, fc().GetFaceNameFromAtom(_latmLFFaceName), &_lf);
        }
        else
        {
            // See comment in GetFontScriptBits.
            // SBITS_ALLLANGS means _never_ fontlink.
            _dwLangBits = SBITS_ALLLANGS;
        }
    }
}

BYTE InWindows1252ButNotInLatin1Helper(WCHAR ch)
{
    for(int i=32;i--;)
    {
        if(ch == g_achLatin1MappingInUnicodeControlArea[i])
        {
            return 0x80+i;
        }
    }

    return 0;
}

//+----------------------------------------------------------------------------
//
//  Function:   CBaseCcs::FixupForFontLink
//
//  Purpose:    Optionally scale a font height when fontlinking.
//
//              This code was borrowed from UniScribe (usp10\str_ana.cxx)
//
//              Let's say you're base font is a 10pt Tahoma, and you need
//              to substitute a Chinese font (e.g. MingLiU) for some ideo-
//              graphic characters.  When you simply ask for a 10pt MingLiU,
//              you'll get a visibly smaller font, due to the difference in
//              in distrubution of the ascenders/descenders.  The purpose of
//              this function is to examine the discrepancy and pick a
//              slightly larger or smaller font for improved legibility. We
//              may also adjust the baseline by 1 pixel.
//              
//-----------------------------------------------------------------------------
void CBaseCcs::FixupForFontLink(HDC hdc, CBaseCcs* pBaseBaseCcs)
{
    int iOriginalDescender = pBaseBaseCcs->_yDescent;
    int iFallbackDescender = _yDescent;
    int iOriginalAscender  = pBaseBaseCcs->_yHeight - iOriginalDescender;
    int iFallbackAscender  = _yHeight - iFallbackDescender;

    if(iFallbackAscender>0 && iFallbackDescender>0)
    {
        int iAscenderRatio  = 1024 * iOriginalAscender  / iFallbackAscender;
        int iDescenderRatio = 1024 * iOriginalDescender / iFallbackDescender;
        int iNewRatio;
        SHORT yDescentAdjust = 0;

        if(iAscenderRatio != iDescenderRatio)
        {
            // We'll allow moving the baseline by one pixel to reduce the amount of
            // scaling required.
            if(iAscenderRatio < iDescenderRatio)
            {
                // Clipping, if any, would happen in the ascender.
                yDescentAdjust = +1;
                iOriginalAscender++;    // Move the baseline down one pixel.
                iOriginalDescender--;
                Trace0("Moving baseline down one pixel to leave more room for ascender\n");
            }
            else
            {
                // Clipping, if any, would happen in the descender.
                yDescentAdjust = -1;
                iOriginalAscender--;    // Move the baseline up one pixel.
                iOriginalDescender++;
                Trace0("Moving baseline up one pixel to leave more room for descender\n");
            }

            // Recalculate ascender and descender ratios based on shifted baseline
            iAscenderRatio  = 1024 * iOriginalAscender  / iFallbackAscender;
            iDescenderRatio = 1024 * iOriginalDescender / iFallbackDescender;
        }

        // Establish extent of worst mismatch, either too big or too small
        iNewRatio = iAscenderRatio;
        iNewRatio = max(iNewRatio, 768); // Never reduce size by over 25%
        if(iNewRatio<1000 || iNewRatio>1048)
        {
            LONG lfHeightCurrent = _lf.lfHeight;
            LONG lAdjust = (iNewRatio<1024) ? 1023 : 0; // round towards 100% (1024)
            LONG lfHeightNew = (lfHeightCurrent*iNewRatio-lAdjust) / 1024;

            Assert(lfHeightCurrent < 0); // lfHeight should be negative; otherwise rounding will be incorrect

            if(lfHeightNew != lfHeightCurrent)
            {
                SHORT yHeightCurrent = _yHeight;
                SHORT sCodePageCurrent = _sCodePage;
                TCHAR achNewFaceName[LF_FACESIZE];
                HFONT hfontCurrent = _hfont;

                // Reselect with new ratio
                Trace0("Reselecting fallback font to improve legibility");
                Trace((_T(" Original font ascender %4d, descender %4d, lfHeight %4d, \'%S\'"),
                    iOriginalAscender, iOriginalDescender, pBaseBaseCcs->_yHeight, fc().GetFaceNameFromAtom(pBaseBaseCcs->_latmLFFaceName)));
                Trace((_T(" Fallback font ascender %4d, descender %4d, -> lfHeight %4d, \'%s\'\n"),
                    iFallbackAscender, iFallbackDescender, _yHeight*iNewRatio/1024, fc().GetFaceNameFromAtom(_latmLFFaceName)));

                _lf.lfHeight = lfHeightNew;

                if(GetFontWithMetrics(hdc, achNewFaceName, CP_UCS_2, 0))
                {
                    DeleteFontEx(hfontCurrent);

                    _yDescent += yDescentAdjust;
                }
                else
                {
                    Assert(_hfont == NULL);

                    _lf.lfHeight = lfHeightCurrent;
                    _yHeight = yHeightCurrent;
                    _hfont = hfontCurrent;
                }

                _sCodePage = sCodePageCurrent;
            }
        }
    }

    _fHeightAdjustedForFontlinking = TRUE;
}
