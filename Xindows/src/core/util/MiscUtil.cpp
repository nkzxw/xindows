
#include "stdafx.h"
#include "MiscUtil.h"

LONG    g_cMetricChange = 0;

SIZE    g_sizeSystemChar;
BOOL    g_fScreenReader = FALSE;

void GetSystemNumberSettings(NUMSHAPE* piNumShape, DWORD* plangNationalDigits)
{
    NUMSHAPE iNumShape = NUMSHAPE_NONE;
    DWORD   langDigits = LANG_NEUTRAL;
    HKEY    hkey = NULL;
    DWORD   dwType;
    DWORD   cbData;
    char    achBufferData[41]; // Max Size user can fit in our variables edit field
    WCHAR   achDigits[16];

    Assert(piNumShape!=NULL && plangNationalDigits!=NULL);

    if(RegOpenKeyEx(HKEY_CURRENT_USER, TEXT("Control Panel\\International"), 0L, KEY_READ, &hkey))
    {
        goto Cleanup;
    }

    Assert(hkey != NULL);

    cbData = sizeof(achBufferData);
    if(RegQueryValueExA(hkey, "NumShape",
        0L, &dwType, (LPBYTE)achBufferData, &cbData)==ERROR_SUCCESS &&
        achBufferData[0]!=TEXT('\0') && (dwType&REG_SZ) && !(dwType&REG_NONE))
    {
        iNumShape = (NUMSHAPE)max(min(atoi(achBufferData), (int)NUMSHAPE_NATIVE), (int)NUMSHAPE_CONTEXT);
    }

    RegCloseKey(hkey);

    if(GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SNATIVEDIGITS, achDigits, 16))
    {
        switch(achDigits[4])
        {
        case 0x664: // Arabic-Indic digits
            langDigits = LANG_ARABIC;
            break;
        case 0x6F4: // Eastern Arabic-Indic digits
            langDigits = LANG_FARSI;
            break;
        case 0x96A: // Devanagari digits
            langDigits = LANG_HINDI;
            break;
        case 0x9EA: // Bengali digits
            langDigits = LANG_BENGALI;
            break;
        case 0xA6A: // Gurmukhi digits
            langDigits = LANG_PUNJABI;
            break;
        case 0xAEA: // Gujarati digits
            langDigits = LANG_GUJARATI;
            break;
        case 0xB6A: // Oriya digits
            langDigits = LANG_ORIYA;
            break;
        case 0xBEA: // Tamil digits
            langDigits = LANG_TAMIL;
            break;
        case 0xC6A: // Telugu digits
            langDigits = LANG_TELUGU;
            break;
        case 0xCEA: // Kannada digits
            langDigits = LANG_KANNADA;
            break;
        case 0xD6A: // Malayalam digits
            langDigits = LANG_MALAYALAM;
            break;
        case 0xE54: // Thai digits
            langDigits = LANG_THAI;
            break;
        case 0xED4: // Lao digits
            langDigits = LANG_LAO;
            break;
        case 0xF24: // Tibetan digits
            langDigits = LANG_TIBETAN;
            break;
        default:
            langDigits = LANG_NEUTRAL;
            break;
        }
    }
    else
    {
        // Work from the platform's locale.
        langDigits = PRIMARYLANGID(GetUserDefaultLangID());
    }

    if (langDigits!=LANG_ARABIC &&
        langDigits!=LANG_FARSI &&
        langDigits!=LANG_HINDI &&
        langDigits!=LANG_BENGALI &&
        langDigits!=LANG_PUNJABI &&
        langDigits!=LANG_GUJARATI &&
        langDigits!=LANG_ORIYA &&
        langDigits!=LANG_TAMIL &&
        langDigits!=LANG_TELUGU &&
        langDigits!=LANG_KANNADA &&
        langDigits!=LANG_MALAYALAM &&
        langDigits!=LANG_THAI &&
        langDigits!=LANG_LAO &&
        langDigits!=LANG_TIBETAN)
    {
        langDigits = LANG_NEUTRAL;
        iNumShape = NUMSHAPE_NONE;
    }

Cleanup:
    *piNumShape = iNumShape;
    *plangNationalDigits = langDigits;
}

HRESULT InitSystemMetricValues(THREADSTATE* pts)
{
    HIGHCONTRAST    hc;
    HFONT           hfontOld;
    TEXTMETRIC      tm;

    InterlockedIncrement(&g_cMetricChange);

    if(!pts->hdcDesktop)
    {
        pts->hdcDesktop = CreateCompatibleDC(NULL);
        if(!pts->hdcDesktop)
        {
            RRETURN(E_OUTOFMEMORY);
        }
    }

    _afxGlobalData._sizePixelsPerInch.cx = GetDeviceCaps(pts->hdcDesktop, LOGPIXELSX);
    _afxGlobalData._sizePixelsPerInch.cy = GetDeviceCaps(pts->hdcDesktop, LOGPIXELSY);

    g_sizeDragMin.cx = GetSystemMetrics(SM_CXDRAG);
    g_sizeDragMin.cy = GetSystemMetrics(SM_CYDRAG);

    _afxGlobalData._sizeScrollbar.cx = GetSystemMetrics(SM_CXVSCROLL);
    _afxGlobalData._sizeScrollbar.cy = GetSystemMetrics(SM_CYHSCROLL);
    _afxGlobalData._sizelScrollbar.cx = HimetricFromHPix(_afxGlobalData._sizeScrollbar.cx);
    _afxGlobalData._sizelScrollbar.cy = HimetricFromVPix(_afxGlobalData._sizeScrollbar.cy);

    // System font info
    hfontOld = (HFONT)SelectObject(pts->hdcDesktop, GetStockObject(SYSTEM_FONT));
    if(hfontOld)
    {
        GetTextMetrics(pts->hdcDesktop, &tm);

        g_sizeSystemChar.cx = tm.tmAveCharWidth;
        g_sizeSystemChar.cy = tm.tmHeight;

        SelectObject(pts->hdcDesktop, hfontOld);
    }
    else
    {
        g_sizeSystemChar.cx = g_sizeSystemChar.cy = 10;
    }

    // Locale info
    _afxGlobalData._cpDefault = GetACP();
    _afxGlobalData._lcidUserDefault = GetSystemDefaultLCID(); // Set Global Locale ID

    GetSystemNumberSettings(&_afxGlobalData._iNumShape, &_afxGlobalData._uLangNationalDigits);

    // Accessibility info
    SystemParametersInfo(SPI_GETSCREENREADER, 0, &g_fScreenReader, FALSE);

    memset(&hc, 0, sizeof(HIGHCONTRAST));
    hc.cbSize = sizeof(HIGHCONTRAST);
    if(SystemParametersInfo(SPI_GETHIGHCONTRAST, sizeof(HIGHCONTRAST), &hc, 0))
    {
        _afxGlobalData._fHighContrastMode = !!(hc.dwFlags & HCF_HIGHCONTRASTON);
    }
    else
    {
        Trace1("SPI failed with error %x\n", GetLastError());
    }

    RRETURN(S_OK);
}

//+-------------------------------------------------------------------------
//
//  Function:   DeinitSystemMetricValues
//
//  Synopsis:   Deinitializes globals holding system metric values.
//
//--------------------------------------------------------------------------
void DeinitSystemMetricValues(THREADSTATE* pts)
{
    if(pts->hdcDesktop)
    {
        Verify(DeleteDC(pts->hdcDesktop));
    }
}

//+---------------------------------------------------------------------------
//
//  Function:       IsFarEastLCID(lcid)
//
//  Returns:        True iff lcid is a Far East locale.
//
//----------------------------------------------------------------------------
BOOL IsFarEastLCID(LCID lcid)
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

//+---------------------------------------------------------------------------
//  COMPLEXSCRIPT
//  Function:       IsBidiLCID(lcid)
//
//  Returns:        True iff lcid is a right to left locale.
//
//----------------------------------------------------------------------------
BOOL IsBidiLCID(LCID lcid)
{
    switch(PRIMARYLANGID(LANGIDFROMLCID(lcid)))
    {
    case LANG_ARABIC:
    case LANG_FARSI:
    case LANG_HEBREW:
        return TRUE;
    }

    return FALSE;
}

#ifndef LANG_BURMESE
#define LANG_BURMESE    0x55 // Burma
#endif
//+---------------------------------------------------------------------------
//  COMPLEXSCRIPT
//  Function:       IsComplexLCID(lcid)
//
//  Returns:        True iff lcid is a complex script locale.
//
//----------------------------------------------------------------------------
BOOL IsComplexLCID(LCID lcid)
{
    switch(PRIMARYLANGID(LANGIDFROMLCID(lcid)))
    {
    case LANG_ARABIC:
    case LANG_ASSAMESE:
    case LANG_BENGALI:
    case LANG_BURMESE:
    case LANG_FARSI:
    case LANG_GUJARATI:
    case LANG_HEBREW:
    case LANG_HINDI:
    case LANG_KANNADA:
    case LANG_KASHMIRI:
    case LANG_KHMER:
    case LANG_KONKANI:
    case LANG_LAO:
    case LANG_MALAYALAM:
    case LANG_MANIPURI:
    case LANG_MARATHI:
    case LANG_MONGOLIAN:
    case LANG_NAPALI:
    case LANG_ORIYA:
    case LANG_PUNJABI:
    case LANG_SANSKRIT:
    case LANG_SINDHI:
    case LANG_TAMIL:
    case LANG_TELUGU:
    case LANG_THAI:
    case LANG_TIBETAN:
    case LANG_URDU:
    case LANG_VIETNAMESE:
    case LANG_YIDDISH:
        return TRUE;
    }

    return FALSE;
}