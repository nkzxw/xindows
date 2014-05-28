
#include "stdafx.h"
#include "Settings.h"

#define MAX_REG_VALUE_LENGTH    50

enum RKI_TYPE
{
    RKI_KEY,
    RKI_CPKEY,
    RKI_BOOL,
    RKI_COLOR,
    RKI_FONT,
    RKI_SIZE,
    RKI_CP,
    RKI_BYTEBOOL,
    RKI_STRING,
    RKI_ANCHORUNDERLINE,
    RKI_DWORD,
    RKI_TYPE_Last_Enum
};

struct REGKEYINFORMATION
{
    TCHAR*  pszName;            // Name of the value or key
    BYTE    rkiType;            // Type of entry
    size_t  cbOffset;           // Offset of member to store data in
    size_t  cbOffsetCondition;  // Offset of member that must be false (true only
                                // for type RKI_STRING) to use this value. If
                                // it's true (false for type RKI_STRING)
                                // this value is left as its default value. Assumed
                                // to be data that's sizeof(BYTE). If 0 then
                                // no condition is used.
    BOOL    fLocalState;        // True if this is local state which may not always get
                                // updated when read again from the registry.
};

HRESULT ReadSettingsFromRegistry(
         TCHAR* pchKeyPath,
         const REGKEYINFORMATION* pAryKeys,
         int    iKeyCount,
         void*  pBase,
         DWORD  dwFlags,
         BOOL   fSettingsRead,
         void*  pUserData)
{
    LONG    lRet;
    HKEY    hKeyRoot = NULL;
    HKEY    hKeySub  = NULL;
    int     i;
    DWORD   dwType;
    DWORD   dwSize;

    // IEUNIX
    // Note we access dwDataBuf as a byte array through bDataBuf
    // but DWORD align it by declaring it a DWORD array (dwDataBuf).
    DWORD   dwDataBuf[pdlUrlLen/sizeof(DWORD)+1];
    BYTE*   bDataBuf = (BYTE*)dwDataBuf;

    TCHAR   achCustomKey[64], *pch;
    BYTE*   pbData;
    BYTE    bCondition;
    BOOL    fUpdateLocalState;
    const REGKEYINFORMATION* prki;
    LONG*   pl;

    Assert(pBase);

    // Do not re-read unless explictly asked to do so.
    if(fSettingsRead && !(dwFlags&REGUPDATE_REFRESH))
    {
        return S_OK;
    }

    // Always read local settings at least once
    fUpdateLocalState = !fSettingsRead || (dwFlags&REGUPDATE_OVERWRITELOCALSTATE);

    // Get a registry key handle
    lRet = RegOpenKeyEx(HKEY_CURRENT_USER, pchKeyPath, 0, KEY_READ, &hKeyRoot);
    if(lRet != ERROR_SUCCESS)
    {
        return S_FALSE;
    }

    for(i=0; i<iKeyCount; i++)
    {
        prki = &pAryKeys[i];
        // Do not update local state unless asked to do so.
        if(!fUpdateLocalState && prki->fLocalState)
        {
            continue;
        }
        switch(prki->rkiType)
        {
        case RKI_KEY:
        case RKI_CPKEY:
            if(!prki->pszName)
            {
                hKeySub = hKeyRoot;
            }
            else
            {
                if(hKeySub && (hKeySub!=hKeyRoot))
                {
                    RegCloseKey(hKeySub);
                    hKeySub = NULL;
                }

                if(prki->rkiType == RKI_CPKEY)
                {
                    // N.B. (johnv) It is assumed here that pUserData points
                    // to a codepage if we are looking for codepage settings.
                    // RKI_CPKEY entries are per family codepage.
                    Assert(pUserData);
                    Format(_afxGlobalData._hInstResource, 0, achCustomKey, ARRAYSIZE(achCustomKey),
                        prki->pszName, *(DWORD*)pUserData);

                    pch = achCustomKey;
                }
                else
                {
                    pch = prki->pszName;
                }

                lRet = RegOpenKeyEx(
                    hKeyRoot,
                    pch,
                    0,
                    KEY_READ,
                    &hKeySub);

                if(lRet != ERROR_SUCCESS)
                {
                    // We couldn't get this key, skip it.
                    i++;
                    while(i<iKeyCount && pAryKeys[i].rkiType!=RKI_KEY && pAryKeys[i].rkiType!=RKI_CPKEY)
                    {
                        i++;
                    }

                    i--; // Account for the fact that continue will increment i again.
                    hKeySub = NULL;
                    continue;
                }
            }
            break;

        case RKI_SIZE:
            Assert(hKeySub);
            dwSize = MAX_REG_VALUE_LENGTH;
            lRet = RegQueryValueEx(
                hKeySub,
                prki->pszName,
                0,
                &dwType,
                bDataBuf,
                &dwSize);

            if(lRet == ERROR_SUCCESS)
            {
                short s;

                if(dwType == REG_BINARY)
                {
                    s = (short)*(BYTE*)bDataBuf;
                }
                else if(dwType == REG_DWORD)
                {
                    s = (short)*(DWORD*)bDataBuf;
                }
                else
                {
                    break;
                }

                *(short*)((BYTE*)pBase+prki->cbOffset) =
                    min(short(BASELINEFONTMAX), max(short(BASELINEFONTMIN), s));
            }
            break;

        case RKI_BOOL:
            Assert(hKeySub);

            dwSize = MAX_REG_VALUE_LENGTH;
            lRet = RegQueryValueEx(
                hKeySub,
                prki->pszName,
                0,
                &dwType,
                bDataBuf,
                &dwSize);

            if(lRet == ERROR_SUCCESS)
            {
                pbData = (BYTE*)((BYTE*)pBase + prki->cbOffset);

                if(dwType == REG_DWORD)
                {
                    *pbData = (*(DWORD*)bDataBuf != 0);
                }
                else if(dwType == REG_SZ)
                {
                    TCHAR ch = *(TCHAR*)bDataBuf;

                    if(ch==_T('1') || ch==_T('y') || ch==_T('Y'))
                    {
                        *pbData = TRUE;
                    }
                    else
                    {
                        *pbData = FALSE;
                    }
                }
                else if(dwType == REG_BINARY)
                {
                    *pbData = (*(BYTE*)bDataBuf != 0);
                }
                // Can't convert other types. Just leave it the default.
            }
            break;

        case RKI_FONT:
            Assert(hKeySub);
            dwSize = LF_FACESIZE * sizeof(TCHAR);

            lRet = RegQueryValueEx(
                hKeySub,
                prki->pszName,
                0,
                &dwType,
                bDataBuf,
                &dwSize);

            pl = (LONG*)((BYTE*)pBase + prki->cbOffset);

            if(lRet==ERROR_SUCCESS && dwType==REG_SZ && *(TCHAR*)bDataBuf)
            {
                *pl = fc().GetAtomFromFaceName((TCHAR*)bDataBuf);
            }
            break;

        case RKI_COLOR:
            Assert(hKeySub);

            dwSize = MAX_REG_VALUE_LENGTH;

            bCondition = *(BYTE*)((BYTE*)pBase + prki->cbOffsetCondition);
            if(prki->cbOffsetCondition && bCondition)
            {
                // The appropriate flag is set that says we should not pay
                // attention to this value, so just skip it and keep the
                // default.
                break;
            }

            lRet = RegQueryValueEx(
                hKeySub,
                prki->pszName,
                0,
                &dwType,
                bDataBuf,
                &dwSize);

            if(lRet == ERROR_SUCCESS)
            {
                if(dwType == REG_SZ)
                {
                    // Crack the registry format for colors which is a string
                    // of the form "R,G,B" where R, G, and B are decimal
                    // values for Red, Green, and Blue, respectively.
                    DWORD   colors[3];
                    TCHAR*  pchStart = (TCHAR*)bDataBuf;
                    TCHAR*  pchEnd;
                    int     i;

                    pbData = (BYTE*)((BYTE*)pBase + prki->cbOffset);

                    for(i=0; i<3; i++)
                    {
                        colors[i] = wcstol(pchStart, &pchEnd, 10);

                        if(*pchEnd==_T('\0') && i!=2)
                        {
                            break;
                        }

                        pchStart  = pchEnd + 1;
                    }

                    if(i != 3) // We didn't get all the colors. Abort.
                    {
                        break;
                    }
                    *(COLORREF*)pbData = RGB(colors[0], colors[1], colors[2]);
                }
                // Can't convert other types. Just leave it the default.
            }
            break;

        case RKI_CP:
            Assert(hKeySub);
            dwSize = sizeof(DWORD);
            lRet = RegQueryValueEx(
                hKeySub,
                prki->pszName,
                0,
                &dwType,
                bDataBuf,
                &dwSize);

            if(lRet==ERROR_SUCCESS && dwType==REG_BINARY)
            {
                *(CODEPAGE*)((BYTE*)pBase+prki->cbOffset) = *(CODEPAGE*)bDataBuf;
            }
            break;

        case RKI_BYTEBOOL:
            Assert(hKeySub);
            dwSize = sizeof(DWORD);
            lRet = RegQueryValueEx(
                hKeySub,
                prki->pszName,
                0,
                &dwType,
                bDataBuf,
                &dwSize);

            if(lRet==ERROR_SUCCESS && dwType==REG_BINARY)
            {
                *(BYTE*)((BYTE*)pBase+prki->cbOffset) = (BYTE)!!(*((DWORD*)bDataBuf));
            }
            break;

        case RKI_DWORD:
            Assert(hKeySub);
            dwSize = sizeof(DWORD);
            lRet = RegQueryValueEx(
                hKeySub,
                prki->pszName,
                0,
                &dwType,
                bDataBuf,
                &dwSize);

            if(lRet==ERROR_SUCCESS && (dwType==REG_BINARY || dwType==REG_DWORD))
            {
                *(DWORD*)((BYTE*)pBase+prki->cbOffset) = *(DWORD*)bDataBuf;
            }
            break;

        case RKI_ANCHORUNDERLINE:
            Assert(hKeySub);
            dwSize = MAX_REG_VALUE_LENGTH;
            lRet = RegQueryValueEx(
                hKeySub,
                prki->pszName,
                0,
                &dwType,
                bDataBuf,
                &dwSize);

            if(lRet==ERROR_SUCCESS && dwType==REG_SZ)
            {
                int nAnchorunderline = ANCHORUNDERLINE_YES;

                LPTSTR pchBuffer = (TCHAR*)bDataBuf;
                Assert(pchBuffer != NULL);

                if(pchBuffer)
                {
                    if(_tcsicmp(pchBuffer, _T("yes")) == 0)
                    {
                        nAnchorunderline = ANCHORUNDERLINE_YES;
                    }
                    else if(_tcsicmp(pchBuffer, _T("no")) == 0)
                    {
                        nAnchorunderline = ANCHORUNDERLINE_NO;
                    }
                    else if(_tcsicmp(pchBuffer, _T("hover")) == 0)
                    {
                        nAnchorunderline = ANCHORUNDERLINE_HOVER;
                    }
                }

                *(int*)((BYTE*)pBase+prki->cbOffset) = nAnchorunderline;

            }
            break;

        case RKI_STRING:
            Assert(hKeySub);
            dwSize = 0;
            bCondition = *(BYTE*)((BYTE*)pBase + prki->cbOffsetCondition);
            if((prki->cbOffsetCondition&&bCondition) || !prki->cbOffsetCondition)
            {
                // The appropriate flag is set that says we should pay
                // attention to this value or the flag is 0, then do it
                // else skip and keep the default.

                // get the size of string
                lRet = RegQueryValueEx(
                    hKeySub,
                    prki->pszName,
                    0,
                    &dwType,
                    NULL,
                    &dwSize);

                if(lRet == ERROR_SUCCESS)
                {
                    lRet = RegQueryValueEx(
                        hKeySub,
                        prki->pszName,
                        0,
                        &dwType,
                        bDataBuf,
                        &dwSize);

                    if(lRet==ERROR_SUCCESS && dwType==REG_SZ)
                    {
                        ((CString*)((BYTE*)pBase+prki->cbOffset))->Set((LPCTSTR)bDataBuf);
                    }
                }
            }

            break;

        default:
            AssertSz(FALSE, "Unrecognized RKI Type");
            break;
        }
    }

    if(hKeySub && (hKeySub!=hKeyRoot))
    {
        RegCloseKey(hKeySub);
    }

    RegCloseKey(hKeyRoot);

    return S_OK;
}

//+---------------------------------------------------------------------------
//
//  Member:     CDocument::ReadOptionSettingsFromRegistry, public
//
//  Synopsis:   Read the general option settings from the registry.
//
//  Arguments:  dwFlags - See UpdateFromRegistry.
//
//  Returns:    HRESULT
//
//----------------------------------------------------------------------------
HRESULT CDocument::ReadOptionSettingsFromRegistry(DWORD dwFlags)
{
    HRESULT hr = S_OK;

    // keys in main registry location
    static const REGKEYINFORMATION aOptionKeys[] =
    {
        { NULL,                         RKI_KEY, 0 },
        { _T("Show_FullURL"),           RKI_BOOL,  offsetof(OPTIONSETTINGS, fShowFriendlyUrl), 0, FALSE },
        { _T("SmartDithering"),         RKI_BOOL,  offsetof(OPTIONSETTINGS, fSmartDithering),  0, FALSE },

        { _T("Main"),                   RKI_KEY, 0 },
        { _T("Use_DlgBox_Colors"),      RKI_BOOL, offsetof(OPTIONSETTINGS, fUseDlgColors),    0, FALSE },
        { _T("Anchor Underline"),       RKI_ANCHORUNDERLINE, offsetof(OPTIONSETTINGS, nAnchorUnderline), 0, FALSE },
        { _T("Expand Alt Text"),        RKI_BOOL, offsetof(OPTIONSETTINGS, fExpandAltText),   0, FALSE },
        { _T("Display Inline Images"),  RKI_BOOL, offsetof(OPTIONSETTINGS, fShowImages),      0, FALSE },
        { _T("Display Inline Videos"),  RKI_BOOL, offsetof(OPTIONSETTINGS, fShowVideos),      0, FALSE },
        { _T("Play_Background_Sounds"), RKI_BOOL, offsetof(OPTIONSETTINGS, fPlaySounds),      0, FALSE },
        { _T("Play_Animations"),        RKI_BOOL, offsetof(OPTIONSETTINGS, fPlayAnimations),  0, FALSE },
        { _T("Use Stylesheets"),        RKI_BOOL, offsetof(OPTIONSETTINGS, fUseStylesheets),  0, FALSE },
        { _T("SmoothScroll"),           RKI_BOOL, offsetof(OPTIONSETTINGS, fSmoothScrolling), 0, FALSE },
        { _T("Show image placeholders"),RKI_BOOL, offsetof(OPTIONSETTINGS, fShowImagePlaceholder), 0, FALSE },
        { _T("Move System Caret"),      RKI_BOOL, offsetof(OPTIONSETTINGS, fMoveSystemCaret), 0, FALSE },

        { _T("International"),          RKI_KEY,   0 },
        { _T("Default_CodePage"),       RKI_CP,   offsetof(OPTIONSETTINGS, codepageDefault),  0, TRUE },
        { _T("AutoDetect"),             RKI_BOOL, offsetof(OPTIONSETTINGS, fCpAutoDetect),    0, FALSE },

        { _T("International\\Scripts"), RKI_KEY,   0 },
        { _T("Default_IEFontSize"),     RKI_SIZE,   offsetof(OPTIONSETTINGS, sBaselineFontDefault),  0, FALSE },

        { _T("Settings"),               RKI_KEY, 0 },
        { _T("Background Color"),       RKI_COLOR, offsetof(OPTIONSETTINGS, colorBack),          offsetof(OPTIONSETTINGS, fUseDlgColors), FALSE },
        { _T("Text Color"),             RKI_COLOR, offsetof(OPTIONSETTINGS, colorText),          offsetof(OPTIONSETTINGS, fUseDlgColors), FALSE },
        { _T("Anchor Color"),           RKI_COLOR, offsetof(OPTIONSETTINGS, colorAnchor),        0, FALSE },
        { _T("Anchor Color Visited"),   RKI_COLOR, offsetof(OPTIONSETTINGS, colorAnchorVisited), 0, FALSE },
        { _T("Anchor Color Hover"),     RKI_COLOR, offsetof(OPTIONSETTINGS, colorAnchorHovered), 0, FALSE },
        { _T("Always Use My Colors"),   RKI_BOOL, offsetof(OPTIONSETTINGS, fAlwaysUseMyColors),  0, FALSE },
        { _T("Always Use My Font Size"),RKI_BOOL, offsetof(OPTIONSETTINGS, fAlwaysUseMyFontSize),  0, FALSE },
        { _T("Always Use My Font Face"),RKI_BOOL, offsetof(OPTIONSETTINGS, fAlwaysUseMyFontFace),  0, FALSE },
        { _T("Use Anchor Hover Color"), RKI_BOOL, offsetof(OPTIONSETTINGS, fUseHoverColor),        0, FALSE },

        { _T("Styles"),                 RKI_KEY, 0 },
        { _T("Use My Stylesheet"),      RKI_BOOL, offsetof(OPTIONSETTINGS, fUseMyStylesheet),  0, FALSE },
        { _T("User Stylesheet"),        RKI_STRING, offsetof(OPTIONSETTINGS, cstrUserStylesheet),  offsetof(OPTIONSETTINGS, fUseMyStylesheet), FALSE },
    };

    // keys in windows location
    static TCHAR achWindowsSettingsPath[] = _T("Software\\Microsoft\\Windows\\CurrentVersion");

    static const REGKEYINFORMATION aOptionKeys2[] =
    {
        { _T("Policies\\ActiveDesktop"),RKI_KEY, 0 },
        { _T("Policies"), RKI_KEY, 0 },
    };

    hr = EnsureOptionSettings();
    if(hr)
    {
        goto Cleanup;
    }

    // Make sure we get back the default windows colors, etc.
    if(dwFlags & REGUPDATE_REFRESH)
    {
        _pOptionSettings->SetDefaults();
    }
    ReadSettingsFromRegistry(
        _pOptionSettings->achKeyPath,
        aOptionKeys,
        ARRAYSIZE(aOptionKeys),
        _pOptionSettings,
        dwFlags,
        _pOptionSettings->fSettingsRead,
        (void*)&_pOptionSettings->codepageDefault);

    ReadSettingsFromRegistry(
        achWindowsSettingsPath,
        aOptionKeys2,
        ARRAYSIZE(aOptionKeys2),
        _pOptionSettings, dwFlags,
        _pOptionSettings->fSettingsRead, NULL);

    _pOptionSettings->fSettingsRead = TRUE;

Cleanup:
    RRETURN(hr);
}

//+---------------------------------------------------------------------------
//
//  Member:     CODEPAGESETTINGS::SetDefaults, public
//
//  Synopsis:   Sets the default values for the CODEPAGESETTINGS struct.
//
//----------------------------------------------------------------------------
void CODEPAGESETTINGS::SetDefaults(CODEPAGE codepage, SHORT sOptionSettingsBaselineFontDefault)
{
    fSettingsRead = FALSE;
    bCharSet = DEFAULT_CHARSET;
    sBaselineFontDefault = sOptionSettingsBaselineFontDefault;
    uiFamilyCodePage = codepage;
    latmFixedFontFace = -1;
    latmPropFontFace  = -1;
}

//+---------------------------------------------------------------------------
//
//  Member:     OPTIONSETTINGS::Init, public
//
//  Synopsis:   Initializes data that should only be initialized once.
//
//  Arguments:  [psz] -- String to initialize achKeyPath to. Cannot be NULL.
//
//----------------------------------------------------------------------------
void OPTIONSETTINGS::Init(TCHAR* psz, BOOL fUseCodePageBasedFontLinkingArg)
{
    _tcscpy(achKeyPath, psz);
    fSettingsRead = FALSE;
    fUseCodePageBasedFontLinking = !!fUseCodePageBasedFontLinkingArg;
    sBaselineFontDefault = BASELINEFONTDEFAULT;

    memset(alatmProporitionalFonts, -1, sizeof(alatmProporitionalFonts));
    memset(alatmFixedPitchFonts,    -1, sizeof(alatmFixedPitchFonts));
}

//+---------------------------------------------------------------------------
//
//  Member:     OPTIONSETTINGS::SetDefaults, public
//
//  Synopsis:   Sets the default values of the struct
//
//----------------------------------------------------------------------------
void OPTIONSETTINGS::SetDefaults()
{
    colorBack            = OLECOLOR_FROM_SYSCOLOR(COLOR_WINDOW);
    colorText            = OLECOLOR_FROM_SYSCOLOR(COLOR_WINDOWTEXT);
    colorAnchor          = RGB(0, 0, 0xFF);
    colorAnchorVisited   = RGB(0x80, 0, 0x80);
    colorAnchorHovered   = RGB(0, 0, 0x80);

    fUseDlgColors        = TRUE;
    fExpandAltText       = FALSE;
    fShowImages          = TRUE;
    fShowVideos          = TRUE;
    fPlaySounds          = TRUE;
    fPlayAnimations      = TRUE;
    fUseStylesheets      = TRUE;
    fSmoothScrolling     = TRUE;
    fShowImagePlaceholder= FALSE;
    fShowFriendlyUrl     = FALSE;
    fSmartDithering      = TRUE;
    fAlwaysUseMyColors   = FALSE;
    fAlwaysUseMyFontSize = FALSE;
    fAlwaysUseMyFontFace = FALSE;
    fUseMyStylesheet     = FALSE;
    fUseHoverColor       = FALSE;
    fMoveSystemCaret     = FALSE;
    fHaveAcceptLanguage  = FALSE;
    fCpAutoDetect        = FALSE;

    nAnchorUnderline     = ANCHORUNDERLINE_YES;

    codepageDefault      = (_afxGlobalData._cpDefault==932) ? CP_AUTO_JP : _afxGlobalData._cpDefault;
}

//+---------------------------------------------------------------------------
//
//  Member:     OPTIONSETTINGS::ReadCodePageSettingsFromRegistry
//
//  Synopsis:   Read the fixed and proportional
//  
//----------------------------------------------------------------------------
void OPTIONSETTINGS::ReadCodepageSettingsFromRegistry(CODEPAGESETTINGS* pCS, DWORD dwFlags, SCRIPT_ID sid)
{
    static const REGKEYINFORMATION aScriptBasedFontKeys[] =
    {
        { _T("International\\Scripts\\<0d>"),      RKI_CPKEY, (long)0 },
        { _T("IEFontSize"),                        RKI_SIZE, offsetof(CODEPAGESETTINGS, sBaselineFontDefault),    0, FALSE },
        { _T("IEPropFontName"),                    RKI_FONT, offsetof(CODEPAGESETTINGS, latmPropFontFace),  1, FALSE },
        { _T("IEFixedFontName"),                   RKI_FONT, offsetof(CODEPAGESETTINGS, latmFixedFontFace), 0, FALSE },
    };

    static const REGKEYINFORMATION aCodePageBasedFontKeys[] =
    {
        { _T("International\\<0d>"),               RKI_CPKEY, (long)0 },
        { _T("IEFontSize"),                        RKI_SIZE, offsetof(CODEPAGESETTINGS, sBaselineFontDefault),    0, FALSE },
        { _T("IEPropFontName"),                    RKI_FONT, offsetof(CODEPAGESETTINGS, latmPropFontFace),  1, FALSE },
        { _T("IEFixedFontName"),                   RKI_FONT, offsetof(CODEPAGESETTINGS, latmFixedFontFace), 0, FALSE },
    };

    // NB (cthrash) (CP_UCS_2,sidLatin) is for Unicode documents.  So for OE, pick the Unicode font.
    // (CP_UCS_2,!sidAsciiLatin), on the other hand, is for codepageless fontlinking.  In OE, we obviously
    // can't use codepage-based fontlinking; use instead IE5 fontlinking.
    fUseCodePageBasedFontLinking &= (sid==sidAsciiLatin || sid==sidLatin
        || DefaultCharSetFromScriptAndCharset(sid, DEFAULT_CHARSET)!=DEFAULT_CHARSET);

    DWORD dwArg = fUseCodePageBasedFontLinking ? DWORD(pCS->uiFamilyCodePage) : DWORD(sid);

    ReadSettingsFromRegistry(
        achKeyPath,
        fUseCodePageBasedFontLinking?aCodePageBasedFontKeys:aScriptBasedFontKeys,
        ARRAYSIZE(aCodePageBasedFontKeys),
        pCS,
        dwFlags,
        pCS->fSettingsRead,
        (void*)&dwArg);

    // Determine the appropriate GDI charset
    pCS->bCharSet = DefaultCharSetFromScriptAndCodePage(sid, pCS->uiFamilyCodePage);

    // Do a little fixup on the fonts if not present.  Note that we avoid
    // doing this in CODEPAGESETTINGS::SetDefault as this could be expensive
    // and often unnecessary.
    if(pCS->latmFixedFontFace==-1 || pCS->latmPropFontFace==-1)
    {
        SCRIPTINFO si;
        HRESULT hr;

        hr = MlangGetDefaultFont(sid, &si);

        if(pCS->latmFixedFontFace == -1)
        {
            pCS->latmFixedFontFace = OK(hr)
                ? fc().GetAtomFromFaceName(si.wszFixedWidthFont)
                : 0; // 'System'
        }

        if(pCS->latmPropFontFace == -1)
        {
            pCS->latmPropFontFace = OK(hr)
                ? fc().GetAtomFromFaceName(si.wszProportionalFont)
                : 0; // 'System'
        }
    }

    pCS->fSettingsRead = TRUE;
}