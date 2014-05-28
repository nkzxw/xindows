
#ifndef __XINDOWS_SITE_BASE_SETTING_H__
#define __XINDOWS_SITE_BASE_SETTING_H__

enum REGUPDATE_FLAGS
{
    REGUPDATE_REFRESH = 1,              // always read from registry, not cache
    REGUPDATE_OVERWRITELOCALSTATE = 2   // re-read local state settings as well
};

// Anchor underline options
#define ANCHORUNDERLINE_NO      0x0
#define ANCHORUNDERLINE_YES     0x1
#define ANCHORUNDERLINE_HOVER   0x2

struct CODEPAGESETTINGS
{
    void     Init() { fSettingsRead = FALSE; }
    void     SetDefaults(CODEPAGE codepage, SHORT sOptionSettingsBaselineFontDefault);

    BYTE     fSettingsRead;             // true if we read once from registry
    BYTE     bCharSet;                  // the character set for this codepage
    SHORT    sBaselineFontDefault;      // must be [0-4]
    UINT     uiFamilyCodePage;          // the family codepage for which these settings are for
    LONG     latmFixedFontFace;         // fixed font
    LONG     latmPropFontFace;          // proportional font
};

struct OPTIONSETTINGS
{
    // These operators are defined since we want to allocate extra memory for the string
    //  at the end of the structure.  We cannot simply call MemAlloc directly since we want
    //  to make sure that the constructor for aryCodepageSettings gets called.
    void* operator new(size_t size, size_t extra) { return MemAlloc(size+extra); }
    void operator delete(void* pOS) { MemFree(pOS); }

    void Init(TCHAR* psz, BOOL fUseCodePageBasedFontLinking);
    void SetDefaults();
    void ReadCodepageSettingsFromRegistry(CODEPAGESETTINGS* pCS, DWORD dwFlags, SCRIPT_ID sid);

    COLORREF  crBack()          { return ColorRefFromOleColor(colorBack); }
    COLORREF  crText()          { return ColorRefFromOleColor(colorText); }
    COLORREF  crAnchor()        { return ColorRefFromOleColor(colorAnchor); }
    COLORREF  crAnchorVisited() { return ColorRefFromOleColor(colorAnchorVisited); }
    COLORREF  crAnchorHovered() { return ColorRefFromOleColor(colorAnchorHovered); }

    OLE_COLOR colorBack;
    OLE_COLOR colorText;
    OLE_COLOR colorAnchor;
    OLE_COLOR colorAnchorVisited;
    OLE_COLOR colorAnchorHovered;

    int     nAnchorUnderline;

    //$BUGBUG (dinartem) Make these bit fields in beta2 to save space!
    BYTE    fSettingsRead;
    BYTE    fUseDlgColors;
    BYTE    fExpandAltText;
    BYTE    fShowImages;
    BYTE    fShowVideos;
    BYTE    fPlaySounds;
    BYTE    fPlayAnimations;
    BYTE    fUseStylesheets;
    BYTE    fSmoothScrolling;   // Set to TRUE if the smooth scrolling is allowed
    BYTE    fShowFriendlyUrl;
    BYTE    fSmartDithering;
    BYTE    fAlwaysUseMyColors;
    BYTE    fAlwaysUseMyFontSize;
    BYTE    fAlwaysUseMyFontFace;
    BYTE    fUseMyStylesheet;
    BYTE    fUseHoverColor;
    BYTE    fMoveSystemCaret;
    BYTE    fHaveAcceptLanguage; // true if cstrLang is valid
    BYTE    fCpAutoDetect;       // true if cp autodetect mode is set
    BYTE    fShowImagePlaceholder;
    BYTE    fUseCodePageBasedFontLinking;

    SHORT   sBaselineFontDefault;

    CString cstrUserStylesheet;
    CODEPAGE codepageDefault;   // codepage of last resort

    CString cstrLang;

    DECLARE_CPtrAry(OSCodePageAry, CODEPAGESETTINGS*)

    // We keep an array for the codepage settings here
    OSCodePageAry aryCodepageSettingsCache;

    // Script-based font info.  Array of LONG atoms.
#ifndef NO_UTF16
    LONG alatmProporitionalFonts[sidLimPlusSurrogates];
    LONG alatmFixedPitchFonts[sidLimPlusSurrogates];
#else
    LONG alatmProporitionalFonts[sidLim];
    LONG alatmFixedPitchFonts[sidLim];
#endif

    // The following member must be the last member in the struct.
    TCHAR achKeyPath[1];
};

#endif //__XINDOWS_SITE_BASE_SETTING_H__