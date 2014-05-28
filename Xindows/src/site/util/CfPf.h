
#ifndef __XINDOWS_SITE_UTIL_CFPF_H__
#define __XINDOWS_SITE_UTIL_CFPF_H__

struct OPTIONSETTINGS;
struct CODEPAGESETTINGS;

#define ISBORDERSIDECLRSETUNIQUE(pFF,iSide) (pFF->_bBorderColorsSetUnique&(1<<iSide))
#define SETBORDERSIDECLRUNIQUE(pFF,iSide)   (pFF->_bBorderColorsSetUnique |= (1<<iSide));

// Border sides.  These are used to make loops of common set/get operations on border properties.
enum
{
    BORDER_TOP      = 0,
    BORDER_RIGHT    = 1,
    BORDER_BOTTOM   = 2,
    BORDER_LEFT     = 3
};

#define BASELINEFONTMIN         0
#define BASELINEFONTMAX         4
#define BASELINEFONTDEFAULT     2

#define MAX_TAB_STOPS           1
#define lDefaultTab             960

// This enum needs to be kept in sync with the htmlBlockAlign & xindowsControlAlign enums
enum htmlAlign
{
    htmlAlignNotSet = 0,
    htmlAlignLeft = 1,
    htmlAlignCenter = 2,
    htmlAlignRight = 3,
    htmlAlignTextTop = 4,
    htmlAlignAbsMiddle = 5,
    htmlAlignBaseline = 6,
    htmlAlignAbsBottom = 7,
    htmlAlignBottom = 8,
    htmlAlignMiddle = 9,
    htmlAlignTop = 10,
    htmlAlign_Last_Enum
};

int ConvertHtmlSizeToTwips(int nHtmlSize);
int ConvertTwipsToHtmlSize(int nFontSize);


// DECLARE_SPECIAL_OBJECT_FLAGS
//
// These CF flags are used for text runs that lie within special line
// services object boundaries.
//
//      _fIsRuby:        true if this is part of a ruby object
//      _fIsRubyText:    true if this is part of the ruby pronunciation text
//      _fBidiEmbed:     true if this is a directional embedding
//      _fBidiOverride:  true if this is a directional override
#define DECLARE_SPECIAL_OBJECT_FLAGS()                      \
    union                                                   \
    {                                                       \
        BYTE _bSpecialObjectFlagsVar;                       \
        struct                                              \
        {                                                   \
            unsigned char _fIsRuby                 : 1;     \
            unsigned char _fIsRubyText             : 1;     \
            unsigned char _fBidiEmbed              : 1;     \
            unsigned char _fBidiOverride           : 1;     \
            unsigned char _fNoBreak                : 1;     \
            unsigned char _fNoBreakInner           : 1;     \
                                                            \
            unsigned char _fSpecialObjectUnused    : 2;     \
        };                                                  \
    };                                                      \
    BYTE& _bSpecialObjectFlags() { return _bSpecialObjectFlagsVar; }

/*
*  CCharFormat
*
*  @class
*      Collects related CHARFORMAT methods and inherits from CHARFORMAT2
*
*  @devnote
*      Could add extra data for round tripping RTF and TOM info, e.g.,
*      save style handles. This data wouldn't be exposed at the API level.
*/
class CCharFormat
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    // (srinib) - WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //            WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //            WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //
    // font specific flags should follow _wFlags, ComputeFontCrc
    // depends on it, otherwise we screw up font caching
    union
    {
        WORD _wFlagsVar; // For now use _wFlags()
        struct
        {
            unsigned short _fUnderline        : 1;
            unsigned short _fOverline         : 1;
            unsigned short _fStrikeOut        : 1;
            unsigned short _fNarrow           : 1; // for brkcls

            unsigned short _fVisibilityHidden : 1;
            unsigned short _fDisplayNone      : 1; // display nothing if set
            unsigned short _fDisabled         : 1;
            unsigned short _fAccelerator      : 1;

            unsigned short _fHasBgColor       : 1;
            unsigned short _fHasBgImage       : 1;
            unsigned short _fRelative         : 1; // relatively positioned chunk
            unsigned short _fExplicitFace     : 1; // font face set from font attr

            unsigned short _fRTL              : 1; // right to left direction

            unsigned short _fUnused0          : 3;
        };
    };
    WORD& _wFlags() { return _wFlagsVar; }

    // (srinib) - WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //            WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //            WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //
    // Look at the the comments above before adding any bits here
    union
    {
        WORD _wFontSpecificFlagsVar; // For now use _wFontSpecificFlags()
        struct
        {
            unsigned short _fBold           : 1;
            unsigned short _fItalic         : 1;
            unsigned short _fSuperscript    : 1;
            unsigned short _fSubscript      : 1;

            unsigned short _fBumpSizeDown   : 1;
            unsigned short _fPassword       : 1;
            unsigned short _fProtected      : 1;
            unsigned short _fSizeDontScale  : 1;

            unsigned short _fDownloadedFont : 1;
            unsigned short _fSubSuperSized  : 1;
            unsigned short _fUnused         : 5;

        };
    };
    WORD& _wFontSpecificFlags() { return _wFontSpecificFlagsVar; }

    // (srinib) - WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //            WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //            WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //
    // Look at the the comments below before adding any members below
    WORD        _wWeight;            // Font weight (LOGFONT value)
    WORD        _wKerning;           // Twip size above which to kern char pair
    LONG        _yHeight;            // This will always be in twips
    LONG        _yOffset;            // > 0 for superscript, < 0 for subscript
    LCID        _lcid;               // Locale ID
    LONG        _latmFaceName;       // was _szFaceName.  ref's CAtomTable in g_fontcache.
    BYTE        _bCharSet;
    BYTE        _bPitchAndFamily;

    // (srinib) - WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //            WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //            WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:WARNING:
    //
    // ComputeFontCrc depends on the offset of _bCursorIdx, members below _bCursorIdx
    // are not font specific. So if you are adding anything to the CF and is font related
    // add it above _bCursorIdx or else add below. It is necessary otherwise we create
    // many fonts even if they are same.
    BYTE        _bCursorIdx;         // CssAttribute. enum for Cursor type
    BYTE        _bTextTransform  :4;
    BYTE        _fTextAutospace  :4; // CSS text-autospace
    CUnitValue  _cuvLetterSpacing;   // CssAttribute for space between characters
    CUnitValue  _cuvLineHeight;      // CssAttribute for line height multiplier
    CColorValue _ccvTextColor;       // text color

    // Please see above declaration for details.  This currently
    // takes up a BYTE of space
    DECLARE_SPECIAL_OBJECT_FLAGS();

    BYTE    _fBranchFiltered        :1; // Element or ancestor has a filter
    BYTE    _fLineBreakStrict       :1; // CSS line-break == strict?
    BYTE    _fCharGridSizeOverride  :1; // Determines whether paragraf grid char size is overriden
    BYTE    _fLineGridSizeOverride  :1; // Determines whether paragraf grid line size is overriden
    BYTE    _uLayoutGridType        :2; // Outer layout grid type
    BYTE    _uLayoutGridTypeInner   :2; // Inner layout grid type
    BYTE    _uLayoutGridMode        :3; // Outer layout grid mode
    BYTE    _uLayoutGridModeInner   :3; // Inner layout grid mode
    BYTE    _fHasGridValues         :1;

    BYTE    _fHasDirtyInnerFormats  :1; // CAUTION: add any flags above here.

    // The Compare function assumes that all simple types are placed before
    // the _bCrcFont member, & all "complex" ones after
    BYTE    _bCrcFont;                  // used to cache the font specific crc
                                        // set by ComputeFontCrc, and must be cached
                                        // before caching it in the global cache
                                        // any members below this are not used in crc computation
public:
    CCharFormat()
    {
        _wFlags() = 0;
        _wFontSpecificFlags() = 0;
    }
    CCharFormat(const CCharFormat& cf) { *this = cf; }
    CCharFormat& operator=(const CCharFormat& cf)
    {
        memcpy(this, &cf, sizeof(*this));
        return *this;
    }

    // Clone this paraformat
    HRESULT Clone(CCharFormat** ppCF)
    {
        Assert(ppCF);
        *ppCF = new CCharFormat(*this);
        MemSetName((*ppCF, "cloned CCharFormat"));
        return *ppCF?S_OK:E_OUTOFMEMORY;
    }

    BOOL    AreInnerFormatsDirty();
    void    ClearInnerFormats();

    // Compare two CharFormat
    BOOL    Compare(const CCharFormat* pCF) const;
    BOOL    CompareForLayout (const CCharFormat* pCF) const;
    BOOL    CompareForLikeFormat(const CCharFormat* pCF) const;

    TCHAR*  GetFaceName() const // replaces _szFaceName
    {
        // must cast away the const-ness.
        return (TCHAR*)fc().GetFaceNameFromAtom(_latmFaceName);
    }

    void SetFaceName(const TCHAR* szFaceName) // mutator
    {
        _latmFaceName= fc().GetAtomFromFaceName(szFaceName);
    }

    const TCHAR* GetFamilyName() const;

    // Compute and return CharFormat Crc
    WORD ComputeCrc() const;

    BYTE ComputeFontCrc() const;

    // Initialization routines
    HRESULT InitDefault(HFONT hfont);
    HRESULT InitDefault(OPTIONSETTINGS* pOS, CODEPAGESETTINGS* pCS, BOOL fKeepFaceIntact=FALSE);

    void SetHeightInTwips(LONG twips);
    void SetHeightInNonscalingTwips(LONG twips);
    void ChangeHeightRelative(int diff);
    LONG GetHeightInTwips(CDocument* pDoc) const;
    LONG GetHeightInPixels(CDocInfo* pdci);

    BOOL SwapSelectionColors() const;

    BOOL IsDisplayNone() const { return _fDisplayNone; }
    BOOL IsVisibilityHidden() const { return _fVisibilityHidden; }

    TCHAR PasswordChar(TCHAR ch) const { return _fPassword?ch:0; }

    // Return TRUE if a text transform is needed.
    BOOL IsTextTransformNeeded() const
    {
        return ((_bTextTransform!=styleTextTransformNotSet)&&(_bTextTransform!=styleTextTransformNone));
    }

    BOOL HasBgColor(BOOL fInner) const { return fInner?FALSE:_fHasBgColor; }
    BOOL HasBgImage(BOOL fInner) const { return fInner?FALSE:_fHasBgImage; }
    BOOL IsRelative(BOOL fInner) const { return fInner?FALSE:_fRelative;   }
    BOOL HasNoBreak(BOOL fInner) const { return fInner?_fNoBreakInner:_fNoBreak; }

    styleLayoutGridMode GetLayoutGridMode(BOOL fInner) const 
    {
        return (styleLayoutGridMode)(fInner?_uLayoutGridModeInner:_uLayoutGridMode);
    }
    styleLayoutGridType GetLayoutGridType(BOOL fInner) const 
    {
        return (styleLayoutGridType)(fInner?_uLayoutGridTypeInner:_uLayoutGridType);
    }
    BOOL HasCharGridSizeOverrriden() const { return _fCharGridSizeOverride; }
    BOOL HasLineGridSizeOverrriden() const { return _fLineGridSizeOverride; }
    BOOL HasCharGrid(BOOL fInner) const 
    { 
        return (styleLayoutGridModeChar&(fInner?_uLayoutGridModeInner:_uLayoutGridMode)); 
    }
    BOOL HasLineGrid(BOOL fInner) const 
    {
        return (styleLayoutGridModeLine&(fInner?_uLayoutGridModeInner:_uLayoutGridMode)); 
    }
};

struct CListing
{
public:
    enum LISTING_CONSTS { MAXLEVELS = 256 };

    enum LISTING_TYPE
    {
        NONE = 0,
        BULLET = 1,
        NUMBERING = 2,
        IMAGE = 3,
        DEFINITION = 4,
        LAST
    };

private:
    WORD wBits;
    WORD wStyle;

    enum LISTING_MASKS
    {
        TYPE       = 0x0007,       // bits 0-2
        GETVALUE   = 0x0008,       // bit 3
        GETSTYLE   = 0x0010,       // bit 4
        INLIST     = 0x0020,       // bit 5
        UNUSED     = 0x00c0,       // bits 6-7
        VALUE      = 0xff00        // bits 8-15
    };

public:
    void Reset() { wBits = wStyle = 0; }
    BOOL IsReset() const { return(wBits==0 && wStyle==0); }

    LISTING_TYPE GetType() const { return (LISTING_TYPE)(wBits&TYPE); }
    void SetType(LISTING_TYPE lt)
    {
        Assert(lt>0 && lt<LAST);
        wBits &= ~TYPE;
        wBits |= (WORD)lt;
    }

    int GetLevel() const { return (int)(wBits>>8); }
    void SetLevel(int level)
    {
        Assert(level>=0 && level<256);
        wBits &= ~VALUE;
        wBits |= (WORD)level<<8;
    }

    BOOL HasAdornment() const
    {
        int type = (wBits & TYPE);

        return type == BULLET||type==NUMBERING||type==IMAGE;
    }

    WORD IsInList() const     { return wBits&INLIST; }
    void SetInList()          { wBits |= INLIST; }
    void SetNotInList()       { wBits &= ~INLIST; }
    WORD IsValueValid() const { return wBits&GETVALUE; }
    void SetValueValid()      { wBits |= GETVALUE; }
    WORD IsStyleValid() const { return wBits&GETSTYLE; }
    void SetStyleValid()      { wBits |= GETSTYLE; }

    enum _styleListStyleType GetStyle() const
    {
        return (enum _styleListStyleType)wStyle;
    }
    void SetStyle(enum _styleListStyleType style) { wStyle = (WORD)style; }
};

/*
*  CParaFormat
*
*  @class
*      Collects related PARAFORMAT methods and inherits from PARAFORMAT2
*
*  @devnote
*      Could add extra data for round tripping RTF and TOM info, e.g., to
*      save style handles
*/
// Space between paragraphs.
#define DEFAULT_VERTICAL_SPACE_POINTS   14
#define DEFAULT_VERTICAL_SPACE_TWIPS    TWIPS_FROM_POINTS(DEFAULT_VERTICAL_SPACE_POINTS)

// Amount to indent for a blockquote.
#define BLOCKQUOTE_INDENT_POINTS        30
#define BLOCKQUOTE_INDENT_TWIPS         TWIPS_FROM_POINTS(BLOCKQUOTE_INDENT_POINTS)

// Amount to indent for lists.
#define LIST_INDENT_POINTS              30
#define LIST_FIRST_REDUCTION_POINTS     12
#define LIST_INDENT_TWIPS               TWIPS_FROM_POINTS(LIST_INDENT_POINTS)
#define LIST_FIRST_REDUCTION_TWIPS      TWIPS_FROM_POINTS(LIST_FIRST_REDUCTION_POINTS)

class CParaFormat
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    union
    {
        DWORD _dwFlagsVar; // For now use _dwFlags()
        struct
        {
            unsigned _bBlockAlignInner      : 3;   // Inner block align
            unsigned _bListPosition         : 3;   // hanging bullet?

            unsigned _fRTL                  : 1;   // COMPLEXSCRIPT paragraph is right to left
            unsigned _fTabStops             : 1;
            unsigned _fCompactDL            : 1;
            unsigned _fResetDLLevel         : 1;

            unsigned _fPadBord              : 1;   // Someone above me has borders or padding set.
            unsigned _fPre                  : 1;

            unsigned _fInclEOLWhite         : 1;
            unsigned _fPreInner             : 1;
            unsigned _fInclEOLWhiteInner    : 1;

            unsigned _fRTLInner             : 1;
            unsigned _fFirstLineIndentForDD : 1;

            unsigned _fWordBreak            : 2;   // For CSS word-break
            unsigned _fWordBreakInner       : 2;   // For CSS word-break

            unsigned _uTextJustify          : 3;
            unsigned _uTextJustifyTrim      : 2;

            unsigned _fUnused               : 5;

            // add any new flags above here (srinib)
            unsigned _fHasDirtyInnerFormats : 1;

        };
    };

    DWORD& _dwFlags() { return _dwFlagsVar; }

    BYTE        _bBlockAlign;               // Block Alignment L,R,C alignemnt
    BYTE        _bTableVAlignment;          // new for tables
    SHORT       _cTabCount;
    CListing    _cListing;                  // info for listing
    LONG        _lNumberingStart;           // Starting value for numbering
    CUnitValue  _cuvTextIndent;             // points value of first line left indent
    CUnitValue  _cuvLeftIndentPoints;       // accum. points value of left indent
    CUnitValue  _cuvLeftIndentPercent;      // accum. percent value of left indent
    CUnitValue  _cuvRightIndentPoints;      // accum. points value of right indent
    CUnitValue  _cuvRightIndentPercent;     // accum. percent value of right indent
    CUnitValue  _cuvNonBulletIndentPoints;  // accum. points value of non bullet left indent
    CUnitValue  _cuvNonBulletIndentPercent; // accum. percent value of non bullet left indent
    CUnitValue  _cuvOffsetPoints;           // accum. points value of bullet offset
    CUnitValue  _cuvOffsetPercent;          // accum. percent value of bullet offset
    CUnitValue  _cuvCharGridSize;           // outer character grid size for layout-grid-char
    CUnitValue  _cuvCharGridSizeInner;      // inner character grid size for layout-grid-char
    CUnitValue  _cuvLineGridSize;           // inner line grid size for layout-grid-line
    CUnitValue  _cuvLineGridSizeInner;      // outer line grid size for layout-grid-line
    CUnitValue  _cuvTextKashida;            // CSS text-kashida

    LONG        _lImgCookie;                // image cookie for bullet
    LONG        _lFontHeightTwips;          // if UV's are in EMs we need the fontheight.

    // Warning: Once again, the simple types appear before the rgxTabs
    LONG        _rgxTabs[MAX_TAB_STOPS];

    CParaFormat()
    {
        _lFontHeightTwips = 1; // one because its used in multiplication in CUnitValue::Converto
    }

    CParaFormat(const CParaFormat& pf) { *this = pf; }

    CParaFormat& operator=(const CParaFormat& pf)
    {
        memcpy(this, &pf, sizeof(*this));
        return *this;
    }

    // Clone this paraformat
    HRESULT Clone(CParaFormat** ppPF)
    {
        Assert(ppPF);
        *ppPF = new CParaFormat(*this);
        MemSetName((*ppPF, "cloned CParaFormat"));
        return *ppPF?S_OK:E_OUTOFMEMORY;
    }

    BOOL AreInnerFormatsDirty();
    void ClearInnerFormats();

    // Compare this ParaFormat with passed one
    BOOL Compare(const CParaFormat* pPF) const;

    // Compute and return Crc for this ParaFormat
    WORD ComputeCrc() const;

    // Tab management
    HRESULT AddTab(LONG tabPos, LONG tabType, LONG tabStyle);

    HRESULT DeleteTab(LONG tabPos);

    HRESULT GetTab(long iTab, long* pdxptab, long* ptbt, long* pstyle) const;

    LONG GetTabPos(LONG tab) const { return tab&0xffffff; }

    // Initialization
    HRESULT InitDefault();

    LONG GetTextIndent(CCalcInfo* pci) const
    {
        return _cuvTextIndent.XGetPixelValue(pci, pci->_sizeParent.cx, _lFontHeightTwips);
    }

    LONG GetLeftIndent(CParentInfo* ppri, BOOL fInner) const
    {
        return fInner ? 0
            : (_cuvLeftIndentPoints.XGetPixelValue(ppri, 0, _lFontHeightTwips)+
            (_cuvLeftIndentPercent.IsNull()?0:_cuvLeftIndentPercent.XGetPixelValue(ppri, ppri->_sizeParent.cx, _lFontHeightTwips)));
    }

    LONG GetRightIndent(CParentInfo* ppri, BOOL fInner) const
    {
        return fInner ? 0
            : (_cuvRightIndentPoints.XGetPixelValue(ppri, 0, _lFontHeightTwips)+
            (_cuvRightIndentPercent.IsNull()?0:_cuvRightIndentPercent.XGetPixelValue(ppri, ppri->_sizeParent.cx, _lFontHeightTwips)));
    }
    LONG GetBulletOffset(CParentInfo* ppri, BOOL fInner) const
    {
        return (fInner||(styleListStylePosition(_bListPosition)==styleListStylePositionInside)) ? 0
            : (_cuvOffsetPoints.XGetPixelValue(ppri, 0, _lFontHeightTwips)+
            (_cuvOffsetPercent.IsNull()?0:_cuvOffsetPercent.XGetPixelValue(ppri, ppri->_sizeParent.cx, _lFontHeightTwips)));
    }

    LONG GetNonBulletIndent(CParentInfo* ppri, BOOL fInner) const
    {
        return fInner ? 0
            : (_cuvNonBulletIndentPoints.XGetPixelValue(ppri, 0, _lFontHeightTwips)+
            (_cuvNonBulletIndentPercent.IsNull()?0:_cuvNonBulletIndentPercent.XGetPixelValue(ppri, ppri->_sizeParent.cx, _lFontHeightTwips)));
    }

    BOOL HasTabStops     (BOOL fInner) const { return fInner ? FALSE : _fTabStops;     }
    BOOL HasCompactDL    (BOOL fInner) const { return fInner ? FALSE : _fCompactDL;    }
    BOOL HasResetDLLevel (BOOL fInner) const { return fInner ? FALSE : _fResetDLLevel; }
    BOOL HasPadBord      (BOOL fInner) const { return fInner ? FALSE : _fPadBord;      }
    BOOL HasRTL          (BOOL fInner) const { return fInner ? _fRTLInner : _fRTL;     }
    BOOL HasPre          (BOOL fInner) const { return fInner ? _fPreInner : _fPre;     }
    BOOL HasInclEOLWhite (BOOL fInner) const { return fInner ? _fInclEOLWhiteInner : _fInclEOLWhite; }

    SHORT GetTabCount       (BOOL fInner) const { return fInner ? 0 : _cTabCount;       }
    long  GetNumberingStart (BOOL fInner) const { return fInner ? 0 : _lNumberingStart; }
    long  GetImgCookie      (BOOL fInner) const { return fInner ? 0 : _lImgCookie;      }

    styleListStylePosition GetListStylePosition (BOOL fInner) const
    {
        return fInner?styleListStylePositionNotSet:styleListStylePosition(_bListPosition);
    }

    htmlBlockAlign GetBlockAlign(BOOL fInner) const
    {
        return (htmlBlockAlign)(fInner?_bBlockAlignInner:_bBlockAlign);
    }

    CListing GetListing(BOOL fInner) const
    {
        if(fInner)
        {
            CListing cl;
            cl.Reset();
            return cl;
        }
        else
        {
            return _cListing;
        }
    }

    const CUnitValue& GetCharGridSize(BOOL fInner) const 
    {
        return (fInner?_cuvCharGridSizeInner:_cuvCharGridSize);
    }
    const CUnitValue& GetLineGridSize(BOOL fInner) const 
    {
        return (fInner?_cuvLineGridSizeInner:_cuvLineGridSize);
    }
};


/*
*  CFancyFormat
*
*  @class
*      Collects related CSite
*
*/
class CFancyFormat
{
public:
    DECLARE_MEMALLOC_NEW_DELETE();

    union
    {
        DWORD _dwFlagsVar; // _dwFlags()
        struct
        {
            unsigned int _fBgRepeatX             : 1;
            unsigned int _fBgRepeatY             : 1;
            unsigned int _fBgFixed               : 1;
            unsigned int _fRelative              : 1;   // relatively positioned element

            unsigned int _bBorderColorsSetUnique : 4;   // Has border-color explicitly been set on particular sides?
                                                        // See ISBORDERSIDECLRSETUNIQUE() and SETBORDERSIDECLRUNIQUE() macros above.
                                                        // These flags are used to determine when to compute
                                                        // light/highlight/dark/shadow colors.

            unsigned int _bBorderSoftEdges       : 1;   // Corresponds to the BF_SOFT flag on wEdges in the BORDERINFO struct.
            unsigned int _fExplicitTopMargin     : 1;   // Explicit margin overrides BODY margin.
            unsigned int _fExplicitBottomMargin  : 1;   // Explicit margin overrides BODY margin.
            unsigned int _fBlockNess             : 1;   // has blockness (element is block element).

            unsigned int _fHasLayout             : 1;   // has 'layoutness'
            unsigned int _fCtrlAlignFromCSS      : 1;   // Control alignment from CSS instead of HTML attr.
            unsigned int _fHeightPercent         : 1;   // Height is a percent.
            unsigned int _fWidthPercent          : 1;   // Width is a percent.

            unsigned int _fAlignedLayout         : 1;
            unsigned int _bTableLayout           : 1;   // table-layout property: fixed|auto
            unsigned int _bBorderCollapse        : 1;   // border-collapse property: collapse|separate  (used by tables)
            unsigned int _fOverrideTablewideBorderDefault : 1;  // set by other table elements when the override border properties,
                                                                // cleared by the table, read by tcell.
            unsigned int _fPositioned            : 1;   // set if the element is positioned
            unsigned int _fZParent               : 1;   // set if the element is a ZParent
            unsigned int _fAutoPositioned        : 1;   // set if the element is auto positioned (relative or absolute with top/left auto)
            unsigned int _fScrollingParent       : 1;   // set if the element is a scrolling parent

            unsigned int _fClearLeft             : 1;
            unsigned int _fClearRight            : 1;
            unsigned int _fPercentHorzPadding    : 1;
            unsigned int _fPercentVertPadding    : 1;

            unsigned int _fNoScroll              : 1;
            unsigned int _fHasMargins            : 1;
            unsigned int _fHasExpressions        : 1;
            unsigned int _fHasNoWrap             : 1;
        };
    };
    DWORD& _dwFlags() { return _dwFlagsVar; }

    CUnitValue      _cuvPaddingTop;
    CUnitValue      _cuvPaddingRight;
    CUnitValue      _cuvPaddingBottom;
    CUnitValue      _cuvPaddingLeft;
    CUnitValue      _cuvSpaceBefore;            // Vertical spacing before para
    CUnitValue      _cuvSpaceAfter;             // Vertical spacing after para
    CUnitValue      _cuvWidth;
    CUnitValue      _cuvHeight;
    CColorValue     _ccvBorderColors[4];
    CUnitValue      _cuvBorderWidths[4];
    BYTE            _bBorderStyles[4];          // fmControlBorderStyles
    CColorValue     _ccvBorderColorLight;
    CColorValue     _ccvBorderColorDark;
    CColorValue     _ccvBorderColorHilight;
    CColorValue     _ccvBorderColorShadow;
    CUnitValue      _cuvClipTop;
    CUnitValue      _cuvClipRight;
    CUnitValue      _cuvClipBottom;
    CUnitValue      _cuvClipLeft;
    CUnitValue      _cuvMarginTop;              // Margin on top
    CUnitValue      _cuvMarginRight;            // Margin on right
    CUnitValue      _cuvMarginBottom;           // Margin on bottom
    CUnitValue      _cuvMarginLeft;             // Margin on left

    LONG            _lZIndex;                   // CSS-P: z-index
    CUnitValue      _cuvTop;                    // top of absolutely/relatively positioned chunk
    CUnitValue      _cuvBottom;                 // top of absolutely/relatively positioned chunk
    CUnitValue      _cuvLeft;                   // left of absolutely/relatively positioned chunk
    CUnitValue      _cuvRight;                  // right of absolutely/relatively positioned chunk

    CColorValue     _ccvBackColor;              // background color on the element
    CUnitValue      _cuvBgPosX;
    CUnitValue      _cuvBgPosY;
    styleListStyleType _ListType;

    LONG            _lImgCtxCookie;             // background image context cookie in
                                                // the doc's bgUrl-imgCtx cache
    WORD            _wPadding;
    SHORT           _iExpandos;                 // index to cached attrarray that contains style expandos affecting us

    BYTE            _bStyleFloat;
    BYTE            _bControlAlign;             // Control Alignment L,R,C,absTop, absBottom
    BYTE            _bOverflowX;
    BYTE            _bOverflowY;

    BYTE            _bPositionType;             // CSS-P: static, relative, or absolute
    BYTE            _bDisplay;                  // Display property
    BYTE            _bVisibility;               // Visibility property


    // WARNING: This MUST be the LAST item in the FF!!
    // The Compare function assumes that all simple types (that can be compared by memcmp)
    // are placed before the _pszFilters member, & all "complex" ones after
    TCHAR*          _pszFilters;                // Multimedia filters (CSS format string)

    CFancyFormat();
    ~CFancyFormat();
    CFancyFormat(const CFancyFormat&ff);

    CFancyFormat& Copy(const CFancyFormat& ff);

    CFancyFormat& operator=(const CFancyFormat& ff);

    // Clone this FancyFormat
    HRESULT Clone(CFancyFormat** ppFF)
    {
        Assert(ppFF);
        *ppFF = new CFancyFormat(*this);
        MemSetName((*ppFF, "cloned CFancyFormat"));
        return *ppFF?S_OK:E_OUTOFMEMORY;
    }

    // Compare this FancyFormat with passed one
    BOOL    Compare(const CFancyFormat* pFF) const;

    // Compute and return Crc for this FancyFormat
    WORD    ComputeCrc() const;
    void    InitDefault();

    BOOL    ElementNeedsFlowLayout() const;

    BOOL    IsTopAuto()         const { return _cuvTop.IsNullOrEnum(); }
    BOOL    IsLeftAuto()        const { return _cuvLeft.IsNullOrEnum(); }
    BOOL    IsRightAuto()       const { return _cuvRight.IsNullOrEnum(); }
    BOOL    IsBottomAuto()      const { return _cuvBottom.IsNullOrEnum(); }
    BOOL    IsWidthAuto()       const { return _cuvWidth.IsNullOrEnum(); }
    BOOL    IsHeightAuto()      const { return _cuvHeight.IsNullOrEnum(); }
    BOOL    IsAbsolute()        const { return stylePosition(_bPositionType)==stylePositionabsolute; }
    BOOL    IsRelative()        const { return stylePosition(_bPositionType)==stylePositionrelative; };
    BOOL    IsPositionStatic()  const { return !_fPositioned; }
    BOOL    IsPositioned()      const { return _fPositioned; }
    BOOL    IsAligned()         const { return _fAlignedLayout; }
    BOOL    IsAutoPositioned()  const { return _fAutoPositioned; }
    BOOL    IsZParent()         const { return _fZParent; }
};



inline void ApplyDefaultBeforeSpace(CFancyFormat* pFF, int defPoints=-1)
{
    if(!pFF->_fExplicitTopMargin)
    {
        pFF->_cuvSpaceBefore.SetPoints(defPoints==-1?DEFAULT_VERTICAL_SPACE_POINTS:defPoints);
    }
}
inline void ApplyDefaultAfterSpace(CFancyFormat* pFF, int defPoints=-1)
{
    // This newest tag in the tree determines post para spacing.
    if(!pFF->_fExplicitBottomMargin)
    {
        pFF->_cuvSpaceAfter.SetPoints(defPoints==-1?DEFAULT_VERTICAL_SPACE_POINTS:defPoints);
    }
}
inline void ApplyDefaultVerticalSpace(CFancyFormat* pFF, int defPoints=-1)
{
    ApplyDefaultBeforeSpace(pFF, defPoints);
    ApplyDefaultAfterSpace(pFF, defPoints);
}



class CStyleInfo;
class CFormatInfo;
class CBehaviorInfo;

enum ApplyPassType { APPLY_All, APPLY_ImportantOnly, APPLY_NoImportant, APPLY_Behavior };

HRESULT ApplyAttrArrayValues(
          CStyleInfo* pStyleInfo,
          CAttrArray** ppAA,
          ApplyPassType passType=APPLY_All,
          BOOL* pfContainsImportant=NULL,
          BOOL fApplyExpandos=TRUE,
          DISPID dispidSetValue=0);
HRESULT ApplyFormatInfoProperty(
         const PROPERTYDESC* pPropertyDesc,
         DISPID dispID,
         VARIANT varValue,
         CFormatInfo* pCFI);

HRESULT ApplyBehaviorProperty(CAttrArray* pAA, CBehaviorInfo* pBehaviorInfo);

void ApplySiteAlignment(CFormatInfo* pCFI, htmlControlAlign at, CElement* pElem);

#endif //__XINDOWS_SITE_UTIL_CFPF_H__