
#ifndef __XINDOWS_SITE_BASE_FORMATINFO_H__
#define __XINDOWS_SITE_BASE_FORMATINFO_H__

class CTreeNode;

inline const CCharFormat* GetCharFormatEx(long iCF)
{
    const CCharFormat* pCF = &(*TLS(_pCharFormatCache))[iCF];
    Assert(pCF);
    return pCF;
}

inline const CParaFormat* GetParaFormatEx(long iPF)
{
    const CParaFormat* pPF = &(*TLS(_pParaFormatCache))[iPF];
    Assert(pPF);
    return pPF;
}

inline const CFancyFormat* GetFancyFormatEx(long iFF)
{
    const CFancyFormat* pFF = &(*TLS(_pFancyFormatCache))[iFF];
    Assert(pFF);
    return pFF;
}

inline CAttrArray* GetExpandosAttrArrayFromCacheEx(long iStyleAttrArray)
{
    const CAttrArray* pExpandosAttrArray = &(*TLS(_pStyleExpandoCache))[iStyleAttrArray];
    Assert(pExpandosAttrArray);
    return (CAttrArray*)pExpandosAttrArray;
}

//+----------------------------------------------------------------
//
//  Class:      CStyleInfo
//
//
//-----------------------------------------------------------------
class CStyleInfo
{
public:
    CStyleInfo() {}
    CStyleInfo(CTreeNode* pNodeContext) { _pNodeContext = pNodeContext; }
    CTreeNode* _pNodeContext;
};

//+----------------------------------------------------------------------------
//
//  Class:      CBehaviorInfo
//
//
//-----------------------------------------------------------------------------
class CBehaviorInfo : public CStyleInfo
{
public:
    CBehaviorInfo(CTreeNode* pNodeContext) : CStyleInfo(pNodeContext)
    {
    }
    ~CBehaviorInfo()
    {
        _acstrBehaviorUrls.Free();
    }

    CAtomTable _acstrBehaviorUrls; // urls of behaviors
};

typedef enum
{
    ComputeFormatsType_Normal = 0x0,
    ComputeFormatsType_GetValue = 0x1,
    ComputeFormatsType_GetInheritedValue = -1,  // 0xFFFFFFFF
    ComputeFormatsType_ForceDefaultValue = 0x2,
} COMPUTEFORMATSTYPE;

//+--------------------------------------------------------------------
//
//  Class:      CFormatInfo
//
//   Note:      This class exists for optimization purposes only.
//              This is so we only have to pass 1 object on the stack
//              while we are applying format.
//
//   WARNING    This class is allocated on the stack so you cannot count on
//              the new operator to clear things out for you.
//
//---------------------------------------------------------------------
class CFormatInfo : public CStyleInfo
{
private:
    DECLARE_MEMALLOC_NEW_DELETE();
    CFormatInfo();

public:

    void            Reset() { memset(this, 0, offsetof(CFormatInfo, _icfSrc)); }
    void            Cleanup();
    CAttrArray*     GetAAExpando();
    void            PrepareCharFormatHelper();
    void            PrepareParaFormatHelper();
    void            PrepareFancyFormatHelper();
    HRESULT         ProcessImgUrl(CElement* pElem, LPCTSTR lpszUrl, DISPID dispID, LONG* plCookie, BOOL fHasLayout);

    void            UnprepareForDebug() {}
    void            PrepareCharFormat() { if(!_fPreparedCharFormat) PrepareCharFormatHelper(); }
    void            PrepareParaFormat() { if(!_fPreparedParaFormat) PrepareParaFormatHelper(); }
    void            PrepareFancyFormat(){ if(!_fPreparedFancyFormat) PrepareFancyFormatHelper(); }
    CCharFormat&    _cf()               { return(_cfDst); }
    CParaFormat&    _pf()               { return(_pfDst); }
    CFancyFormat&   _ff()               { return(_ffDst); }

    unsigned                _fPreparedCharFormat    :1; //  0
    unsigned                _fPreparedParaFormat    :1; //  1
    unsigned                _fPreparedFancyFormat   :1; //  2
    unsigned                _fPreparedAAExpando     :1; //  3
    unsigned                _fHasImageUrl           :1; //  4
    unsigned                _fHasBgColor            :1; //  5
    unsigned                _fNotInUse              :1; //  6
    unsigned                _fAlwaysUseMyColors     :1; //  7
    unsigned                _fAlwaysUseMyFontSize   :1; //  8
    unsigned                _fAlwaysUseMyFontFace   :1; //  9
    unsigned                _fHasFilters            :1; // 10
    unsigned                _fVisibilityHidden      :1; // 11
    unsigned                _fDisplayNone           :1; // 12
    unsigned                _fRelative              :1; // 13
    unsigned                _uTextJustify           :3; // 14,15,16
    unsigned                _uTextJustifyTrim       :2; // 17,18
    unsigned                _fPadBord               :1; // 19
    unsigned                _fHasBgImage            :1; // 20
    unsigned                _fNoBreak               :1; // 21
    unsigned                _fCtrlAlignLast         :1; // 22
    unsigned                _fPre                   :1; // 23
    unsigned                _fInclEOLWhite          :1; // 24
    unsigned                _fHasExpandos           :1; // 25
    unsigned                _fBidiEmbed             :1; // 26
    unsigned                _fBidiOverride          :1; // 27

    BYTE                    _bBlockAlign;               // Alignment set by DISPID_BLOCKALIGN
    BYTE                    _bControlAlign;             // Alignment set by DISPID_CONTROLALIGN
    BYTE                    _bCtrlBlockAlign;           // For elements with TAGDESC_OWNLINE, they also set the block alignment.
                                                        // Combined with _fCtrlAlignLast we use this to figure
                                                        // out what the correct value for the block alignment is.
    CUnitValue              _cuvTextIndent;
    CUnitValue              _cuvTextKashida;
    CAttrArray*             _pAAExpando;                // AA for style expandos
    CString                 _cstrBgImgUrl;              // URL for background image
    CString                 _cstrLiImgUrl;              // URL for <LI> image
    CString                 _cstrFilters;               // New filters string


    // ^^^^^ All of the above fields are cleared by Reset() ^^^^^
    LONG                    _icfSrc;                    // _icf being inherited
    LONG                    _ipfSrc;                    // _ipf being inherited
    LONG                    _iffSrc;                    // _iff being inherited
    const CCharFormat*      _pcfSrc;                    // Original CCharFormat being inherited
    const CParaFormat*      _ppfSrc;                    // Original CParaFormat being inherited
    const CFancyFormat*     _pffSrc;                    // Original CFancyFormat being inherited
    const CCharFormat*      _pcf;                       // Same as _pcfSrc until _cf is prepared
    const CParaFormat*      _ppf;                       // Same as _ppfSrc until _pf is prepared
    const CFancyFormat*     _pff;                       // Same as _pffSrc until _ff is prepared

    // We can call ComputeFormats to get a style attribute that affects given element
    // _eExtraValues is uesd to request the special mode
    COMPUTEFORMATSTYPE      _eExtraValues;              // If not ComputeFormatsType_Normal next members are used
    DISPID                  _dispIDExtra;               // DISPID of the value requested
    VARIANT*                _pvarExtraValue;            // Returned value. Type depends on _dispIDExtra

    // We can pass in a style object from which to force values.  That is, no inheritance,
    // or cascading.  Just jam the values from this style obj into the formatinfo.  
    CStyle*                 _pStyleForce;               // Style object that's forced in.
private:
    CCharFormat             _cfDst;
    CParaFormat             _pfDst;
    CFancyFormat            _ffDst;
    CAttrArray              _AAExpando;
};

#endif //__XINDOWS_SITE_BASE_FORMATINFO_H__