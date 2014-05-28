
#ifndef __XINDOWS_SITE_STYLE_STYLE_H__
#define __XINDOWS_SITE_STYLE_STYLE_H__

#define _hxx_
#include "../gen/style.hdl"

class CElement;

#define TD_NONE         0x01
#define TD_UNDERLINE    0x02
#define TD_OVERLINE     0x04
#define TD_LINETHROUGH  0x08
#define TD_BLINK        0x10

enum TEXTAUTOSPACETYPE
{
    TEXTAUTOSPACE_NONE         = 0,
    TEXTAUTOSPACE_NUMERIC      = 1,
    TEXTAUTOSPACE_ALPHA        = 2,
    TEXTAUTOSPACE_SPACE        = 4,
    TEXTAUTOSPACE_PARENTHESIS  = 8
};

//+---------------------------------------------------------------------------
//
//  Class:      CStyle (Style Sheets)
//
//   Note:      Even though methods exist for this structure to be cached, we
//              are currently not doing so.
//
//----------------------------------------------------------------------------
class CStyle : public CBase
{
    DECLARE_CLASS_TYPES(CStyle, CBase)

private:
    CElement*   _pElem;
    DISPID      _dispIDAA; // DISPID of Attr Array on _pElem

protected:
    DWORD       _dwFlags;

    union
    {
        WORD _wflags;
        struct
        {
            BOOL _fSeparateFromElem: 1;
        };
    };

public:
    IDispatch* _pStyleSource;

    enum STYLEFLAGS
    {
        STYLE_MASKPROPERTYCHANGES   =   0x1,
        STYLE_SEPARATEFROMELEM      =   0x2,
        STYLE_NOCLEARCACHES         =   0x4, // BUGBUG: (anandra) HACK ALERT
    };

    CStyle(CElement* pElem, DISPID dispID, DWORD dwFlags);
    ~CStyle();

    DECLARE_MEMCLEAR_NEW_DELETE()

    // Support for sub-objects created through pdl's
    // CStyle & Celement implement this differently
    // In the future this may return a NULL ptr!!
    CElement* GetElementPtr(void) { return _pElem; }
    void ClearElementPtr(void) { _pElem = NULL; }

    //Data access
    HRESULT PrivatizeProperties(CVoid** ppv) { return S_OK; }
    void    MaskPropertyChanges(BOOL fMask);

    //Parsing methods
    HRESULT OnPropertyChange(DISPID dispid, DWORD dwFlags);

    DECLARE_TEAROFF_TABLE(IServiceProvider)
    // IServiceProvider methods
    NV_DECLARE_TEAROFF_METHOD(QueryService, queryservice, (REFGUID guidService, REFIID iid, LPVOID* ppv));

    // CBase methods
    DECLARE_TEAROFF_METHOD(PrivateQueryInterface, privatequeryinterface, (REFIID iid, void** ppv));
    HRESULT QueryInterface(REFIID iid, void** ppv) { return PrivateQueryInterface(iid, ppv); }

    NV_DECLARE_TEAROFF_METHOD_(ULONG, PrivateAddRef, privateaddref, ());
    NV_DECLARE_TEAROFF_METHOD_(ULONG, PrivateRelease, privaterelease, ());

    virtual CAtomTable* GetAtomTable(BOOL* pfExpando=NULL);

    BOOL    TestFlag (DWORD dwFlags) const { return 0!=(_dwFlags&dwFlags); };
    void    SetFlag  (DWORD dwFlags) { _dwFlags |= dwFlags; };
    void    ClearFlag(DWORD dwFlags) { _dwFlags &= ~dwFlags; };

    HRESULT getValueHelper(long* plValue, DWORD dwFlags, const PROPERTYDESC* pPropertyDesc);
    HRESULT putValueHelper(long lValue, DWORD dwFlags, DISPID dispid);
    HRESULT getfloatHelper(float* pfValue, DWORD dwFlags, const PROPERTYDESC* pPropertyDesc);
    HRESULT putfloatHelper(float fValue, DWORD dwFlags, DISPID dispid);

    // misc
    BOOL    NeedToDelegateGet(DISPID dispid);
    HRESULT DelegateGet(DISPID dispid, VARTYPE varType, void* pv);

    void    Passivate();
    struct CLASSDESC
    {
        CBase::CLASSDESC _classdescBase;
        void* _apfnTearOff;
    };

    // Helper for HDL file
    virtual CAttrArray** GetAttrArray(void) const;

    NV_DECLARE_TEAROFF_METHOD(put_textDecorationNone, PUT_textDecorationNone, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(get_textDecorationNone, GET_textDecorationNone, (VARIANT_BOOL* p));
    NV_DECLARE_TEAROFF_METHOD(put_textDecorationUnderline, PUT_textDecorationUnderline, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(get_textDecorationUnderline, GET_textDecorationUnderline, (VARIANT_BOOL* p));
    NV_DECLARE_TEAROFF_METHOD(put_textDecorationOverline, PUT_textDecorationOverline, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(get_textDecorationOverline, GET_textDecorationOverline, (VARIANT_BOOL* p));
    NV_DECLARE_TEAROFF_METHOD(put_textDecorationLineThrough, PUT_textDecorationLineThrough, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(get_textDecorationLineThrough, GET_textDecorationLineThrough, (VARIANT_BOOL* p));
    NV_DECLARE_TEAROFF_METHOD(put_textDecorationBlink, PUT_textDecorationBlink, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(get_textDecorationBlink, GET_textDecorationBlink, (VARIANT_BOOL* p));

    // Override all property gets/puts so that the invtablepropdesc points to
    // our local functions to give us a chance to pass the pAA of CStyleRule and
    // not CRuleStyle.
    NV_DECLARE_TEAROFF_METHOD(put_StyleComponent, PUT_StyleComponent, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(put_Url, PUT_Url, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(put_String, PUT_String, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(put_Short, PUT_Short, (short v));
    NV_DECLARE_TEAROFF_METHOD(put_Long, PUT_Long, (long v));
    NV_DECLARE_TEAROFF_METHOD(put_Bool, PUT_Bool, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(put_Variant, PUT_Variant, (VARIANT v));
    NV_DECLARE_TEAROFF_METHOD(get_Url, GET_Url, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD(get_StyleComponent, GET_StyleComponent, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD(get_Property, GET_Property, (void* p));

//#ifndef NO_EDIT
//    virtual IOleUndoManager* UndoManager(void) { return _pElem->UndoManager(); }
//    virtual BOOL QueryCreateUndo(BOOL fRequiresParent, BOOL fDirtyChange=TRUE) 
//    {
//        return _pElem->QueryCreateUndo(fRequiresParent, fDirtyChange);
//    }
//#endif // NO_EDIT

#define _CStyle_
#include "../gen/style.hdl"

protected:
    DECLARE_CLASSDESC_MEMBERS;
};

// Generic Parsing Helpers
TCHAR* NextParenthesizedToken(TCHAR* pszToken);


// Style Property Helpers
HRESULT WriteStyleToBSTR(CBase* pObject, CAttrArray* pAA, BSTR* pbstr, BOOL fQuote);
HRESULT WriteBackgroundStyleToBSTR(CAttrArray* pAA, BSTR* pbstr);
HRESULT WriteFontToBSTR(CAttrArray* pAA, BSTR *pbstr);
HRESULT WriteLayoutGridToBSTR(CAttrArray* pAA, BSTR* pbstr);
HRESULT WriteTextDecorationToBSTR(CAttrArray* pAA, BSTR* pbstr);
HRESULT WriteTextAutospaceToBSTR(CAttrArray* pAA, BSTR* pbstr);
void    WriteTextAutospaceFromLongToBSTR(LONG lTextAutospace, BSTR* pbstr, BOOL fWriteNone);
HRESULT WriteBorderToBSTR(CAttrArray* pAA, BSTR* pbstr);
HRESULT WriteExpandedPropertyToBSTR( DWORD dispid, CAttrArray* pAA, BSTR* pbstr);
HRESULT WriteListStyleToBSTR(CAttrArray* pAA, BSTR* pbstr);
HRESULT WriteBorderSidePropertyToBSTR(DWORD dispid, CAttrArray* pAA, BSTR* pbstr);
HRESULT WriteBackgroundPositionToBSTR(CAttrArray* pAA, BSTR* pbstr);
HRESULT WriteClipToBSTR(CAttrArray* pAA, BSTR* pbstr);


// Style Property Sub-parsers
HRESULT ParseBackgroundProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBGString, BOOL bValidate);
HRESULT ParseFontProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszFontString);
HRESULT ParseLayoutGridProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszGridString);
HRESULT ParseExpandProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBGString, DWORD dwDispId, BOOL fIsMeasurements);
HRESULT ParseBorderProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBorderString);
HRESULT ParseAndExpandBorderSideProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBorderString, DWORD dwDispId);
HRESULT ParseTextDecorationProperty(CAttrArray** ppAA, LPCTSTR pcszTextDecoration, WORD wFlags);
HRESULT ParseTextAutospaceProperty(CAttrArray** ppAA, LPCTSTR pcszTextAutospace, WORD wFlags);
HRESULT ParseListStyleProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszListStyle);
HRESULT ParseBackgroundPositionProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszBackgroundPosition);
HRESULT ParseClipProperty(CAttrArray** ppAA, CBase* pObject, DWORD dwOpCode, LPCTSTR pcszClip);


typedef enum
{
    sysfont_caption = 0,
    sysfont_icon = 1,
    sysfont_menu = 2,
    sysfont_messagebox = 3,
    sysfont_smallcaption = 4,
    sysfont_statusbar = 5,
    sysfont_non_system = -1
} Esysfont;
Esysfont FindSystemFontByName(const TCHAR* szString);

size_t ValidStyleUrl(TCHAR* pch);
HRESULT RemoveStyleUrlFromStr(TCHAR** pch); // Removes "url(" if present from the string


#endif //__XINDOWS_SITE_STYLE_STYLE_H__