
#ifndef __XINDOWS_CORE_BASE_BASE_H__
#define __XINDOWS_CORE_BASE_BASE_H__

class CBase;

// flags for get/Set Attribute/Expression
#define GETMEMBER_CASE_SENSITIVE    0x00000001
#define GETMEMBER_AS_SOURCE         0x00000002

struct PROPERTYDESC;
struct VTABLEDESC
{
    const PROPERTYDESC* pPropDesc;
    UINT                uVTblEntry;
};

// Sorted property desc array.
struct HDLDESC
{
    // Mondo disp interface IID
    const GUID*                 _piidOfMondoDispInterface;

    // PropDesc array
    const PROPERTYDESC* const*  ppPropDescs;
    UINT                        uNumPropDescs;  // Max value binary search.

    //VTable array
    VTABLEDESC const*           pVTableArray;
    UINT                        uVTblIndex;     // Max value binary search.
    // PropDesc array in Vtbl order
    const PROPERTYDESC* const*  ppVtblPropDescs;
    const PROPERTYDESC* FindPropDescForName(LPCTSTR szName, BOOL fCaseSensitive=FALSE, long* pidx=NULL) const;
};

struct ENUMDESC
{
    WORD    cEnums;
    DWORD   dwMask;
    struct ENUMPAIR
    {
        TCHAR*  pszName;
        long    iVal;
    } aenumpairs[];

    HRESULT EnumFromString(LPCTSTR pbstr, long* lValue, BOOL fCaseSensitive=FALSE) const;
    HRESULT StringFromEnum(long lValue, BSTR* pbstr) const;
    LPCTSTR StringPtrFromEnum(long lValue) const;
};
#define ENUMFROMSTRING(enumname, pbstr, pEnumValue) s_enumdesc##enumname.EnumFromString((LPTSTR)pbstr, pEnumValue, FALSE)
#define STRINGFROMENUM(enumname, EnumValue, pbstr)  s_enumdesc##enumname.StringFromEnum(EnumValue, pbstr)
#define STRINGPTRFROMENUM(enumname, EnumValue)      s_enumdesc##enumname.StringPtrFromEnum(EnumValue)

struct BASICPROPPARAMS
{
    DWORD   dwPPFlags;      // PROPPARAM_FLAGS |'d together
    DISPID  dispid;         // automation id
    DWORD   dwFlags;        // flags pass to OnPropertyChange (ELEMCHNG_FLAG)
    WORD    wInvFunc;       // Index into function Invoke table
    WORD    wMaxstrlen;     // maxlength for string and Variant types that can contain strings

    const PROPERTYDESC* GetPropertyDesc(void) const;
    HRESULT GetAvString(void* pObject, const void* pvParams, CString* pstr, BOOL* pfValuePresent=NULL) const;
    DWORD   GetAvNumber(void* pObject, const void* pvParams, UINT uNumBytes, BOOL* pfValuePresent=NULL) const;
    HRESULT SetAvNumber(void*pObject, DWORD dwNumber, const void* pvParams, UINT uNumberBytes, WORD wFlags) const;

    void GetGetMethodP(const void* pvParams, void* pfn) const;
    void GetSetMethodP(const void* pvParams, void* pfn) const;

    HRESULT GetColor(CVoid* pObject, BSTR* pbstrColor, BOOL fReturnAsHex=FALSE) const;
    HRESULT GetColor(CVoid* pObject, CString* pcstr, BOOL fReturnAsHex=FALSE, BOOL* pfValuePresent=NULL) const;
    HRESULT SetColor(CVoid* pObject, TCHAR* pch, WORD wFlags) const;
    HRESULT GetColor(CVoid* pObject, DWORD* pdwValue) const;
    HRESULT SetColor(CVoid* pObject, DWORD dwValue, WORD wFlags) const;
    HRESULT GetColorProperty(VARIANT* pbstr, CBase* pObject, CVoid* pSubObject) const;
    HRESULT SetColorProperty(VARIANT  bstr,  CBase* pObject, CVoid* pSubObject, WORD wFlags=0) const;

    HRESULT GetString(CVoid* pObject, CString* pcstr, BOOL* pfValuePresent=NULL) const;
    HRESULT SetString(CVoid* pObject, TCHAR* pch, WORD wFlags) const;
    HRESULT GetStringProperty(BSTR* pbstr, CBase* pObject, CVoid* pSubObject) const;
    HRESULT SetStringProperty(BSTR  bstr,  CBase* pObject, CVoid* pSubObject, WORD wFlags=0) const;

    HRESULT GetUrl(CVoid* pObject, CString* pcstr) const;
    HRESULT SetUrl(CVoid* pObject, TCHAR* pch, WORD wFlags) const;
    HRESULT GetUrlProperty(BSTR* pbstr, CBase* pObject, CVoid* pSubObject) const;
    HRESULT SetUrlProperty(BSTR  bstr,  CBase* pObject, CVoid* pSubObject, WORD wFlags=0) const;

    HRESULT SetCodeProperty(VARIANT* pvarCode, CBase* pObject, CVoid*) const;
    HRESULT GetCodeProperty(VARIANT* pvarCode, CBase* pObject, CVoid* pSubObject=NULL) const;

    HRESULT GetStyleComponentProperty(BSTR* pbstr, CBase* pObject, CVoid* pSubObject) const;
    HRESULT SetStyleComponentProperty(BSTR bstr,   CBase* pObject, CVoid* pSubObject, WORD wFlags=0) const;
    HRESULT GetStyleComponentBooleanProperty(VARIANT_BOOL* p, CBase* pObject, CVoid* pSubObject) const;
    HRESULT SetStyleComponentBooleanProperty(VARIANT_BOOL  v, CBase* pObject, CVoid* pSubObject) const;
};

struct NUMPROPPARAMS
{
    BASICPROPPARAMS bpp;
    BYTE            vt;
    BYTE            cbMember; // Only used for member access
    LONG            lMin;
    LONG_PTR        lMax;

    const PROPERTYDESC* GetPropertyDesc(void) const;

    HRESULT GetNumber(CBase* pObject, CVoid* pSubObject, long* plNum, BOOL* pfValuePresent=NULL) const;
    HRESULT SetNumber(CBase* pObject, CVoid* pSubObject, long plNum, WORD wFlags) const;

    HRESULT ValidateNumberProperty(long* lArg, CBase* pObject) const;
    HRESULT GetNumberProperty(void* pv, CBase* pObject, CVoid* pSubObject) const;
    HRESULT SetNumberProperty(long lValueNew, CBase* pObject, CVoid* pSubObject, BOOL fValidate=TRUE, WORD wFlags=0) const;
    HRESULT GetEnumStringProperty(BSTR* pbstr, CBase* pObject, CVoid* pSubObject) const;
    HRESULT SetEnumStringProperty(BSTR bstr, CBase* pObject, CVoid* pSubObject, WORD wFlags=0) const;
};

typedef ULONG AAINDEX;
const AAINDEX AA_IDX_UNKNOWN = (AAINDEX)~0UL;

//+-------------------------------------------------------------------
//
// CAttrValue.
//
//--------------------------------------------------------------------
class CAttrValue
{
public:
    //  31              24  23              16  15             8  7           0
    //  -----------------------------------------------------------------------
    //  |                 |                   |                 |             |
    //  |     VT_TYPE     |              AAExtraBits            |    AATYPE   |
    //  |                 |                   |                 |             |
    //  -----------------------------------------------------------------------
    struct AttrFlags
    {
        BYTE _aaType;
        WORD _aaExtraBits;
        BYTE _aaVTType;
    };

    enum AATYPE
    {
        AA_Attribute,
        AA_UnknownAttr, 
        AA_Expando,
        AA_Internal, 
        AA_AttrArray,   
        AA_StyleAttribute, // Used to simplify get/set
        AA_Expression,
        AA_AttachEvent,
        AA_Undefined = ~0UL
    };

    enum AAExtraBits
    {
        AA_Extra_Empty        = 0x0000,   // No bits set in extrabits area

        AA_Extra_Style        = 0x0001,   // The bit used to indicate AA_StyleAttribute
        AA_Extra_Important    = 0x0002,   // Stylesheets: "! important"
        AA_Extra_Implied      = 0x0004,   // Stylesheets: implied property, e.g. backgroundColor: transparent in "background: none"
        AA_Extra_Propdesc     = 0x0008,   // Does the union u1 contain a propdesc * or a dispid?

        AA_Extra_DefaultValue = 0x0020,   // This AV stores a "tag not asssigned" default value
    };

    // other AV types that we use. These are stored in the higher byte of _wFlags
    // in CAttrValue just like the normal VARTYPEs. Their values fall within the range
    // of normal VARIANT types, but are currently not pre-defined.
    enum
    {
        VT_ATTRARRAY = 0x80,
    };

    inline BOOL IsPropdesc() const            { return _wFlags._aaExtraBits&AA_Extra_Propdesc; }
    inline BOOL IsStyleAttribute (void) const { return _wFlags._aaExtraBits&AA_Extra_Style; } 
    inline BOOL IsImportant(void) const       { return _wFlags._aaExtraBits&AA_Extra_Important; } 
    inline BOOL IsImplied(void) const         { return _wFlags._aaExtraBits&AA_Extra_Implied; } 
    inline BOOL IsDefault(void) const         { return _wFlags._aaExtraBits&AA_Extra_DefaultValue; } 

    inline void SetImportant(BOOL fImportant)
    {
        if(fImportant)
        {
            _wFlags._aaExtraBits |= AA_Extra_Important;
        }
        else
        {
            _wFlags._aaExtraBits &= ~AA_Extra_Important;
        }
    }
    inline void SetImplied(BOOL fImplied)
    {
        if(fImplied)
        {
            _wFlags._aaExtraBits |= AA_Extra_Implied;
        }
        else
        {
            _wFlags._aaExtraBits &= ~AA_Extra_Implied;
        }
    }
    inline void SetDefault(BOOL fDefault)
    {
        if(fDefault)
        {
            _wFlags._aaExtraBits |= AA_Extra_DefaultValue;
        }
        else
        {
            _wFlags._aaExtraBits &= ~AA_Extra_DefaultValue;
        }
    }

    // Internal wrapper for AAType - use when you don't care about the style bit
    AATYPE AAType(void) const
    {
        return (AATYPE)(_wFlags._aaType);
    }
    // External wrapper to get AAType - call when need to crack the style bit
    AATYPE GetAAType(void) const
    { 
        AATYPE aa = AAType();
        if(aa==AA_Attribute && IsStyleAttribute())
        {
            aa = AA_StyleAttribute;
        }
        return aa;
    }
    void SetAAType(AATYPE aaType);
    DISPID GetDISPID(void) const;
    void SetDISPID(DISPID dispID)
    {
        u1._dispid = dispID;
        _wFlags._aaExtraBits &= ~AA_Extra_Propdesc;
    }

    const PROPERTYDESC* GetPropDesc() const;
    void SetPropDesc(const PROPERTYDESC* pPropdesc)
    {
        u1._pPropertyDesc = pPropdesc;
        _wFlags._aaExtraBits |= AA_Extra_Propdesc;
    }

    VARTYPE GetAVType() const
    {
        return (VARTYPE)(_wFlags._aaVTType);
    }
    void SetAVType(VARTYPE avType)
    {
        _wFlags._aaVTType = avType;
    }

    // NOTE (SRamani/TerryLu): If you need to use more AAtypes, or extra bits,
    // I've increased _wFlags to a DWORD (the extra WORD is used on win32 due
    // to padding) if you need more then the 16 extra bits do the following: 
    //
    //      1) Compress the AVType which is the higher byte of _wFlags into the
    //         top nibble and use some sort of a look up table to map the
    //         VARTYPEs to the AVTYPEs, thus gaining 4 more bits. I chose not
    //         to do this now due to perf reasons.
    AttrFlags _wFlags;

    union
    {
        const PROPERTYDESC* _pPropertyDesc;
        struct
        {
            DISPID _dispid;
#ifdef _WIN64
            DWORD _dwCookie; // Used for Advise/Unadvise on Win64 platforms
#endif
        };
    } u1;
    union
    {
        IUnknown*   _pUnk;
        IDispatch*  _pDisp;
        LONG        _lVal;
        BOOL        _boolVal;
        SHORT       _iVal;
        FLOAT       _fltVal;
        LPTSTR      _lpstrVal;
        BSTR        _bstrVal;
        VARIANT*    _pvarVal;
        double*     _pdblVal;
        CAttrArray* _pAA;
        void*       _pvVal;
    } u2;

    CAttrValue() {}

    int     CompareWith(DISPID dispID, AATYPE aaType);

    HRESULT Copy(const CAttrValue* pAV);
    void    CopyEmpty(const CAttrValue* pAV);
    BOOL    Compare(const CAttrValue* pAV) const;
    void    Free();
    HRESULT InitVariant(const VARIANT* pvarNew, BOOL fCloning=FALSE);
    HRESULT GetIntoVariant(VARIANT* pvar) const;
    HRESULT GetIntoString(BSTR* pbStr, LPCTSTR* ppStr) const;
    // copies the attr value into a variant, w/o copying strings or addref'ing objects
    void    GetAsVariantNC(VARIANT* pvar) const;

    // Sets the value of the AV to the given variant by copying it into a newly allocated
    // VARIANT referenced as a ptr
    HRESULT SetVariant(const VARIANT* pvar);
    inline VARIANT* GetVariant() const  { Assert(GetAVType() == VT_VARIANT); return u2._pvarVal; }
    HRESULT SetDouble(double dblVal);
    inline double GetDouble() const     { Assert(u2._pdblVal); Assert(GetAVType() == VT_R8); return *(u2._pdblVal); }
    inline double* GetDoublePtr() const { Assert(u2._pdblVal); Assert(GetAVType() == VT_R8); return u2._pdblVal; }

    // inline put helpers that set the value and type of the AV
    // use these when changing the value as well as the type of an AV
    inline void SetLong(LONG lVal, VARTYPE avType)      { u2._lVal = lVal; SetAVType(avType); }
    inline void SetShort(SHORT iVal, VARTYPE avType)    { u2._iVal = iVal; SetAVType(avType); }
    inline void SetBSTR(BSTR bstrVal)                   { u2._bstrVal = bstrVal; SetAVType(VT_BSTR); }
    inline void SetLPWSTR(LPWSTR lpstrVal)              { u2._lpstrVal = lpstrVal; SetAVType(VT_LPWSTR); }
    inline void SetPointer(void* pvVal, VARTYPE avType) { u2._pvVal = pvVal; SetAVType(avType); }
    inline void SetAA(CAttrArray* pAA)                  { u2._pAA = pAA; SetAVType(VT_ATTRARRAY); }

#ifdef _WIN64
    inline void SetCookie(DWORD dwCookie) { u1._dwCookie = dwCookie; }
#endif

    // inline put helpers that Assert the type of the AV and then set its value
    // use these when changing the value of an AV of the same type
    inline void PutLong(LONG lVal)          { Assert(GetAVType() == VT_I4); u2._lVal = lVal; }

    // inline get helpers that Assert the type of the AV and then return its value. 
    // use these when getting the value of an AV whose type is known
    inline LONG GetLong() const             { Assert(GetAVType() == VT_I4); return u2._lVal; }
    inline SHORT GetShort() const           { Assert(GetAVType() == VT_I2); return u2._iVal; }
    inline FLOAT GetFloat() const           { Assert(GetAVType() == VT_R4); return u2._fltVal; }
    inline BSTR GetBSTR() const             { Assert(GetAVType() == VT_BSTR); return u2._bstrVal; }
    inline LPTSTR GetLPWSTR() const         { Assert(GetAVType() == VT_LPWSTR); return u2._lpstrVal; }
    inline BOOL GetBool() const             { Assert(GetAVType() == VT_BOOL); return u2._boolVal; }
    inline void* GetPointer() const         { Assert(GetAVType() == VT_PTR); return u2._pvVal; }
    inline IUnknown* GetUnknown() const     { Assert(GetAVType() == VT_UNKNOWN); return u2._pUnk; }
    inline IDispatch* GetDispatch() const   { Assert(GetAVType() == VT_DISPATCH); return u2._pDisp; }
    inline CAttrArray* GetAA() const        { Assert(GetAVType() == VT_ATTRARRAY); return u2._pAA; }

#ifdef _WIN64
    inline DWORD GetCookie()                { return u1._dwCookie; }
#endif

    // inline get helpers that get the value of the common VARIANT types, w/o asserting
    // use these when getting the value of an AV whose type is not known
    inline LONG GetlVal() const             { return u2._lVal; }
    inline SHORT GetiVal() const            { return u2._iVal; }
    inline IUnknown* GetpUnkVal() const     { return u2._pUnk; }
    inline LPTSTR GetString() const         { return u2._lpstrVal; }
    inline void* GetPointerVal() const      { return u2._pvVal; }
    inline CAttrArray** GetppAA() const     { return (CAttrArray**)(&(u2._pAA)); }
};

//+---------------------------------------------------------------------------
//
// CAttrArray.
//
//----------------------------------------------------------------------------
class CAttrArray : public CDataAry<CAttrValue>
{
    friend class CBase;

private:
    // Maintain a simple checksum so we can quickly compare to attrarrays
    // Checksum is computed by adding DISPID's of attr array members
    // NOTE: The low bit of the checksum indicates the presence of 
    // style ptr caches in the array.  This speeds up shutdown.
    DWORD _dwChecksum;

    HRESULT Set(DISPID dispID,const PROPERTYDESC* pPropDesc, VARIANT* varNew,
        CAttrValue::AATYPE aaType=CAttrValue::AA_Attribute,
        WORD wExtraBits=0, // Modifiers like "important" or "implied"
        BOOL fAllowMultiple=FALSE);

protected:
    void Destroy(int iIndex);

public:
    DECLARE_MEMALLOC_NEW_DELETE();

    CAttrArray();
    ~CAttrArray() { Free(); }

    void Free(void);
    void FreeSpecial();
    void Clear(CAttrArray* pAAUndo);

    BOOL Compare(CAttrArray* pAA, DISPID* pdispIDDifferent=NULL) const;

    BOOL HasAnyAttribute(BOOL fCountExpandosToo=FALSE);

    static HRESULT Set(CAttrArray** ppAA, DISPID dispID, VARIANT* varNew,
        const PROPERTYDESC* pPropDesc=NULL,
        CAttrValue::AATYPE aaType=CAttrValue::AA_Attribute,
        WORD wExtraBits=0, // Modifiers like "important" or "implied"
        BOOL fAllowMultiple=FALSE);

    HRESULT SetAt(AAINDEX aaIdx, VARIANT* pvarNew);

    CAttrValue* FindAt(AAINDEX aaIdx)
    {
        return (((aaIdx>=0)&&(aaIdx<(ULONG)Size())) ? ((CAttrValue*)*this)+aaIdx : NULL);
    }
    CAttrValue* Find(DISPID dispID, CAttrValue::AATYPE aaType=CAttrValue::AA_Attribute, AAINDEX* paaIdx=NULL, BOOL fAllowMultiple=FALSE);
    static BOOL FindSimple(CAttrArray* pAA, const PROPERTYDESC* pPropertyDesc, DWORD* pdwValue);
    static BOOL FindString(CAttrArray* pAA, const PROPERTYDESC* pPropertyDesc, LPCTSTR* ppStr);

    BOOL FindString(DISPID dispID, LPCTSTR* ppStr, CAttrValue::AATYPE aaType=CAttrValue::AA_Attribute, const PROPERTYDESC**ppPropertyDesc=NULL);
    BOOL FindSimple(DISPID dispID, DWORD* pdwValue, CAttrValue::AATYPE aaType=CAttrValue::AA_Attribute, const PROPERTYDESC**ppPropertyDesc=NULL);
    BOOL FindSimpleInt4AndDelete(DISPID dispID, DWORD* pdwValue, CAttrValue::AATYPE aaType=CAttrValue::AA_Attribute, const PROPERTYDESC**ppPropertyDesc=NULL);
    BOOL FindSimpleAndDelete(DISPID dispID, CAttrValue::AATYPE aaType=CAttrValue::AA_Attribute, const PROPERTYDESC** ppPropertyDesc=NULL);

    static HRESULT SetSimple(CAttrArray** ppAA, const PROPERTYDESC* pPropertyDesc, DWORD dwSimple, WORD wFlags=0);
    static HRESULT SetString(CAttrArray** ppAA, const PROPERTYDESC* pPropertyDesc, LPCTSTR pch, BOOL fIsUnknown=FALSE, WORD wFlags=0);

    AAINDEX FindAAIndex(DISPID dispID, CAttrValue::AATYPE aaType, AAINDEX aaLastOne=AA_IDX_UNKNOWN, BOOL fAllowMultiple=FALSE);

    HRESULT GetSimpleAt(AAINDEX aaIdx, DWORD* pdwValue);
    static HRESULT AddSimple(CAttrArray** ppAA, DISPID dispID, DWORD dwSimple, CAttrValue::AATYPE aaType)
    {
        VARIANT varNew;

        varNew.vt = VT_I4;
        varNew.lVal = (long)dwSimple;

        RRETURN(CAttrArray::Set(ppAA, dispID, &varNew, NULL, aaType));
    }

    HRESULT Clone(CAttrArray** ppAA) const;
    HRESULT Merge(CAttrArray** ppAA, CBase* pTarget, CAttrArray* pAAUndo, BOOL fFromUndo=FALSE, BOOL fCopyID=FALSE) const;

    // Copy expandos from given attribute array to current one
    HRESULT CopyExpandos(CAttrArray* pAA);

    DWORD GetChecksum() const { return _dwChecksum; }
    // This method is only for the data cache and casts the checksum down to a word
    WORD ComputeCrc() const { return (WORD)GetChecksum(); }

    BOOL IsStylePtrCachePossible()   { return _dwChecksum&1; }
    void SetStylePtrCachePossible()  { _dwChecksum |= 1; }
};

//+---------------------------------------------------------------------------
//
//  Class:      CBase
//
//  Purpose:    A generally useful base class.
//
//----------------------------------------------------------------------------
class NOVTABLE CBase : public CVoid, public IPrivateUnknown
{
    DECLARE_CLASS_TYPES(CBase, CVoid)

private:
    DECLARE_MEMALLOC_NEW_DELETE();

public:
    CBase();
    virtual ~CBase();
    virtual HRESULT Init();
    virtual void Passivate();

    DECLARE_TEAROFF_METHOD_(ULONG, PrivateAddRef, privateaddref, ());
    DECLARE_TEAROFF_METHOD_(ULONG, PrivateRelease, privaterelease, ());
    STDMETHOD(PrivateQueryInterface)(REFIID, void**);

    // IObjectIdentity members
    NV_DECLARE_TEAROFF_METHOD(IsEqualObject, isequalobject, (IUnknown* ppunk));

    STDMETHOD(GetEnabled)(VARIANT_BOOL* pfEnabled);
    STDMETHOD(GetValid)(VARIANT_BOOL* pfEnabled);

    // IOleCommandTarget methods
    STDMETHOD(QueryStatus)(
        GUID*       pguidCmdGroup,
        ULONG       cCmds,
        OLECMD      rgCmds[],
        OLECMDTEXT* pcmdtext)
    {
        return MSOCMDERR_E_NOTSUPPORTED;
    }
    STDMETHOD(Exec)(
        GUID*       pguidCmdGroup,
        DWORD       nCmdID,
        DWORD       nCmdexecopt,
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut)
    {
        return MSOCMDERR_E_NOTSUPPORTED;
    }

    static int  IDMFromCmdID(const GUID* pguidCmdGroup, ULONG ulCmdID);
    static BOOL IsCmdGroupSupported(const GUID* pguidCmdGroup);

    // Fire an event and get the returned value
    HRESULT FireCancelableEvent(DISPID dispidMethod, DISPID dispidProp,
        IDispatch* pEventObject, VARIANT_BOOL* pfRetVal, BYTE* pbTypes, ...);

    HRESULT queryCommandSupported(const BSTR cmdID,VARIANT_BOOL* pfRet);
    HRESULT queryCommandEnabled(const BSTR cmdID,VARIANT_BOOL* pfRet);
    HRESULT queryCommandState(const BSTR cmdID,VARIANT_BOOL* pfRet);
    HRESULT queryCommandIndeterm(const BSTR cmdID,VARIANT_BOOL* pfRet);
    HRESULT queryCommandText(const BSTR cmdID,BSTR* pcmdText);
    HRESULT queryCommandValue(const BSTR cmdID,VARIANT* pcmdValue);
    HRESULT execCommand(const BSTR cmdID,VARIANT_BOOL showUI,VARIANT value);
    HRESULT execCommandShowHelp(const BSTR cmdID);

    // Helpers for non-abstract property get\put
    // NOTE: Any addtion of put/get must have a corresponding helper.
    NV_DECLARE_TEAROFF_METHOD(put_StyleComponent, PUT_StyleComponent, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(put_Url, PUT_Url, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(put_String, PUT_String, (BSTR v));
    NV_DECLARE_TEAROFF_METHOD(put_Short, PUT_Short, (short v));
    NV_DECLARE_TEAROFF_METHOD(put_Long, PUT_Long, (long v));
    NV_DECLARE_TEAROFF_METHOD(put_Bool, PUT_Bool, (VARIANT_BOOL v));
    NV_DECLARE_TEAROFF_METHOD(put_Variant, PUT_Variant, (VARIANT v));
    NV_DECLARE_TEAROFF_METHOD(put_DataEvent, PUT_DataEvent, (VARIANT v));
    NV_DECLARE_TEAROFF_METHOD(get_Url, GET_Url, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD(get_StyleComponent, GET_StyleComponent, (BSTR* p));
    NV_DECLARE_TEAROFF_METHOD(get_Property, GET_Property, (void* p));

    NV_DECLARE_TEAROFF_METHOD(put_StyleComponentHelper, PUT_StyleComponentHelper, (BSTR v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(put_UrlHelper, PUT_UrlHelper, (BSTR v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(put_StringHelper, PUT_StringHelper, (BSTR v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(put_ShortHelper, PUT_ShortHelper, (short v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(put_LongHelper, PUT_LongHelper, (long v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(put_BoolHelper, PUT_BoolHelper, (VARIANT_BOOL v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(put_VariantHelper, PUT_VariantHelper, (VARIANT v, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(get_UrlHelper, GET_UrlHelper, (BSTR *p, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(get_StyleComponentHelper, GET_StyleComponentHelper, (BSTR* p, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));
    NV_DECLARE_TEAROFF_METHOD(get_PropertyHelper, GET_PropertyHelper, (void* p, const PROPERTYDESC* pPropDesc, CAttrArray** ppAttr=NULL));

protected:
    HRESULT ExecSetGetProperty(
        VARIANTARG* pvarargIn,
        VARIANTARG* pvarargOut,
        UINT        uPropName,
        VARTYPE     vt);
    HRESULT ExecSetGetKnownProp(
        VARIANTARG* pvarargIn, 
        VARIANTARG* pvarargOut,
        DISPID      dispidProp, 
        VARTYPE     vt);
    HRESULT ExecSetGetHelper(
        VARIANTARG* pvarargIn, 
        VARIANTARG* pvarargOut,
        IDispatch*  pDispatch, 
        DISPID      dispid,    
        VARTYPE     vt);

    HRESULT ExecSetPropertyCmd(UINT uPropName, DWORD value);
    HRESULT ExecToggleCmd(UINT uPropName);

    HRESULT QueryStatusProperty(OLECMD* pCmd, UINT uPropName, VARTYPE vt);

    HRESULT GetDispatchForProp(UINT uPropName, IDispatch** ppDisp, DISPID* pdispid);

    // Translates command name into command ID
    static HRESULT CmdIDFromCmdName(BSTR cmdName, ULONG* pcmdValue);

    // Returns the expected VARIANT type of the command value (like VT_BSTR for font name)
    static VARTYPE GetExpectedCmdValueType(ULONG ulCmdID);

    // QueryCommandXXX helper function
    HRESULT QueryCommandHelper(BSTR cmdID, DWORD* cmdf, BSTR* pcmdTxt);

    struct CMDINFOSTRUCT
    {
        TCHAR* cmdName;
        ULONG cmdID;
    };

protected:
    HRESULT toString(BSTR* bstrString);

public:
    ULONG GetRefs() { return _ulAllRefs; }
    ULONG GetObjectRefs() { return _ulRefs; }

    ULONG SubAddRef() { return ++_ulAllRefs; }
    ULONG SubRelease();

    IUnknown* PunkInner() { return (IUnknown*)(IPrivateUnknown*)this; }
    virtual IUnknown* PunkOuter() { return PunkInner(); }

    virtual HRESULT STDMETHODCALLTYPE QueryService(REFGUID, REFIID, void** ppv);

    // Connection Point Helpers.
    HRESULT DoAdvise(
        REFIID      riid,
        DISPID      dispidBase,
        IUnknown*   pUnkSinkAdvise,
        IUnknown**  ppUnkSinkOut,
        DWORD*      pdwCookie);
    HRESULT DoUnadvise(DWORD dwCookie, DISPID dispidBase);
    HRESULT FindAdviseIndex(DISPID dispidBase, DWORD dwCookie, IUnknown* pUnkCookie, AAINDEX* paaidx);

    HRESULT FindEventName(ITypeInfo* pTISrc, DISPID dispid, BSTR* pBSTR);

    HRESULT FireEventV(
        DISPID      dispidEvent, 
        DISPID      dispidProp, 
        IDispatch*  pEventObject,
        VARIANT*    pv, 
        BYTE*       pbTypes, 
        va_list     valParms);
    HRESULT FireEvent(
        DISPID      dispidEvent, 
        DISPID      dispidProp, 
        VARIANT*    pvarResult, 
        DISPPARAMS* pdispparams,
        EXCEPINFO*  pexcepinfo,
        UINT*       puArgErr,
        ITypeInfo*  pTIEventSrc=NULL,
        IDispatch*  pEventObject=NULL);

    HRESULT FirePropertyNotify(DISPID dispid, BOOL fOnChanged);
    HRESULT InvokeDispatchWithThis(
        IDispatch*  pDisp,
        VARIANT*    pExtraArg,
        REFIID      riid,
        LCID        lcid,
        WORD        wFlags,
        DISPPARAMS* pdispparams,
        VARIANT*    pvarResult,
        EXCEPINFO*  pexcepinfo,
        IServiceProvider* pSrvProvider);
    HRESULT InvokeAA(
        DISPID              dispidMember,
        CAttrValue::AATYPE  aaType,
        LCID                lcid,
        WORD                wFlags,
        DISPPARAMS*         pdispparams,
        VARIANT *           pvarResult,
        EXCEPINFO*          pexcepinfo,
        IServiceProvider*   pSrvProvider);

    HRESULT FireOnChanged(DISPID dispid) { return FirePropertyNotify(dispid, TRUE); }
    HRESULT FireRequestEdit(DISPID dispid) { return FirePropertyNotify(dispid, FALSE); }
    virtual HRESULT OnPropertyChange(DISPID dispid, DWORD dwFlags) { return S_OK; }

    // Error Helpers
    virtual void PreSetErrorInfo() {}
    HRESULT SetErrorInfo(HRESULT hr);
    HRESULT SetErrorInfoPSet(HRESULT hr, DISPID dispid);
    HRESULT SetErrorInfoPGet(HRESULT hr, DISPID dispid);
    HRESULT SetErrorInfoInvalidArg();
    HRESULT SetErrorInfoBadValue(DISPID dispid, UINT ids, ...);
    HRESULT SetErrorInfoPBadValue(DISPID dispid, UINT ids, ...);
    HRESULT CloseErrorInfo(HRESULT hr, DISPID id, INVOKEKIND invkind);
    HRESULT CloseErrorInfoPSet(HRESULT hr, DISPID dispid)
    { return CloseErrorInfo(hr, dispid, INVOKE_PROPERTYPUT); }
    HRESULT CloseErrorInfoPGet(HRESULT hr, DISPID dispid)
    { return CloseErrorInfo(hr, dispid, INVOKE_PROPERTYGET); }
    virtual HRESULT CloseErrorInfo(HRESULT hr);

    // Property helpers
    HRESULT SetCodeProperty(DISPID dispidCodeProp, IDispatch* pDispCode, BOOL* pfAnyDeleted=NULL);
    
    // Helper for removeAttribute (also called by recalc engine)
    BOOL removeAttributeDispid(DISPID dispid, const PROPERTYDESC* pPropDesc=NULL);     

    // Helper functions for IDispatchEx:
    HRESULT AddExpando(LPOLESTR pszExpandoName, DISPID* pdispID);
    HRESULT GetExpandoName(DISPID expandoDISPID, LPCTSTR* lpPropName);
    HRESULT SetExpando(LPOLESTR pszExpandoName, VARIANT* pvarPropertyValue);

    HRESULT NextProperty(DISPID currDispID, BSTR* pStrNextName, DISPID* pNextDispID);

    BOOL    IsExpandoDISPID(DISPID dispid, DISPID* pOLESiteExpando=NULL);

    inline const PROPERTYDESC* FindPropDescForName(LPCTSTR szName, BOOL fCaseSensitive=FALSE, long* pidx=NULL)
    {
        if(BaseDesc() && BaseDesc()->_apHdlDesc)
        {
            return BaseDesc()->_apHdlDesc->FindPropDescForName(szName, fCaseSensitive, pidx);
        }
        return NULL;
    }
    
    const VTABLEDESC* FindVTableEntryForName(LPCTSTR szName, BOOL fCaseSensitive, WORD* pVTblOffset=NULL);

    HRESULT ExpandedRelativeUrlInVariant(VARIANT* pVariantURL);

    // Helper functions for custom Invoke:
    HRESULT FindPropDescFromDispID(DISPID dispidMember, PROPERTYDESC** ppDesc, WORD* pwEntry, WORD* pwClassType);

    //+-----------------------------------------------------------------------
    //
    //  CBase::CLock
    //
    //------------------------------------------------------------------------
    class CLock
    {
    public:
        CLock(CBase* pBase);
        ~CLock();
    private:
        CBase* _pBase;
    };

    // 类描述
    struct CLASSDESC
    {
        const CLSID*    _pclsid;            // class's unique identifier
        const CONNECTION_POINT_INFO* _pcpi; // connection pt info, NULL terminated
        DWORD           _dwFlags;           // any combination of the *_FLAG
        const IID*      _piidDispinterface; // class's dispinterface IID
        HDLDESC*        _apHdlDesc;         // Description arrays (NULL terminated) in .HDL files
    };

    // 最终类通过DECLARE_CLASSDESC_MEMBERS声明类描述
    // 对应多种控件的类需承载GetClassDesc函数进行对不同类型控件单独处理
    // 比如：Input Button
    virtual const CBase::CLASSDESC* GetClassDesc() const = 0;
    const CLASSDESC* BaseDesc() const { return GetClassDesc(); }

    // 获取属性列表
    const PROPERTYDESC* const* GetPropDescArray(void) 
    { 
        if(BaseDesc() && BaseDesc()->_apHdlDesc)
        {
            return BaseDesc()->_apHdlDesc->ppPropDescs;
        }
        else
        {
            return NULL;
        }
    }
    // 获取属性数量
    UINT GetPropDescCount(void)
    { 
        if(BaseDesc() && BaseDesc()->_apHdlDesc)
        {
            return BaseDesc()->_apHdlDesc->uNumPropDescs;
        }
        else
        {
            return 0;
        }
    }

    const VTABLEDESC* GetVTableArray(void)
    { 
        if(BaseDesc() && BaseDesc()->_apHdlDesc)
        {
            return BaseDesc()->_apHdlDesc->pVTableArray;
        }
        else
        {
            return NULL;
        }
    }
    UINT GetVTableCount(void)
    { 
        if(BaseDesc() && BaseDesc()->_apHdlDesc)
        {
            return BaseDesc()->_apHdlDesc->uVTblIndex;
        }
        else
        {
            return 0;
        }
    }

    virtual CAtomTable* GetAtomTable(BOOL* pfExpando=NULL)
    {
        if(pfExpando)
        {
            *pfExpando = TRUE;
        }
        return NULL;
    }

    CAttrArray** GetAttrArray() const { return const_cast<CAttrArray**>(&_pAA); }
    void SetAttrArray(CAttrArray* pAA) { _pAA = pAA; }

    AAINDEX FindAAIndex(DISPID dispID, CAttrValue::AATYPE aaType, AAINDEX aaLastOne=AA_IDX_UNKNOWN, BOOL fAllowMultiple=FALSE)
    {
        return (_pAA && _pAA->Find(dispID, aaType, &aaLastOne, fAllowMultiple))?aaLastOne:AA_IDX_UNKNOWN;
    }

    AAINDEX FindAAIndex(DISPID dispID, CAttrValue::AATYPE aaType, AAINDEX aaLastOne,
        BOOL fAllowMultiple, CAttrArray* pAA)
    {
        return (pAA&&pAA->Find(dispID, aaType, &aaLastOne, fAllowMultiple))?aaLastOne:AA_IDX_UNKNOWN;
    }

    BOOL DidFindAAIndexAndDelete(DISPID dispID, CAttrValue::AATYPE aaType);

    void FindAAIndexAndDelete(DISPID dispID, CAttrValue::AATYPE aaType)
    {
        DidFindAAIndexAndDelete(dispID, aaType);
    }

    AAINDEX FindAAType(CAttrValue::AATYPE aaType, AAINDEX lastIndex);
    AAINDEX FindNextAAIndex(DISPID dispid, CAttrValue::AATYPE aatype, AAINDEX aaLastOne);

    BOOL FindString(DISPID dispID, LPCTSTR* ppStr, CAttrValue::AATYPE aaType=CAttrValue::AA_Attribute,
        const PROPERTYDESC** ppPropertyDesc=NULL)
    {
        if(_pAA)
        {
            return _pAA->FindString(dispID, ppStr, aaType, ppPropertyDesc);
        }
        else
        {
            return FALSE;
        }
    }

    HRESULT GetStringAt(AAINDEX aaIdx, LPCTSTR* ppStr);
    HRESULT GetIntoStringAt(AAINDEX aaIdx, BSTR* pbStr, LPCTSTR* ppStr);
    HRESULT GetIntoBSTRAt(AAINDEX aaIdx, BSTR* pbStr);

    HRESULT GetSimpleAt(AAINDEX aaIdx, DWORD* pdwValue)
    {
        if(_pAA)
        {
            return _pAA->GetSimpleAt(aaIdx, pdwValue);
        }
        else
        {
            return  DISP_E_MEMBERNOTFOUND;
        }
    }
    HRESULT GetPointerAt(AAINDEX aaIdx, void** pdwValue);
    HRESULT GetVariantAt(AAINDEX aaIdx, VARIANT* pVar, BOOL fAllowEmptyVariant=TRUE);
    HRESULT GetUnknownObjectAt(AAINDEX aaIdx, IUnknown** ppUnk)
    {
        RRETURN(GetObjectAt(aaIdx, VT_UNKNOWN, (void**)ppUnk));
    }
    HRESULT GetDispatchObjectAt(AAINDEX aaIdx, IDispatch** ppDisp)
    {
        RRETURN(GetObjectAt(aaIdx, VT_DISPATCH, (void**)ppDisp));
    }

    DISPID GetDispIDAt(AAINDEX aaIdx)
    { 
        if(_pAA)
        {
            CAttrValue* pAV;

            pAV = _pAA->FindAt(aaIdx);
            if(pAV)
            {
                return pAV->GetDISPID();
            }
        }
        return DISPID_UNKNOWN;
    }
    CAttrValue::AATYPE GetAAtypeAt(AAINDEX aaIdx);
    CAttrValue* GetAttrValueAt(AAINDEX aaIdx) { return _pAA?_pAA->FindAt(aaIdx):NULL; }
    UINT    GetVariantTypeAt(AAINDEX aaIdx);
    void    DeleteAt(AAINDEX aaIdx) { if(_pAA) _pAA->Destroy(aaIdx); }
    HRESULT AddSimple(DISPID dispID, DWORD dwSimple, CAttrValue::AATYPE aaType);
    HRESULT AddPointer(DISPID dispID, void* pValue, CAttrValue::AATYPE aaType);
    HRESULT AddString(DISPID dispID, LPCTSTR pch, CAttrValue::AATYPE aaType);
    HRESULT AddBSTR(DISPID dispID, LPCTSTR pch, CAttrValue::AATYPE aaType);
    HRESULT AddUnknownObject(DISPID dispID, IUnknown* pUnk, CAttrValue::AATYPE aaType);
    HRESULT AddDispatchObject(DISPID dispID, IDispatch* pDisp, CAttrValue::AATYPE aaType);
    HRESULT AddVariant(DISPID dispID, VARIANT* pVar, CAttrValue::AATYPE aaType);
    HRESULT AddAttrArray(DISPID dispID, CAttrArray* pAttrArray, CAttrValue::AATYPE aaType);
    HRESULT AddUnknownObjectMultiple(DISPID dispid, IUnknown* pUnk, CAttrValue::AATYPE aaType,
        CAttrValue::AAExtraBits wFlags=CAttrValue::AA_Extra_Empty);
    HRESULT AddDispatchObjectMultiple(DISPID dispid, IDispatch* pDisp, CAttrValue::AATYPE aaType);

    HRESULT ChangeSimpleAt(AAINDEX aaIdx, DWORD dwSimple);
    HRESULT ChangeStringAt(AAINDEX aaIdx, LPCTSTR pch);
    HRESULT ChangeUnknownObjectAt(AAINDEX aaIdx, IUnknown* pUnk);
    HRESULT ChangeDispatchObjectAt(AAINDEX aaIdx, IDispatch* pDisp);
    HRESULT ChangeVariantAt(AAINDEX aaIdx, VARIANT* pVar);
    HRESULT ChangeAATypeAt(AAINDEX aaIdx, CAttrValue::AATYPE aaType);

    HRESULT FetchObject(CAttrValue* pAV, VARTYPE vt, void** ppvoid);
    HRESULT GetObjectAt(AAINDEX aaIdx, VARTYPE vt, void** ppVoid);

#ifdef _WIN64
    HRESULT GetCookieAt(AAINDEX aaIdx, DWORD* pdwCookie);
    HRESULT SetCookieAt(AAINDEX aaIdx, DWORD dwCookie);
#endif

protected:
    ULONG _ulRefs;
    ULONG _ulAllRefs;
    CAttrArray* _pAA;

#ifdef _DEBUG
public:
    // These "social security" numbers help us identify and track object
    // allocations and manipulations in the debugger.  At age 65 all objects retire.
    ULONG _ulSSN;
    static ULONG  s_ulLastSSN;
#endif

protected:

#define _CBase_
#include "../gen/types.hdl"

    // Tear off interfaces
    DECLARE_TEAROFF_TABLE(IOleCommandTarget)
    DECLARE_TEAROFF_TABLE(IObjectIdentity)

    static CMDINFOSTRUCT cmdTable[];
};

#define DECLARE_CLASSDESC_MEMBERS \
    static const CLASSDESC s_classdesc; \
    virtual const CBase::CLASSDESC* GetClassDesc() const \
    { return (CBase::CLASSDESC*)&s_classdesc; }

// The following flags specify which data follows relevant structures.  They
// follow in the order specified, member, indirect, etc...
enum PROPPARAM_FLAGS
{
    PROPPARAM_MEMBER                = 0x00000001,
    PROPPARAM_GETMFHandler          = 0x00000004,
    PROPPARAM_SETMFHandler          = 0x00000008,
    PROPPARAM_RESTRICTED            = 0x00000020,
    PROPPARAM_HIDDEN                = 0x00000040,
    // PROPTYPE_NOPERSIST is used to mark entries in the descriptor table which are
    // no longer supported by the object.  They are not written out, and are
    // only parsed (without actually reading the data) when read in.  This
    // flag should be OR'd together with the original type, so that the data
    // size can be inferred.
    PROPPARAM_NOPERSIST             = 0x00000080,

    // Property supports Get and/or Put (Read-only, R/W, or Write-only).  This
    // also signals if the the item is a property or method (either set it's a
    // property otherwise a method).
    PROPPARAM_INVOKEGet             = 0x00000100,
    PROPPARAM_INVOKESet             = 0x00000200,

    // The lMax member of NUMPROPPARAMS structure points to the ENUMDESC
    // structure if this flag is set
    PROPPARAM_ENUM                  = 0x00000400,
    // Value is a CUnitValue-encoded long
    PROPPARAM_UNITVALUE             = 0x00000800,

    // Validation masks for PROPPARAM_UNITVALUE types
    PROPPARAM_BITMASK               = 0x00001000,

    // Specifies that attribute can represent a relative fontsize 1..7 or -1..+7
    PROPPARAM_FONTSIZE              = 0x00002000,

    // Specifies that attribute can represent any legal unit measurement type e.g. 1em, 2px etc.
    PROPPARAM_LENGTH                = 0x00004000,

    // Specifies that attribute can represent a percentage value  e.g. 25%
    PROPPARAM_PERCENTAGE            = 0x00008000,

    // Specifies that attribute can indicate table relative i.e. *
    PROPPARAM_TIMESRELATIVE         = 0x00010000,

    // Specifies that attribute can be a number without a unit specified e.g. 1, +1, 23 etc
    PROPPARAM_ANUMBER               = 0x00020000,

    // specifies whether the ulTagNotAssignedDefault value should be used.  If not, its type will still be used for CUnitValues.
    PROPPARAM_DONTUSENOTASSIGNED    = 0x00040000,

    // specifies whether illegal values should be set to the MIN or to the DEFAULT
    PROPPARAM_MINOUT                = 0x00080000,

    // Specifies that attribute is located in the AttrArray
    PROPPARAM_ATTRARRAY             = 0x00100000,

    // Specifies that attribute is case sensitive (eg. TYPE_TYPE)
    PROPPARAM_CASESENSITIVE         = 0x00200000,

    // This marks a property as containing a scriptlet.  Uses DISPID_THIS for
    // when invoked.
    PROPPARAM_SCRIPTLET             = 0x01000000,

    // This marks a property as applying to a CF/PF/FF/SF
    PROPPARAM_STYLISTIC_PROPERTY    = 0x02000000,

    // This marks a property as being a stylesheet property (used for quoting, etc.)
    PROPPARAM_STYLESHEET_PROPERTY   = 0x04000000,

    // If this flag is set, when we see "TAG PROPERTY=" OR "PROPERTY"
    // we apply the default ( rather than the noassigndefault ) AND if
    // The property is invalid we apply the noassigndefault (rather than the default)
    PROPPARAM_INVALIDASNOASSIGN     = 0x08000000,
    PROPPARAM_NOTPRESENTASDEFAULT   = 0x10000000,
};

// Handy unit value combos
//
// Plain font
#define PP_UV_FONT                  (PROPPARAM_FONTSIZE|PROPPARAM_UNITVALUE)

// CSS font - does not allow unadorned numbers
#define PP_UV_UNITFONT              (PROPPARAM_UNITVALUE|PROPPARAM_LENGTH|PROPPARAM_PERCENTAGE)

// HTML ATRIBUTES
#define PP_UV_LENGTH                (PROPPARAM_UNITVALUE)
#define PP_UV_LENGTH_OR_PERCENT     (PP_UV_LENGTH|PROPPARAM_PERCENTAGE)

// CSS values
#define PP_UV_SCALER                (PROPPARAM_LENGTH|PROPPARAM_UNITVALUE)
#define PP_UV_SCALER_OR_PERCENT     (PP_UV_SCALER|PROPPARAM_PERCENTAGE)

// Our property types are the same as the first 32 VT_ VARENUM values
// The private types start at 32
enum PROPTYPE_FLAGS
{
    PROPTYPE_EMPTY          = 0,
    PROPTYPE_NULL           = 1,
    PROPTYPE_I2             = 2,
    PROPTYPE_I4             = 3,
    PROPTYPE_R4             = 4,
    PROPTYPE_R8             = 5,
    PROPTYPE_CY             = 6,
    PROPTYPE_DATE           = 7,
    PROPTYPE_BSTR           = 8,
    PROPTYPE_DISPATCH       = 9,
    PROPTYPE_ERROR          = 10,
    PROPTYPE_BOOL           = 11,
    PROPTYPE_VARIANT        = 12,
    PROPTYPE_UNKNOWN        = 13,
    PROPTYPE_DECIMAL        = 14,
    PROPTYPE_I1             = 16,
    PROPTYPE_UI1            = 17,
    PROPTYPE_UI2            = 18,
    PROPTYPE_UI4            = 19,
    PROPTYPE_I8             = 20,
    PROPTYPE_UI8            = 21,
    PROPTYPE_INT            = 22,
    PROPTYPE_UINT           = 23,
    PROPTYPE_VOID           = 24,
    PROPTYPE_HRESULT        = 25,
    PROPTYPE_PTR            = 26,
    PROPTYPE_SAFEARRAY      = 27,
    PROPTYPE_CARRAY         = 28,
    PROPTYPE_USERDEFINED    = 29,
    PROPTYPE_LPSTR          = 30,
    PROPTYPE_LPWSTR         = 31,

    // our private types
    PROPTYPE_HTMLNUM        = 32,   // Private type -- HTML number: %50, etc
    PROPTYPE_ITEM_ARY       = 33,   // Private type -- array of ItemList's
    PROPTYPE_BSTR_ARY       = 34,   // Private type -- array of BSTRs
    PROPTYPE_CSTRING        = 35,   // Private type -- CStr
    PROPTYPE_CSTR_ARY       = 36,   // Private type -- array of CStr's

    PROPTYPE_LAST           = PROPTYPE_LPWSTR,
    PROPTYPE_MAX            = 0xFFFFFFFF
};

// these are helper macros for HandleProperty methods
#define PROPTYPE(dw)        (dw>>16)
#define SETPROPTYPE(dw,t)   dw = (dw&0x0000FFFF) | (t<<16)
#define OPCODE(dw)          (dw&HANDLEPROP_OPMASK)
#define OPERATION(dw)       (dw&HANDLEPROP_OPERATION)
#define SETOPCODE(dw,o)     dw = (dw&~HANDLEPROP_OPMASK) | o
#define ISSET(dw)           (dw&HANDLEPROP_SET)
#define ISSTREAM(dw)        (dw&HANDLEPROP_STREAM)
#define ISSAMPLING(dw)      (dw&HANDLEPROP_SAMPLE)
#define SETSAMPLING(dw)     (dw|=HANDLEPROP_SAMPLE) 
// These flags are used to indicate the operation in the HandleProperty methods
enum HANDLEPROP_FLAGS
{
    HANDLEPROP_DEFAULT      = 0x001,    // property value to default
    HANDLEPROP_VALUE        = 0x002,    // property value
    HANDLEPROP_AUTOMATION   = 0x004,    // property value with validation and extended error handling
    HANDLEPROP_COMPARE      = 0x008,    // property value compare
    HANDLEPROP_STREAM       = 0x010,    // use stream
    HANDLEPROP_OPMASK       = 0x01F,    // mask for operation code above

    HANDLEPROP_SET          = 0x020,    // set

    HANDLEPROP_DONTVALIDATE = 0x040,    // IMG width&height - don't validate when setting
    HANDLEPROP_SAMPLE       = 0x080,    // sniffing for a good parse or not
    HANDLEPROP_IMPORTANT    = 0x100,    // stylesheets: this property is "!important"
    HANDLEPROP_IMPLIED      = 0x200,    // stylesheets: this property has been implied by another.

    HANDLEPROP_MERGE        = 0x400,    // property value must case an onPropertyChange

    HANDLEPROP_OPERATION    = 0xFFFF,   // Opcode and extra bits
    HANDLEPROP_MAX          = 0xFFFFFFFF // another enum we need to make max for win16.
};

#define HANDLEPROP_SETHTML      (HANDLEPROP_SET|HANDLEPROP_VALUE|(PROPTYPE_LPWSTR<<16))
#define HANDLEPROP_MERGEHTML    (HANDLEPROP_SET|HANDLEPROP_VALUE|HANDLEPROP_MERGE|(PROPTYPE_LPWSTR<<16))
#define HANDLEPROP_GETHTML      (HANDLEPROP_STREAM|(PROPTYPE_LPWSTR<<16))

typedef HRESULT (STDMETHODCALLTYPE PROPERTYDESC::*PFN_HANDLEPROPERTY)(DWORD dwOpcode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
#define PROPERTYDESC_METHOD(fn) \
    HRESULT STDMETHODCALLTYPE Handle ## fn ## Property(DWORD, void*, CBase*, CVoid*) const;

struct PROPERTYDESC
{
    // These property handlers are the combination of the old Set/Get property helpers, the
    // persistance helpers and the new HTML string parsing helpers.
    //
    // The first DWORD is encoding the incoming type (PROPTYPE_FLAG) in the upper WORD and
    // the opcode in the lower WORD (HANDLERPROP_FLAG)
    //
    // The pValue is pointing to the 'media' the value is stored for the get and set
    PROPERTYDESC_METHOD(Num);               // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(String);            // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(Enum);              // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(Color);             // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(Variant);           // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(UnitValue);         // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(Style);             // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(StyleComponent);    // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(Url);               // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;
    PROPERTYDESC_METHOD(Code);              // (DWORD dwOpCode, void* pValue, CBase* pObject, CVoid* pSubObject) const;

    inline HRESULT Default(CBase* pObject) const
    {
        return (pfnHandleProperty ? 
            CALL_METHOD(this, pfnHandleProperty,
            (HANDLEPROP_SET|HANDLEPROP_DEFAULT, 
            NULL,
            pObject,
            CVOID_CAST(pObject))) : S_OK);
    }

    ENUMDESC* GetEnumDescriptor(void) const
    {
        ENUMDESC* pEnumDesc;
        NUMPROPPARAMS* ppp = (NUMPROPPARAMS*)(this+1);

        if(!(GetPPFlags() & PROPPARAM_ENUM))
        {
            return NULL;
        }

        if(GetPPFlags() & PROPPARAM_ANUMBER)
        {
            // enumref ptr is after the sizeof() member
            pEnumDesc = *(ENUMDESC**)((BYTE*)(ppp+1) + sizeof(DWORD_PTR));
        }
        else
        {
            // enumref is coded in the lMax
            pEnumDesc = (ENUMDESC*)ppp->lMax;
        }
        return pEnumDesc;
    }
    const BASICPROPPARAMS* GetBasicPropParams() const
    {
        return (const BASICPROPPARAMS*)(this+1);
    }
    const NUMPROPPARAMS* GetNumPropParams() const
    {
        return (const NUMPROPPARAMS*)(this+1);
    }

    const DWORD GetPPFlags() const
    {
        return GetBasicPropParams()->dwPPFlags;
    }
    const DWORD GetdwFlags() const
    {
        return GetBasicPropParams()->dwFlags;
    }
    const DISPID GetDispid() const
    {
        return GetBasicPropParams()->dispid;
    }
    HRESULT TryLoadFromString(DWORD dwOpCode, LPCTSTR szString, CBase* pBaseObject, CAttrArray** ppAA) const
    {
        return (CALL_METHOD(this, pfnHandleProperty,
            (dwOpCode|(PROPTYPE_LPWSTR<<16|HANDLEPROP_SAMPLE), 
            const_cast<TCHAR*>(szString), pBaseObject, 
            (CVoid*)(void*)ppAA)));
    }
    HRESULT CallHandler(CBase* pBaseObject, DWORD dwFlags, void* pvArg) const
    {
        BASICPROPPARAMS* pbpp = (BASICPROPPARAMS*)(this + 1);
        return (CALL_METHOD(this, pfnHandleProperty,
            (dwFlags, pvArg, pBaseObject, 
            pbpp->dwPPFlags&PROPPARAM_ATTRARRAY?
            (CVoid*)(void*)pBaseObject->GetAttrArray():(CVoid*)(void*)pBaseObject)));
    }
    HRESULT HandleLoadFromHTMLString(CBase* pBaseObject, LPTSTR szString) const
    {
        return  CallHandler(pBaseObject, HANDLEPROP_SETHTML, (void*)szString); 
    }
    HRESULT HandleMergeFromHTMLString(CBase* pBaseObject, LPTSTR szString) const
    {
        return  CallHandler(pBaseObject, HANDLEPROP_MERGEHTML, (void*)szString); 
    }
    HRESULT HandleSaveToHTMLStream(CBase* pBaseObject, void* pStream) const
    {
        return  CallHandler(pBaseObject, HANDLEPROP_GETHTML, (void*)pStream); 
    }
    HRESULT HandleCompare(CBase* pBaseObject, void* pWith) const
    {
        return  CallHandler(pBaseObject, HANDLEPROP_COMPARE, (void*)pWith); 
    }
    // Gets property into a VARIANT in most compact representation
    HRESULT HandleGetIntoVARIANT(CBase* pBaseObject, VARIANT* pVariant) const
    {
        return  CallHandler(pBaseObject,
            (HANDLEPROP_VALUE|(PROPTYPE_VARIANT<<16)), (void*)pVariant); 
    }
    // Get property into VARIANT compatable with readinf property thru automation
    HRESULT HandleGetIntoAutomationVARIANT(CBase* pBaseObject, VARIANT* pVariant) const
    {
        return  CallHandler(pBaseObject,
            (HANDLEPROP_AUTOMATION|(PROPTYPE_VARIANT<<16)), (void*)pVariant); 
    }
    HRESULT HandleGetIntoBSTR(CBase* pBaseObject, BSTR* pBSTR) const
    {
        return  CallHandler(pBaseObject,
            (HANDLEPROP_VALUE|(PROPTYPE_BSTR<<16)), (void*)pBSTR); 
    }

    BOOL IsBOOLProperty(void) const;
    BOOL IsNotPresentAsDefaultSet(void) const 
    {
        return GetPPFlags()&PROPPARAM_NOTPRESENTASDEFAULT;
    }
    BOOL UseNotAssignedValue(void) const 
    {
        return !(GetPPFlags()&PROPPARAM_DONTUSENOTASSIGNED);
    }
    BOOL IsInvalidAsNoAssignSet(void) const 
    {
        return GetPPFlags()&PROPPARAM_INVALIDASNOASSIGN;
    }
    BOOL IsStyleSheetProperty(void) const 
    {
        return GetPPFlags()&PROPPARAM_STYLESHEET_PROPERTY;
    }

    ULONG GetNotPresentDefault(void) const
    {
        if(IsNotPresentAsDefaultSet() || !UseNotAssignedValue())
        {
            return ulTagNotPresentDefault;
        }
        else
        {
            return ulTagNotAssignedDefault;
        }
    }
    ULONG GetInvalidDefault(void) const
    {
        if(IsInvalidAsNoAssignSet() && UseNotAssignedValue())
        {
            return ulTagNotAssignedDefault;
        }
        else
        {
            return ulTagNotPresentDefault;
        }
    }
    PFN_HANDLEPROPERTY pfnHandleProperty;   // Pointer to the handler function
    const TCHAR*    pstrName;               // Name of property
    const TCHAR*    pstrExposedName;        // Exposed name of the property
    // We have two possible "defaults"
    // The first one is applied if the Tag was just plain missing
    ULONG_PTR       ulTagNotPresentDefault;
    // The second is applied if the Tag is present, but not set equal to anything
    // e.g. "SOMETAG=" OR "SOMETAG"
    ULONG_PTR       ulTagNotAssignedDefault;
};



typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_NUMPROPGET)(long* pnVal);
typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_NUMPROPSET)(long nVal);

typedef HRESULT (STDMETHODCALLTYPE CBase::*PFNB_NUMPROPGET)(long* pnVal);
typedef HRESULT (STDMETHODCALLTYPE CBase::*PFNB_NUMPROPSET)(long nVal);

typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_BSTRPROPGET)(BSTR* pbstr);
typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_BSTRPROPSET)(BSTR bstr);

typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_CSTRPROPGET)(CString* pcstr);
typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_CSTRPROPSET)(CString* pcstr);

typedef HRESULT (STDMETHODCALLTYPE CBase::*PFNB_BSTRPROPGET)(BSTR* pbstr);
typedef HRESULT (STDMETHODCALLTYPE CBase::*PFNB_BSTRPROPSET)(BSTR bstr);

typedef HRESULT (STDMETHODCALLTYPE CBase::*PFNB_CSTRPROPGET)(CString* pcstr);
typedef HRESULT (STDMETHODCALLTYPE CBase::*PFNB_CSTRPROPSET)(CString* pcstr);

typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_VARIANTPROP)(VARIANT* pbstr);
typedef HRESULT (STDMETHODCALLTYPE CBase::*PFNB_VARIANTPROP)(VARIANT* pbstr);

typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_ARBITRARYPROPGET)(...);
typedef HRESULT (STDMETHODCALLTYPE CVoid::*PFN_ARBITRARYPROPSET)(...);

struct DEFAULTARGDESC
{
    DWORD_PTR defargs[];
};

//+------------------------------------------------------------------------
//
//  Possible PROPERTY_DESC generated by pdlparser:
//
//      PROPERTYDESC_BASIC_ABSTRACT     (only if abstract set on class)
//      PROPERTYDESC_BASIC
//      PROPERTYDESC_NUMPROP_ABSTRACT   (only if abstract set on class)
//      PROPERTYDESC_NUMPROP
//      PROPERTYDESC_NUMPROP_GETSET
//      PROPERTYDESC_NUMPROP_ENUMREF
//
//+------------------------------------------------------------------------
struct PROPERTYDESC_METHOD
{
    PROPERTYDESC            a;
    BASICPROPPARAMS         b;
    const DEFAULTARGDESC*   c;      // pointer to default value arguments
    WORD                    d;      // total argument count
    WORD                    e;      // required count
};

struct PROPERTYDESC_BASIC_ABSTRACT
{
    PROPERTYDESC        a;
    BASICPROPPARAMS     b;
};

struct PROPERTYDESC_BASIC
{
    PROPERTYDESC        a;
    BASICPROPPARAMS     b;
    DWORD_PTR           c;
};

struct PROPERTYDESC_NUMPROP_ABSTRACT
{
    PROPERTYDESC        a;
    NUMPROPPARAMS       b;
};

struct PROPERTYDESC_NUMPROP
{
    PROPERTYDESC        a;
    NUMPROPPARAMS       b;
    DWORD_PTR           c;
};

struct PROPERTYDESC_NUMPROP_GETSET
{
    PROPERTYDESC        a;
    NUMPROPPARAMS       b;
    PFN_NUMPROPGET      c;
    PFN_NUMPROPSET      d;
};

struct PROPERTYDESC_NUMPROP_ENUMREF
{
    PROPERTYDESC        a;
    NUMPROPPARAMS       b;
    DWORD_PTR           c;
    const ENUMDESC*     pE;
};

struct PROPERTYDESC_CSTR_GETSET
{ 
    PROPERTYDESC a; 
    BASICPROPPARAMS b; 
    PFN_CSTRPROPGET c; 
    PFN_CSTRPROPSET d; 
};


inline const PROPERTYDESC* BASICPROPPARAMS::GetPropertyDesc(void) const
{ 
    return (const PROPERTYDESC*)this-1;
}

inline const PROPERTYDESC* NUMPROPPARAMS::GetPropertyDesc(void) const
{ 
    return (const PROPERTYDESC*)this-1;
}

//+------------------------------------------------------------------------
//
//  Helpers to set/get integer of given size or type at given address
//
//-------------------------------------------------------------------------
void SetNumberOfSize(void* pv, int cb, long i);
long GetNumberOfSize(void* pv, int cb);

void SetNumberOfType(void* pv, VARENUM vt, long i);
long GetNumberOfType(void* pv, VARENUM vt);

HRESULT LookupEnumString(const NUMPROPPARAMS* ppp, LPCTSTR pstr, long* plNewValue);

// Function used by shdocvw to do a case sensitive compare of a typelibrary.
STDAPI MatchExactGetIDsOfNames(ITypeInfo* pTI, REFIID riid, LPOLESTR* rgszNames,
                               UINT cNames, LCID lcid, DISPID* rgdispid, BOOL fCaseSensitive);

#endif //__XINDOWS_CORE_BASE_BASE_H__